import time
import webdriver_manager
import re
import pandas as pd
import os
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests
from bs4 import BeautifulSoup

# Setup logging
logging.basicConfig(filename='scraping_log.txt', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Cetak direktori kerja
print("Direktori kerja saat ini:", os.getcwd())

# Setup Selenium Options
def init_driver(headless=True):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-features=OptimizationHints,MediaFoundationVideoCapture")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# Fungsi untuk mencari email di halaman website
def find_email(website, worker_id, idsbr):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(website, headers=headers, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Cari email dengan regex
        email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        emails = re.findall(email_pattern, soup.get_text())
        
        # Filter email yang valid dan unik
        emails = list(set(emails))
        if emails:
            logging.info(f"[Worker {worker_id}] Email ditemukan untuk idsbr {idsbr}: {emails}")
            print(f"[Worker {worker_id}] Email ditemukan: {emails}")
            return ", ".join(emails)
        else:
            logging.info(f"[Worker {worker_id}] Tidak ada email ditemukan di website untuk idsbr {idsbr}")
            print(f"[Worker {worker_id}] Tidak ada email ditemukan di website")
            return None
    except Exception as e:
        logging.error(f"[Worker {worker_id}] Gagal mencari email untuk idsbr {idsbr}: {str(e)}")
        print(f"[Worker {worker_id}] Gagal mencari email: {str(e)}")
        return None

# Fungsi untuk scraping satu alamat
def scrape_address(query, idx, total, worker_id, idsbr, fallback_query=None):
    driver = init_driver(headless=False)  # Non-headless untuk debugging
    try:
        logging.info(f"[Worker {worker_id}] Memulai scraping untuk idsbr: {idsbr}, query: {query}")
        print(f"[Worker {worker_id}] Memulai scraping untuk idsbr: {idsbr}")
        time.sleep(2)  # Penundaan untuk mencegah rate limiting
        driver.get("https://www.google.com/maps")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchboxinput")))
        
        search_box = driver.find_element(By.ID, "searchboxinput")
        search_box.clear()
        search_box.send_keys(query)
        search_box.send_keys(Keys.ENTER)
        
        # Tunggu hingga halaman hasil pencarian termuat
        WebDriverWait(driver, 10).until(EC.url_contains("@"))
        
        # Klik pada hasil pencarian pertama untuk memastikan panel detail terbuka
        try:
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.hfpxzc")))
            result = driver.find_element(By.CSS_SELECTOR, "a.hfpxzc")
            result.click()
            logging.info(f"[Worker {worker_id}] Hasil pencarian diklik untuk idsbr {idsbr}")
            print(f"[Worker {worker_id}] Hasil pencarian diklik")
            time.sleep(3)  # Tunggu panel detail termuat
        except Exception as e:
            logging.warning(f"[Worker {worker_id}] Gagal mengklik hasil pencarian untuk idsbr {idsbr}: {str(e)}")
            print(f"[Worker {worker_id}] Gagal mengklik hasil pencarian: {str(e)}")

        # Ambil lat lon dari URL
        url = driver.current_url
        lat, lon = None, None
        if "@" in url:
            try:
                coords = url.split("@")[1].split(",")
                lat, lon = coords[0], coords[1]
            except Exception as e:
                logging.error(f"[Worker {worker_id}] Gagal mengambil koordinat untuk idsbr {idsbr}: {str(e)}")
                print(f"[Worker {worker_id}] Gagal mengambil koordinat: {str(e)}")

        # Ambil detail panel
        try:
            nama = driver.find_element(By.CLASS_NAME, "DUwDvf").text
        except:
            nama = None
            logging.warning(f"[Worker {worker_id}] Nama usaha tidak ditemukan untuk idsbr {idsbr}")

        # Jika nama tidak ditemukan dan ada fallback_query, coba ulang
        if nama is None and fallback_query:
            logging.info(f"[Worker {worker_id}] Nama usaha tidak ditemukan untuk idsbr {idsbr}, mencoba fallback query: {fallback_query}")
            print(f"[Worker {worker_id}] Nama usaha tidak ditemukan, mencoba fallback: {fallback_query}")
            driver.get("https://www.google.com/maps")
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "searchboxinput")))
            search_box = driver.find_element(By.ID, "searchboxinput")
            search_box.clear()
            search_box.send_keys(fallback_query)
            search_box.send_keys(Keys.ENTER)
            WebDriverWait(driver, 10).until(EC.url_contains("@"))
            
            # Klik hasil pencarian untuk fallback
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.hfpxzc")))
                result = driver.find_element(By.CSS_SELECTOR, "a.hfpxzc")
                result.click()
                logging.info(f"[Worker {worker_id}] Hasil pencarian fallback diklik untuk idsbr {idsbr}")
                print(f"[Worker {worker_id}] Hasil pencarian fallback diklik")
                time.sleep(3)
            except Exception as e:
                logging.warning(f"[Worker {worker_id}] Gagal mengklik hasil pencarian fallback untuk idsbr {idsbr}: {str(e)}")
                print(f"[Worker {worker_id}] Gagal mengklik hasil pencarian fallback: {str(e)}")
            
            url = driver.current_url
            if "@" in url:
                try:
                    coords = url.split("@")[1].split(",")
                    lat, lon = coords[0], coords[1]
                except Exception as e:
                    logging.error(f"[Worker {worker_id}] Gagal mengambil koordinat fallback untuk idsbr {idsbr}: {str(e)}")
                    print(f"[Worker {worker_id}] Gagal mengambil koordinat fallback: {str(e)}")

        # Ambil alamat
        try:
            alamat = driver.find_element(By.CLASS_NAME, "Io6YTe").text
        except:
            alamat = None
            logging.warning(f"[Worker {worker_id}] Alamat tidak ditemukan untuk idsbr {idsbr}")

        # Ambil nomor telepon
        try:
            telepon = driver.find_element(By.XPATH, "//button[contains(@aria-label,'Telepon') or contains(@aria-label,'Phone')]").text
        except:
            telepon = None
            logging.warning(f"[Worker {worker_id}] Nomor telepon tidak ditemukan untuk idsbr {idsbr}")

        # Ambil website dengan CSS selector yang menargetkan "Salin situs"
        website = None
        try:
            # Tunggu elemen website muncul (max 15 detik)
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[data-item-id*='authority'], a[data-tooltip*='Situs web'], a[data-tooltip*='Website']")))
            website_element = driver.find_element(By.CSS_SELECTOR, "a[data-item-id*='authority'], a[data-tooltip*='Situs web'], a[data-tooltip*='Website']")
            website = website_element.get_attribute("href")
            if website:
                logging.info(f"[Worker {worker_id}] Website berhasil diambil untuk idsbr {idsbr}: {website}")
                print(f"[Worker {worker_id}] Website berhasil diambil: {website}")
            else:
                logging.warning(f"[Worker {worker_id}] Website ditemukan tetapi href kosong untuk idsbr {idsbr}")
                print(f"[Worker {worker_id}] Website ditemukan tetapi href kosong")
        except:
            logging.warning(f"[Worker {worker_id}] Website tidak ditemukan untuk idsbr {idsbr}")
            print(f"[Worker {worker_id}] Website tidak ditemukan")

        # Cari email jika website tersedia
        email = None
        if website:
            email = find_email(website, worker_id, idsbr)

        # Ekstrak kode pos
        kode_pos = None
        if alamat:
            match = re.search(r"\b\d{5}\b", alamat)
            if match:
                kode_pos = match.group(0)

        hasil = {
            "idsbr": idsbr,
            "worker": worker_id,
            "nama": nama,
            "lat": lat,
            "lon": lon,
            "alamat_google": alamat,
            "kode_pos": kode_pos,
            "telepon": telepon,
            "website": website,
            "email": email,
            "query": query,
            "fallback_query": fallback_query if nama is None and fallback_query else None
        }

        print(f"[Worker {worker_id}] ({idx}/{total}) idsbr: {idsbr} | {nama}")
        print(f"    Lat: {lat}, Lon: {lon}")
        print(f"    Alamat: {alamat}")
        print(f"    Kode Pos: {kode_pos}")
        print(f"    Telepon: {telepon}")
        print(f"    Website: {website}")
        print(f"    Email: {email}")
        print(f"    Fallback Query: {fallback_query if nama is None and fallback_query else 'Tidak digunakan'}")
        print("-" * 60)
        logging.info(f"[Worker {worker_id}] Scraping berhasil untuk idsbr: {idsbr}")

        return hasil

    except Exception as e:
        logging.error(f"[Worker {worker_id}] Error pada idsbr {idsbr}: {str(e)}")
        print(f"[Worker {worker_id}] Error pada idsbr {idsbr}: {str(e)}")
        return {
            "idsbr": idsbr,
            "worker": worker_id,
            "nama": None,
            "lat": None,
            "lon": None,
            "alamat_google": None,
            "kode_pos": None,
            "telepon": None,
            "website": None,
            "email": None,
            "query": query,
            "fallback_query": fallback_query
        }
    finally:
        driver.quit()

# Baca data Excel
file_path = r"Scraping Usaha Gede.xlsx"
try:
    df = pd.read_excel(file_path)
except FileNotFoundError:
    print(f"File {file_path} tidak ditemukan. Pastikan path benar.")
    logging.error(f"File {file_path} tidak ditemukan.")
    exit()

# Tes dengan jumlah entri kecil untuk debugging
# df = df.head(5)  # Uncomment untuk tes 5 entri pertama

# Gabungkan alamat
df["alamat_lengkap"] = (
    df["nama_usaha"].astype(str) + ", " +
    df["alamat"].astype(str) + ", " +
    df["nmdesa"].astype(str) + ", " +
    df["nmkec"].astype(str) + ", " +
    df["nmkab"].astype(str) + ", " +
    df["nmprov"].astype(str)
)

# Buat fallback query tanpa nama usaha
df["fallback_alamat"] = (
    df["alamat"].astype(str) + ", " +
    df["nmdesa"].astype(str) + ", " +
    df["nmkec"].astype(str) + ", " +
    df["nmkab"].astype(str) + ", " +
    df["nmprov"].astype(str)
)

queries = df["alamat_lengkap"].tolist()
fallback_queries = df["fallback_alamat"].tolist()
idsbr_list = df["idsbr"].tolist()

# Jalankan scraping paralel
results = []
max_workers = 2  # Kurangi lebih jauh untuk debugging
with ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {}
    for i, (q, fq, idsbr) in enumerate(zip(queries, fallback_queries, idsbr_list)):
        worker_id = (i % max_workers) + 1
        print(f"[Main] Memulai worker {worker_id} untuk idsbr {idsbr}")
        future = executor.submit(scrape_address, q, i+1, len(queries), worker_id, idsbr, fq)
        futures[future] = q

    for future in as_completed(futures):
        try:
            data = future.result()
            results.append(data)
            # Simpan hasil parsial setiap 10 entri
            if len(results) % 10 == 0:
                df_hasil = pd.DataFrame(results)
                partial_output = r"hasilscraping_partial.xlsx"
                df_hasil.to_excel(partial_output, index=False)
                print(f"Hasil parsial disimpan di {os.path.abspath(partial_output)}")
                logging.info(f"Hasil parsial disimpan di {os.path.abspath(partial_output)}")
        except Exception as e:
            logging.error(f"Error saat mengambil hasil: {str(e)}")
            print(f"Error saat mengambil hasil: {str(e)}")

# Simpan hasil akhir ke Excel
output_path = r"hasilscraping.xlsx"
try:
    df_hasil = pd.DataFrame(results)
    df_hasil.to_excel(output_path, index=False)
    print(f"Selesai! Data disimpan di {os.path.abspath(output_path)}")
    logging.info(f"Data disimpan di {os.path.abspath(output_path)}")
except Exception as e:
    print(f"Gagal menyimpan file: {str(e)}")
    logging.error(f"Gagal menyimpan file: {str(e)}")