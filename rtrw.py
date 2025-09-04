import pandas as pd
import re
import numpy as np

def extract_rt_rw(alamat):
    """
    Fungsi untuk mengekstrak RT dan RW dari string alamat dengan berbagai format
    
    Parameters:
    alamat (str): String alamat yang mengandung informasi RT dan RW
    
    Returns:
    str: String format "RT XXX RW XXX" atau NaN jika tidak ditemukan
    """
    if pd.isna(alamat):
        return np.nan
    
    # Konversi ke string dan uppercase untuk memudahkan matching
    alamat_str = str(alamat).upper()
    
    # Inisialisasi variabel RT dan RW
    rt_value = None
    rw_value = None
    
    # Pola regex untuk menangkap berbagai format RT
    rt_patterns = [
        r'RT[\s\.\:]*(\d+)',           # RT.7, RT:7, RT 7, RT7
        r'RT[\s\.\:]*0+(\d+)',          # RT.007, RT 007
        r'RT[\s\.\:]*(\d+)[\s]*/[\s]*RW', # RT.7/RW
        r'RT[\s\.\:]*(\d+)[\s]*/[\s]*(\d+)', # RT.7/8 (RT/RW format)
    ]
    
    # Pola regex untuk menangkap berbagai format RW
    rw_patterns = [
        r'RW[\s\.\:]*(\d+)',           # RW.8, RW:8, RW 8, RW8
        r'RW[\s\.\:]*0+(\d+)',          # RW.008, RW 008
        r'RT[\s\.\:]*\d+[\s]*/[\s]*RW[\s\.\:]*(\d+)', # RT.7/RW.8
        r'RT[\s\.\:]*\d+[\s]*/[\s]*(\d+)', # RT.7/8 (asumsi angka kedua adalah RW)
    ]
    
    # Cari RT
    for pattern in rt_patterns:
        match = re.search(pattern, alamat_str)
        if match:
            rt_value = match.group(1)
            break
    
    # Cari RW
    for pattern in rw_patterns:
        match = re.search(pattern, alamat_str)
        if match:
            # Untuk pola RT.7/8, group terakhir adalah RW
            if '/' in pattern and not 'RW' in pattern:
                if len(match.groups()) > 1:
                    rw_value = match.group(2)
                else:
                    rw_value = match.group(1)
            else:
                rw_value = match.group(1)
            break
    
    # Format hasil dengan padding 3 digit
    if rt_value and rw_value:
        rt_formatted = str(rt_value).zfill(3)
        rw_formatted = str(rw_value).zfill(3)
        return f"RT {rt_formatted} RW {rw_formatted}"
    elif rt_value:
        rt_formatted = str(rt_value).zfill(3)
        return f"RT {rt_formatted}"
    elif rw_value:
        rw_formatted = str(rw_value).zfill(3)
        return f"RW {rw_formatted}"
    else:
        return np.nan

def process_excel_file(input_file, output_file=None, alamat_column='alamat'):
    """
    Fungsi utama untuk memproses file Excel dan menambahkan kolom RT_RW
    
    Parameters:
    input_file (str): Nama file Excel input
    output_file (str): Nama file Excel output (opsional)
    alamat_column (str): Nama kolom yang berisi alamat (default: 'alamat')
    
    Returns:
    DataFrame: DataFrame dengan kolom RT_RW baru
    """
    try:
        # Baca file Excel
        print(f"Membaca file: {input_file}")
        df = pd.read_excel(input_file)
        
        # Cek apakah kolom alamat ada
        if alamat_column not in df.columns:
            print(f"Kolom '{alamat_column}' tidak ditemukan!")
            print(f"Kolom yang tersedia: {list(df.columns)}")
            
            # Cari kolom yang mungkin berisi kata 'alamat' (case insensitive)
            possible_columns = [col for col in df.columns if 'alamat' in col.lower()]
            if possible_columns:
                alamat_column = possible_columns[0]
                print(f"Menggunakan kolom: {alamat_column}")
            else:
                return None
        
        # Terapkan fungsi ekstraksi RT/RW
        print("Mengekstrak RT dan RW...")
        df['RT_RW'] = df[alamat_column].apply(extract_rt_rw)
        
        # Hitung statistik
        total_rows = len(df)
        extracted_rows = df['RT_RW'].notna().sum()
        success_rate = (extracted_rows / total_rows) * 100 if total_rows > 0 else 0
        
        print(f"\nStatistik Ekstraksi:")
        print(f"Total baris: {total_rows}")
        print(f"Berhasil diekstrak: {extracted_rows}")
        print(f"Tingkat keberhasilan: {success_rate:.2f}%")
        
        # Tampilkan beberapa contoh hasil
        print("\nContoh hasil ekstraksi:")
        sample_df = df[[alamat_column, 'RT_RW']].dropna()
        if len(sample_df) > 0:
            print(sample_df.head(10).to_string())
        
        # Simpan ke file baru jika output_file diberikan
        if output_file:
            df.to_excel(output_file, index=False)
            print(f"\nFile berhasil disimpan ke: {output_file}")
        else:
            # Jika tidak ada output_file, simpan dengan suffix '_with_RT_RW'
            output_file = input_file.replace('.xlsx', '_with_RT_RW.xlsx')
            df.to_excel(output_file, index=False)
            print(f"\nFile berhasil disimpan ke: {output_file}")
        
        return df
        
    except FileNotFoundError:
        print(f"Error: File '{input_file}' tidak ditemukan!")
        return None
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

# Fungsi untuk testing berbagai format
def test_extraction():
    """
    Fungsi untuk menguji ekstraksi dengan berbagai format RT/RW
    """
    test_cases = [
        "Jalan Boulevard No.13 RT.7/ RW.8",
        "Jl. Merdeka RT:7 RW:8",
        "Jl. Sudirman RT 7 RW 8",
        "Kompleks ABC RT.7/8",
        "RT 007 RW 008 Kelurahan XYZ",
        "Alamat lengkap RT007/RW008",
        "Gang Mawar rt.5 rw.6",
        "RT.003/RW.004",
        "Jl. Test RT 7/RW 8",
        "Alamat tanpa RT RW",
    ]
    
    print("Testing ekstraksi RT/RW:")
    print("-" * 50)
    for alamat in test_cases:
        result = extract_rt_rw(alamat)
        print(f"Input: {alamat}")
        print(f"Output: {result}")
        print("-" * 50)

# Main execution
if __name__ == "__main__":
    # Nama file input
    input_filename = "SUG_alamat_baru.xlsx"
    
    # Nama file output (opsional, bisa diubah sesuai kebutuhan)
    output_filename = "SUG_alamat_baru_with_RT_RW.xlsx"
    
    # Proses file Excel
    # Ganti 'alamat' dengan nama kolom yang sesuai di file Excel Anda
    result_df = process_excel_file(
        input_file=input_filename,
        output_file=output_filename,
        alamat_column='alamat'  # Sesuaikan dengan nama kolom di file Excel Anda
    )
    
    # Uncomment baris di bawah untuk menjalankan test
    # test_extraction()