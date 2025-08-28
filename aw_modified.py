import pandas as pd

# Baca file excel
df = pd.read_excel("D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx", sheet_name="Sheet1")

# Fungsi untuk membersihkan Latitude
def fix_lat(value):
    if pd.isna(value):
        return value
    s = str(value).replace(".", "").replace(" ", "")
    negative = s.startswith("-")
    if negative:
        s = s[1:]
    # Sisipkan titik setelah digit pertama
    s = s[0] + "." + s[1:]
    if negative:
        s = "-" + s
    
    # Batasi maksimal 15 digit setelah koma
    if "." in s:
        integer_part, decimal_part = s.split(".")
        decimal_part = decimal_part[:15]
        s = integer_part + "." + decimal_part

    return s

# Fungsi untuk membersihkan Longitude
def fix_lon(value):
    if pd.isna(value):
        return value
    s = str(value).replace(".", "").replace(" ", "")
    negative = s.startswith("-")
    if negative:
        s = s[1:]
    # Sisipkan titik setelah digit ketiga
    s = s[:3] + "." + s[3:]
    if negative:
        s = "-" + s
    return s

# Terapkan fungsi
df["Latitude"] = df["Latitude"].apply(fix_lat)
df["Longitude"] = df["Longitude"].apply(fix_lon)

# Simpan hasil ke file baru
df.to_excel("D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx", index=False)

print("Proses selesai! File disimpan di: D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx")
