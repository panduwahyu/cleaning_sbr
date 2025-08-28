import pandas as pd
import numpy as np

def find_empty_coordinates(input_file, output_file, sheet_name='Sheet1'):
    """
    Mencari baris dengan nilai kosong pada kolom Latitude dan/atau Longitude
    
    Parameters:
    input_file (str): Nama file Excel input
    output_file (str): Nama file Excel output
    sheet_name (str): Nama sheet yang akan diproses
    """
    try:
        # Membaca file Excel
        print(f"Membaca file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        # Menampilkan info dasar tentang data
        print(f"Total baris dalam data: {len(df)}")
        print(f"Kolom yang tersedia: {list(df.columns)}")
        
        # Memeriksa apakah kolom Latitude dan Longitude ada
        if 'Latitude' not in df.columns or 'Longitude' not in df.columns:
            print("Error: Kolom 'Latitude' atau 'Longitude' tidak ditemukan!")
            print("Kolom yang tersedia:", list(df.columns))
            return
        
        # Mencari baris dengan nilai kosong pada Latitude atau Longitude
        # Nilai kosong bisa berupa NaN, None, atau string kosong
        empty_lat = df['Latitude'].isna() | (df['Latitude'] == '') | (df['Latitude'] == ' ')
        empty_lon = df['Longitude'].isna() | (df['Longitude'] == '') | (df['Longitude'] == ' ')
        
        # Baris yang memiliki nilai kosong pada Latitude ATAU Longitude
        empty_coordinates = empty_lat | empty_lon
        
        # Filter baris yang memiliki koordinat kosong
        empty_rows = df[empty_coordinates].copy()
        
        # Menampilkan hasil
        print(f"\nHasil pencarian:")
        print(f"Baris dengan Latitude kosong: {empty_lat.sum()}")
        print(f"Baris dengan Longitude kosong: {empty_lon.sum()}")
        print(f"Total baris dengan koordinat kosong: {len(empty_rows)}")
        
        if len(empty_rows) > 0:
            # Menyimpan hasil ke file Excel baru
            empty_rows.to_excel(output_file, index=False, sheet_name='Empty_Coordinates')
            print(f"\nData berhasil disimpan ke: {output_file}")
            
            # Menampilkan preview beberapa baris pertama
            print("\nPreview data yang ditemukan:")
            print(empty_rows[['Latitude', 'Longitude']].head(10))
            
        else:
            print("\nTidak ditemukan baris dengan koordinat kosong!")
            # Tetap buat file kosong sebagai indikator
            empty_df = pd.DataFrame(columns=df.columns)
            empty_df.to_excel(output_file, index=False, sheet_name='Empty_Coordinates')
            print(f"File kosong dibuat: {output_file}")
            
    except FileNotFoundError:
        print(f"Error: File '{input_file}' tidak ditemukan!")
    except Exception as e:
        print(f"Error: {str(e)}")

def find_empty_coordinates_detailed(input_file, output_file, sheet_name='Sheet1'):
    """
    Versi detail yang memisahkan berbagai jenis nilai kosong
    """
    try:
        print(f"Membaca file: {input_file}")
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        
        if 'Latitude' not in df.columns or 'Longitude' not in df.columns:
            print("Error: Kolom 'Latitude' atau 'Longitude' tidak ditemukan!")
            return
        
        # Membuat kondisi untuk berbagai jenis nilai kosong
        conditions = {
            'lat_nan': df['Latitude'].isna(),
            'lon_nan': df['Longitude'].isna(),
            'lat_empty_string': df['Latitude'] == '',
            'lon_empty_string': df['Longitude'] == '',
            'lat_whitespace': df['Latitude'].astype(str).str.strip() == '',
            'lon_whitespace': df['Longitude'].astype(str).str.strip() == ''
        }
        
        # Kombinasi semua kondisi kosong
        any_empty = pd.Series([False] * len(df))
        for condition in conditions.values():
            any_empty = any_empty | condition
        
        empty_rows = df[any_empty].copy()
        
        # Menambahkan kolom indikator untuk setiap jenis nilai kosong
        for name, condition in conditions.items():
            empty_rows[f'is_{name}'] = condition[any_empty]
        
        print(f"\nAnalisis detail:")
        for name, condition in conditions.items():
            print(f"{name}: {condition.sum()} baris")
        
        print(f"Total baris dengan masalah koordinat: {len(empty_rows)}")
        
        if len(empty_rows) > 0:
            empty_rows.to_excel(output_file, index=False, sheet_name='Detailed_Analysis')
            print(f"\nAnalisis detail disimpan ke: {output_file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")

# Penggunaan utama
if __name__ == "__main__":
    # Konfigurasi file
    input_filename = "D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong.xlsx"
    output_filename = "D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx"
    sheet_name = "Sheet1"
    
    print("=== PENCARI BARIS DENGAN KOORDINAT KOSONG ===\n")
    
    # Menjalankan fungsi pencarian standar
    find_empty_coordinates(input_filename, output_filename, sheet_name)
    
    print("\n" + "="*50)
    print("Untuk analisis detail, uncomment baris berikut:")
    print("# find_empty_coordinates_detailed(input_filename, 'analisis_detail.xlsx', sheet_name)")
    
    # Uncomment baris di bawah untuk analisis yang lebih detail
    # find_empty_coordinates_detailed(input_filename, 'analisis_detail.xlsx', sheet_name)