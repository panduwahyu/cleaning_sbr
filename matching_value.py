import pandas as pd
import os

def match_excel_data():
    """
    Mencocokkan data dari dua file Excel berdasarkan kolom idsbr
    dan mengupdate data yang kosong dengan data dari file referensi
    """
    
    # Path file
    base_path = "D:/1. Poltstat STIS/New folder/aw/"
    source_file = os.path.join(base_path, "4506_baris_koordinat_kosong_kosong.xlsx")  # File dengan data lengkap
    target_file = os.path.join(base_path, "4506_baris_koordinat_kosong.xlsx")        # File yang akan diupdate
    output_file = os.path.join(base_path, "kosong_updated.xlsx")                     # File hasil
    
    try:
        # Membaca kedua file Excel
        print("Membaca file sumber...")
        df_source = pd.read_excel(source_file)
        
        print("Membaca file target...")
        df_target = pd.read_excel(target_file)
        
        print(f"Jumlah baris file sumber: {len(df_source)}")
        print(f"Jumlah baris file target: {len(df_target)}")
        
        # Kolom yang akan diupdate
        columns_to_update = [
            'alamat', 'Latitude', 'Longitude', 'idsls', 'iddesa', 
            'idkec', 'idkab', 'idprov', 'nmdesa', 'status', 'Kode Pos'
        ]
        
        # Validasi kolom yang diperlukan ada di kedua file
        if 'idsbr' not in df_source.columns:
            raise ValueError("Kolom 'idsbr' tidak ditemukan di file sumber")
        if 'idsbr' not in df_target.columns:
            raise ValueError("Kolom 'idsbr' tidak ditemukan di file target")
        
        # Cek kolom yang akan diupdate ada di file sumber
        missing_cols_source = [col for col in columns_to_update if col not in df_source.columns]
        if missing_cols_source:
            print(f"Peringatan: Kolom berikut tidak ditemukan di file sumber: {missing_cols_source}")
            columns_to_update = [col for col in columns_to_update if col in df_source.columns]
        
        # Pastikan kolom yang akan diupdate ada di file target
        for col in columns_to_update:
            if col not in df_target.columns:
                df_target[col] = None  # Tambahkan kolom kosong jika tidak ada
        
        # Buat copy dari df_target untuk hasil
        df_result = df_target.copy()
        
        # Set index berdasarkan idsbr untuk mempercepat lookup
        df_source_indexed = df_source.set_index('idsbr')
        
        # Counter untuk tracking update
        update_count = 0
        match_count = 0
        
        print("\nMemulai proses pencocokkan...")
        
        # Iterasi setiap baris di df_target
        for idx, row in df_target.iterrows():
            idsbr_target = row['idsbr']
            
            # Cek apakah idsbr ada di file sumber
            if idsbr_target in df_source_indexed.index:
                match_count += 1
                source_row = df_source_indexed.loc[idsbr_target]
                
                # Update kolom yang kosong atau NaN di target
                row_updated = False
                for col in columns_to_update:
                    # Cek jika nilai di target kosong/NaN dan nilai di source tidak kosong
                    if (pd.isna(df_result.at[idx, col]) or df_result.at[idx, col] == '' or df_result.at[idx, col] is None):
                        if not pd.isna(source_row[col]) and source_row[col] != '' and source_row[col] is not None:
                            df_result.at[idx, col] = source_row[col]
                            row_updated = True
                
                if row_updated:
                    update_count += 1
        
        # Simpan hasil ke file baru
        print(f"\nMenyimpan hasil ke {output_file}...")
        df_result.to_excel(output_file, index=False)
        
        # Tampilkan statistik
        print("\n" + "="*50)
        print("STATISTIK PENCOCOKKAN")
        print("="*50)
        print(f"Total baris di file target: {len(df_target)}")
        print(f"Jumlah idsbr yang cocok: {match_count}")
        print(f"Jumlah baris yang diupdate: {update_count}")
        print(f"Kolom yang diupdate: {', '.join(columns_to_update)}")
        print(f"File hasil disimpan di: {output_file}")
        
        return True
        
    except FileNotFoundError as e:
        print(f"Error: File tidak ditemukan - {e}")
        return False
    except Exception as e:
        print(f"Error: {e}")
        return False

# Jalankan fungsi
if __name__ == "__main__":
    print("Script Pencocokkan Data Excel")
    print("="*40)
    
    success = match_excel_data()
    
    if success:
        print("\nProses selesai dengan sukses!")
    else:
        print("\nProses gagal. Silakan periksa error di atas.")