import pandas as pd

def cocokkan_kode_pos():
    """
    Mencocokkan data kode pos berdasarkan nama kelurahan
    antara file 4506_baris_koordinat_kosong_kosong dan D:/1. Poltstat STIS/New folder/aw/kodeposjkt
    """
    
    try:
        # Membaca file Excel
        print("Membaca file D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx...")
        Jaksel_jagakarsa2_df = pd.read_excel('D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx', sheet_name='Sheet1')
        
        print("Membaca file D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx...")
        kodepos_df = pd.read_excel('D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx', sheet_name='Sheet1')
        
        # Menampilkan info awal
        print(f"Jumlah data 4506_baris_koordinat_kosong_kosong: {len(Jaksel_jagakarsa2_df)}")
        print(f"Jumlah data kodepos: {len(kodepos_df)}")
        
        # Membuat dictionary untuk mapping Kelurahan -> kd_pos
        # Untuk menghindari duplikasi, kita ambil yang pertama jika ada duplikat
        kelurahan_to_kodepos = {}
        for _, row in kodepos_df.iterrows():
            kelurahan = str(row['Kelurahan']).strip().upper()  # Normalisasi text
            if kelurahan not in kelurahan_to_kodepos:
                kelurahan_to_kodepos[kelurahan] = row['kd_pos']
        
        print(f"Unique kelurahan di file kodepos: {len(kelurahan_to_kodepos)}")
        
        # Counter untuk tracking hasil
        matched_count = 0
        not_matched = []
        
        # Melakukan pencocokan dan update kode pos
        for index, row in Jaksel_jagakarsa2_df.iterrows():
            nmdesa = str(row['nmdesa']).strip().upper()  # Normalisasi text
            
            if nmdesa in kelurahan_to_kodepos:
                # Update kode pos jika ditemukan kecocokan
                Jaksel_jagakarsa2_df.at[index, 'kodepos'] = kelurahan_to_kodepos[nmdesa]
                matched_count += 1
            else:
                not_matched.append(nmdesa)
        
        print(f"\nHasil pencocokan:")
        print(f"Data yang berhasil dicocokkan: {matched_count}")
        print(f"Data yang tidak ditemukan: {len(not_matched)}")
        
        # Menampilkan beberapa data yang tidak ditemukan
        if not_matched:
            print(f"\nContoh data yang tidak ditemukan (max 10):")
            for i, item in enumerate(not_matched[:10]):
                print(f"  {i+1}. {item}")
        
        # Simpan hasil ke file baru
        output_filename = 'D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong_updated.xlsx'
        
        # Membaca semua sheet dari file asli untuk mempertahankan sheet lain
        with pd.ExcelFile('D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx') as xls:
            all_sheets = xls.sheet_names
        
        # Simpan dengan semua sheet, tapi update sheet Sheet1
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            for sheet in all_sheets:
                if sheet == 'Sheet1':
                    # Simpan sheet yang sudah diupdate
                    Jaksel_jagakarsa2_df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    # Simpan sheet lain tanpa perubahan
                    original_sheet = pd.read_excel('D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx', sheet_name=sheet)
                    original_sheet.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"\nFile berhasil disimpan sebagai: {output_filename}")
        
        # Tampilkan sample hasil
        print(f"\nSample data hasil update (5 baris pertama):")
        print(Jaksel_jagakarsa2_df[['nmdesa', 'kodepos']].head())
        
        return Jaksel_jagakarsa2_df
        
    except FileNotFoundError as e:
        print(f"Error: File tidak ditemukan - {e}")
        print("Pastikan file 'D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx' dan 'D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx' ada di direktori yang sama")
    except KeyError as e:
        print(f"Error: Kolom tidak ditemukan - {e}")
        print("Pastikan nama kolom sudah benar:")
        print("  - D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx sheet 'Sheet1' harus punya kolom 'nmdesa' dan 'kodepos'")
        print("  - D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx sheet 'Sheet1' harus punya kolom 'Kelurahan' dan 'kd_pos'")
    except Exception as e:
        print(f"Error tidak terduga: {e}")

def preview_data():
    """
    Fungsi untuk melihat preview data sebelum melakukan pencocokan
    """
    try:
        print("=== PREVIEW DATA ===")
        
        # Preview 4506_baris_koordinat_kosong_kosong
        Jaksel_jagakarsa2_df = pd.read_excel('D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx', sheet_name='Sheet1')
        print(f"\nFile D:/1. Poltstat STIS/New folder/aw/4506_baris_koordinat_kosong_kosong.xlsx - Sheet Sheet1:")
        print(f"Jumlah baris: {len(Jaksel_jagakarsa2_df)}")
        print(f"Kolom yang tersedia: {list(Jaksel_jagakarsa2_df.columns)}")
        print("Sample data nmdesa:")
        if 'nmdesa' in Jaksel_jagakarsa2_df.columns:
            print(Jaksel_jagakarsa2_df['nmdesa'].head(10).tolist())
        
        # Preview kodepos
        kodepos_df = pd.read_excel('D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx', sheet_name='Sheet1')
        print(f"\nFile D:/1. Poltstat STIS/New folder/aw/kodeposjkt.xlsx - Sheet1:")
        print(f"Jumlah baris: {len(kodepos_df)}")
        print(f"Kolom yang tersedia: {list(kodepos_df.columns)}")
        print("Sample data Kelurahan:")
        if 'Kelurahan' in kodepos_df.columns:
            print(kodepos_df['Kelurahan'].head(10).tolist())
            
    except Exception as e:
        print(f"Error saat preview: {e}")

# Jalankan fungsi
if __name__ == "__main__":
    print("=== SCRIPT PENCOCOKAN KODE POS ===")
    print("1. Preview data terlebih dahulu")
    preview_data()
    
    print("\n" + "="*50)
    print("2. Mulai proses pencocokan")
    result = cocokkan_kode_pos()