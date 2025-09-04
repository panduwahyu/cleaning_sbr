import pandas as pd

def cocokkan_kode_pos(): 
    """
    Mencocokkan data kode pos berdasarkan nama kelurahan
    antara file SUG_kdps dan kodeposjkt (berdasarkan kemunculan kata)
    """
    try:
        import pandas as pd

        # Membaca file Excel
        print("Membaca file SUG_kdps.xlsx...")
        input_file_df = pd.read_excel('SUG_kdps.xlsx', sheet_name='Sheet1')
        
        print("Membaca file kodeposjkt.xlsx...")
        kodepos_df = pd.read_excel('kodeposjkt.xlsx', sheet_name='Sheet1')
        
        # Menampilkan info awal
        print(f"Jumlah data SUG_kdps: {len(input_file_df)}")
        print(f"Jumlah data kodepos: {len(kodepos_df)}")
        
        # Normalisasi kolom Kelurahan
        kodepos_df['Kelurahan_norm'] = kodepos_df['Kelurahan'].astype(str).str.upper().str.strip()
        
        matched_count = 0
        not_matched = []
        
        # Loop data SUG_kdps
        for index, row in input_file_df.iterrows():
            nmdesa = str(row['nmdesa']).upper().strip()
            nmdesa_words = set(nmdesa.split())  # pisahkan jadi kata
            
            found_kodepos = None
            
            # Cari di kodepos_df
            for _, kp_row in kodepos_df.iterrows():
                kelurahan_words = set(str(kp_row['Kelurahan_norm']).split())
                
                # Jika ada irisan kata
                if nmdesa_words & kelurahan_words:
                    found_kodepos = kp_row['kd_pos']
                    break
            
            if found_kodepos:
                input_file_df.at[index, 'kodepos'] = found_kodepos
                matched_count += 1
            else:
                not_matched.append(nmdesa)
        
        print(f"\nHasil pencocokan:")
        print(f"Data yang berhasil dicocokkan: {matched_count}")
        print(f"Data yang tidak ditemukan: {len(not_matched)}")
        
        if not_matched:
            print(f"\nContoh data yang tidak ditemukan (max 10):")
            for i, item in enumerate(not_matched[:10]):
                print(f"  {i+1}. {item}")
        
        # Simpan hasil ke file baru
        output_filename = 'SUG_kdps_updated.xlsx'
        
        with pd.ExcelFile('SUG_kdps.xlsx') as xls:
            all_sheets = xls.sheet_names
        
        with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
            for sheet in all_sheets:
                if sheet == 'Sheet1':
                    input_file_df.to_excel(writer, sheet_name=sheet, index=False)
                else:
                    original_sheet = pd.read_excel('SUG_kdps.xlsx', sheet_name=sheet)
                    original_sheet.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"\nFile berhasil disimpan sebagai: {output_filename}")
        print(f"\nSample data hasil update (5 baris pertama):")
        print(input_file_df[['nmdesa', 'kodepos']].head())
        
        return input_file_df
    
    except FileNotFoundError as e:
        print(f"Error: File tidak ditemukan - {e}")
    except KeyError as e:
        print(f"Error: Kolom tidak ditemukan - {e}")
        print("Pastikan kolom tersedia di Excel:")
        print("  - SUG_kdps.xlsx Sheet1: 'nmdesa' dan 'kodepos'")
        print("  - kodeposjkt.xlsx Sheet1: 'Kelurahan' dan 'kd_pos'")
    except Exception as e:
        print(f"Error tidak terduga: {e}")

def preview_data():
    """
    Fungsi untuk melihat preview data sebelum melakukan pencocokan
    """
    try:
        print("=== PREVIEW DATA ===")
        
        # Preview SUG_kdps
        input_file_df = pd.read_excel('SUG_kdps.xlsx', sheet_name='Sheet1')
        print(f"\nFile SUG_kdps.xlsx - Sheet Sheet1:")
        print(f"Jumlah baris: {len(input_file_df)}")
        print(f"Kolom yang tersedia: {list(input_file_df.columns)}")
        print("Sample data nmdesa:")
        if 'nmdesa' in input_file_df.columns:
            print(input_file_df['nmdesa'].head(10).tolist())
        
        # Preview kodepos
        kodepos_df = pd.read_excel('kodeposjkt.xlsx', sheet_name='Sheet1')
        print(f"\nFile kodeposjkt.xlsx - Sheet1:")
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