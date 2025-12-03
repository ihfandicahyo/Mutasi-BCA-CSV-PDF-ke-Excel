import pandas as pd
import glob
import os

def clean_and_merge_final():
    print("--- PROGRAM PENGGABUNG DATA (sesuai tanggal 01 -> 31) ---")
    
    files = glob.glob("*.xlsx") + glob.glob("*.csv")
    output_filename = 'Hasil_Gabungan_Mutasi_BCA.xlsx'
    
    files = [f for f in files if output_filename not in f and not f.startswith('~$')]
    
    if not files:
        print("File tidak ditemukan.")
        return

    all_data = []

    for file in files:
        try:
            # 1. Deteksi Header
            if file.endswith('.csv'):
                try: temp_df = pd.read_csv(file, header=None, nrows=25, encoding='utf-8')
                except: temp_df = pd.read_csv(file, header=None, nrows=25, encoding='latin1')
            else:
                temp_df = pd.read_excel(file, header=None, nrows=25)
            
            header_idx = -1
            for idx, row in temp_df.iterrows():
                row_str = row.astype(str).str.cat(sep=' ')
                if 'Tanggal Transaksi' in row_str and 'Keterangan' in row_str:
                    header_idx = idx
                    break
            
            if header_idx == -1: continue

            # 2. Baca Data
            if file.endswith('.csv'):
                try: df = pd.read_csv(file, skiprows=header_idx, dtype=str, encoding='utf-8')
                except: df = pd.read_csv(file, skiprows=header_idx, dtype=str, encoding='latin1')
            else:
                df = pd.read_excel(file, skiprows=header_idx, dtype=str)
                
            df.columns = [str(c).strip() for c in df.columns]
            
            # 3. Filter Tanggal Valid
            date_pattern = r'\d{2}/\d{2}/\d{4}'
            if 'Tanggal Transaksi' in df.columns:
                df = df[df['Tanggal Transaksi'].astype(str).str.match(date_pattern, na=False)]
                
                cols = ['Tanggal Transaksi', 'Keterangan', 'Cabang', 'Jumlah', 'Saldo']
                df = df[cols].copy()
                
                # Konversi ke Datetime UNTUK PENGURUTAN
                df['Tanggal Transaksi'] = pd.to_datetime(df['Tanggal Transaksi'], dayfirst=True)
                
                # Bersihkan Angka
                def clean_money(val):
                    if not isinstance(val, str): return val
                    val = val.replace(',', '')
                    if 'CR' in val: return float(val.replace('CR', '').strip())
                    if 'DB' in val: return -float(val.replace('DB', '').strip())
                    try: return float(val)
                    except: return 0.0

                df['Jumlah'] = df['Jumlah'].apply(clean_money)
                df['Saldo'] = df['Saldo'].apply(clean_money)
                
                all_data.append(df)
                
        except Exception as e:
            print(f"Error: {e}")

    # 4. Gabung & Simpan
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        
        # --- PERBAIKAN DI SINI: URUTKAN ASCENDING (True) ---
        # True = Tanggal kecil (01) ke besar (31)
        # False = Tanggal besar (31) ke kecil (01)
        final_df = final_df.sort_values(by='Tanggal Transaksi', ascending=True)
        
        # Format jadi String agar jam 00:00:00 hilang
        final_df['Tanggal Transaksi'] = final_df['Tanggal Transaksi'].dt.strftime('%d/%m/%Y')
        
        writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
        final_df.to_excel(writer, index=False, sheet_name='Data')
        
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # Format Excel
        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
        text_fmt = workbook.add_format({'num_format': '@'}) # Format Text
        
        worksheet.set_column('A:A', 15, text_fmt)
        worksheet.set_column('D:E', 20, money_fmt)
        
        for i, col in enumerate(['Keterangan', 'Cabang']):
            idx = i + 1
            max_len = max(final_df[col].astype(str).map(len).max(), len(col)) + 3
            worksheet.set_column(idx, idx, max_len)
            
        writer.close()
        print(f"Selesai! File tersimpan: {output_filename}")
        
if __name__ == "__main__":
    clean_and_merge_final()