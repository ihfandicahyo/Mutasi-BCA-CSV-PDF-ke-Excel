import pandas as pd
import os
import sys
import warnings
from difflib import SequenceMatcher

warnings.simplefilter(action='ignore', category=FutureWarning)

def text_similarity(a, b):
    return SequenceMatcher(None, str(a), str(b)).ratio()

def main():
    filename = 'Hasil_Gabungan_Mutasi_BCA.xlsx'
    output_file = 'Laporan_Transfer_Antar_Bank.xlsx'

    if not os.path.exists(filename):
        print(f"File {filename} tidak ditemukan.")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    try:
        xls = pd.read_excel(filename, sheet_name=None)
    except Exception as e:
        print(f"Gagal membaca file: {e}")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    all_data = []

    for sheet_name, df in xls.items():
        df.columns = df.columns.str.strip()
        
        required_columns = ['Tanggal Transaksi', 'Keterangan', 'Debit', 'Kredit']
        if not all(col in df.columns for col in required_columns):
            continue

        df = df.copy()
        df['Akun_Bank'] = str(sheet_name)
        
        df['Tanggal Transaksi'] = pd.to_datetime(df['Tanggal Transaksi'], dayfirst=True, errors='coerce')
        
        for col in ['Debit', 'Kredit']:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        df['Keterangan'] = df['Keterangan'].astype(str).str.strip().str.upper()
        df = df.dropna(subset=['Tanggal Transaksi'])
        
        if not df.empty:
            all_data.append(df)

    if not all_data:
        print("Tidak ada data yang valid untuk diproses.")
        input("Tekan Enter untuk keluar...")
        sys.exit()

    full_df = pd.concat(all_data, ignore_index=True)

    df_debit = full_df[full_df['Debit'] > 0].copy()
    df_kredit = full_df[full_df['Kredit'] > 0].copy()

    df_debit = df_debit.rename(columns={'Akun_Bank': 'Bank_Pengirim', 'Debit': 'Nominal', 'Keterangan': 'Ket_Pengirim'})
    df_kredit = df_kredit.rename(columns={'Akun_Bank': 'Bank_Penerima', 'Kredit': 'Nominal', 'Keterangan': 'Ket_Penerima'})

    cols_debit = ['Tanggal Transaksi', 'Bank_Pengirim', 'Ket_Pengirim', 'Nominal']
    cols_kredit = ['Tanggal Transaksi', 'Bank_Penerima', 'Ket_Penerima', 'Nominal']

    df_debit = df_debit[cols_debit]
    df_kredit = df_kredit[cols_kredit]

    potential_matches = pd.merge(
        df_debit, 
        df_kredit, 
        on=['Tanggal Transaksi', 'Nominal'], 
        how='inner'
    )

    potential_matches = potential_matches[potential_matches['Bank_Pengirim'] != potential_matches['Bank_Penerima']]

    valid_matches = []
    
    for index, row in potential_matches.iterrows():
        similarity = text_similarity(row['Ket_Pengirim'], row['Ket_Penerima'])
        
        if similarity >= 0.7:
            valid_matches.append(row)

    if not valid_matches:
        print("Tidak ditemukan transaksi transfer antar akun yang cocok.")
    else:
        hasil_final = pd.DataFrame(valid_matches)
        
        print(f"Ditemukan {len(hasil_final)} transaksi transfer antar akun terverifikasi:\n")
        
        preview = hasil_final.copy()
        preview['Tanggal Transaksi'] = preview['Tanggal Transaksi'].dt.strftime('%d/%m/%Y')
        
        for index, row in preview.iterrows():
            print(f"Tanggal    : {row['Tanggal Transaksi']}")
            print(f"Nominal    : {row['Nominal']:,.2f}")
            print(f"Dari Bank  : {row['Bank_Pengirim']}")
            print(f"Ke Bank    : {row['Bank_Penerima']}")
            print(f"Ket (Kirim): {row['Ket_Pengirim']}")
            print(f"Ket (Trm)  : {row['Ket_Penerima']}")
            print("-" * 60)

        try:
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                hasil_final.to_excel(writer, index=False, sheet_name='Laporan')
                worksheet = writer.sheets['Laporan']
                
                for i, col in enumerate(hasil_final.columns):
                    column_len = max(
                        hasil_final[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2
                    worksheet.set_column(i, i, column_len)
            
            print(f"\nLaporan berhasil disimpan: {output_file}")
            
        except Exception as e:
            print(f"Gagal menyimpan Excel: {e}")

    input("Tekan Enter untuk keluar...")

if __name__ == "__main__":
    main()