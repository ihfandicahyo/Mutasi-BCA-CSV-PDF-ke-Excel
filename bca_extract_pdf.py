import pdfplumber
import pandas as pd
import re
import glob

def parse_decimal(value_str):
    if not value_str:
        return 0.0
    clean_str = value_str.upper().replace('DB', '').replace('CR', '').strip()
    clean_str = clean_str.replace(',', '')
    try:
        return float(clean_str)
    except ValueError:
        return 0.0

def extract_bca_clean(pdf_path, output_path):
    transactions = []
    x_mutasi_limit = None
    x_saldo_limit = None
    current_tx = None

    print(f"Memproses file: {pdf_path}")

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            words = page.extract_words()
            
            if x_mutasi_limit is None:
                header_mutasi = next((w for w in words if "MUTASI" in w['text'].upper()), None)
                header_saldo = next((w for w in words if "SALDO" in w['text'].upper()), None)
                
                if header_mutasi and header_saldo:
                    x_mutasi_limit = header_mutasi['x0'] - 20
                    x_saldo_limit = header_saldo['x0'] - 10
                elif page_num == 0:
                    width = page.width
                    x_mutasi_limit = width * 0.60
                    x_saldo_limit = width * 0.80

            lines = {}
            for w in words:
                y_coord = round(w['top'] / 3) * 3
                if y_coord not in lines:
                    lines[y_coord] = []
                lines[y_coord].append(w)

            sorted_y = sorted(lines.keys())

            for y in sorted_y:
                line_objs = lines[y]
                line_text = " ".join([w['text'] for w in line_objs])
                date_match = re.match(r'^(\d{2}/\d{2})', line_text)

                if date_match:
                    if current_tx:
                        transactions.append(current_tx)
                    
                    current_tx = {
                        "Tanggal": date_match.group(1) + "/2025",
                        "Keterangan": "",
                        "Debet": 0.0,
                        "Kredit": 0.0,
                        "Saldo": 0.0
                    }

                    desc_words = []
                    
                    for i, w in enumerate(line_objs):
                        text = w['text']
                        x = w['x0']

                        if i == 0 and text in current_tx["Tanggal"]:
                            continue

                        if x_saldo_limit and x >= x_saldo_limit:
                            if re.match(r'[\d,]+\.\d{2}', text):
                                current_tx["Saldo"] = parse_decimal(text)
                        
                        elif x_mutasi_limit and x_saldo_limit and x >= x_mutasi_limit and x < x_saldo_limit:
                            if re.search(r'[\d,]', text):
                                val = parse_decimal(text)
                                is_db = "DB" in text or "DB" in line_text
                                if is_db:
                                    current_tx["Debet"] = val
                                else:
                                    current_tx["Kredit"] = val
                        
                        elif x_mutasi_limit and x < x_mutasi_limit:
                            desc_words.append(text)
                    
                    current_tx["Keterangan"] = " ".join(desc_words)

                elif current_tx:
                    if "SALDO" in line_text and "MUTASI" in line_text: continue
                    if "HALAMAN" in line_text: continue
                    if "BERSAMBUNG" in line_text: continue
                    
                    add_desc = []
                    for w in line_objs:
                        if x_mutasi_limit and w['x0'] < x_mutasi_limit:
                            add_desc.append(w['text'])
                    
                    if add_desc:
                        current_tx["Keterangan"] += " " + " ".join(add_desc)

    if current_tx:
        transactions.append(current_tx)

    df = pd.DataFrame(transactions)
    
    if df.empty:
        print(f"Skipped {pdf_path}: No data found.")
        return

    df['Keterangan'] = df['Keterangan'].str.replace(r'\s+', ' ', regex=True).str.strip()
    df = df[['Tanggal', 'Keterangan', 'Debet', 'Kredit', 'Saldo']]

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Mutasi')
        workbook = writer.book
        worksheet = writer.sheets['Mutasi']
        fmt_num = workbook.add_format({'num_format': '#,##0.00'})
        fmt_text = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        
        worksheet.set_column('A:A', 12, fmt_text)
        worksheet.set_column('B:B', 60, fmt_text)
        worksheet.set_column('C:E', 18, fmt_num)

    print(f"Success: {output_path}")

def main():
    pdf_files = glob.glob("*.pdf")
    if not pdf_files:
        print("Tidak ada PDF ditemukan.")
        return

    print("1. Pindai Perdokumen")
    print("2. Proses Semua Dokumen")
    mode = input("Masukkan (1/2): ")

    if mode == '1':
        for i, f in enumerate(pdf_files):
            print(f"{i+1}. {f}")
        try:
            idx = int(input("Nomor Dokumen: ")) - 1
            if 0 <= idx < len(pdf_files):
                f = pdf_files[idx]
                extract_bca_clean(f, f.replace('.pdf', '_Excel.xlsx'))
            else:
                print("Nomor Salah.")
        except ValueError:
            print("Pilihan Salah.")
            
    elif mode == '2':
        for f in pdf_files:
            extract_bca_clean(f, f.replace('.pdf', '_Excel.xlsx'))

if __name__ == "__main__":
    main()
