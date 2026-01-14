import pandas as pd
import os
import glob
from openpyxl.utils import get_column_letter
from datetime import datetime

def convert_csv_to_excel_autofit():
    current_folder = os.path.dirname(os.path.abspath(__file__))
    print(f"Mencari file CSV di folder: {current_folder}")

    csv_files = glob.glob(os.path.join(current_folder, "*.csv"))

    if not csv_files:
        print("Tidak ditemukan file CSV di folder ini.")
        input("Ketik enter untuk keluar")
        return

    print(f"Ditemukan {len(csv_files)} file CSV. Memulai proses...\n")

    month_map = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MEI", 6: "JUN",
        7: "JUL", 8: "AGU", 9: "SEP", 10: "OKT", 11: "NOV", 12: "DES"
    }

    for input_path in csv_files:
        try:
            print(f"-> Memproses: {os.path.basename(input_path)} ...")

            df = pd.read_csv(input_path, sep=None, engine='python', quotechar='"', header=None)

            def clean_text(text):
                if isinstance(text, str):
                    return text.strip().strip(',')
                return text

            df_clean = df.map(clean_text)

            rek_digits = "0000"
            try:
                raw_rek = str(df_clean.iloc[1, 0])
                digits_only = ''.join(filter(str.isdigit, raw_rek))
                if len(digits_only) >= 4:
                    rek_digits = digits_only[-4:]
                else:
                    rek_digits = digits_only
            except:
                pass

            valid_dates = []
            valid_months = []
            
            if len(df_clean) > 6:
                possible_dates = df_clean.iloc[6:, 0]
                for val in possible_dates:
                    try:
                        dt = datetime.strptime(str(val).strip(), "%d/%m/%Y")
                        valid_dates.append(dt.day)
                        valid_months.append(dt.month)
                    except:
                        continue

            dd_str = "00"
            mm_str = "UNK"

            if valid_dates:
                min_day = min(valid_dates)
                max_day = max(valid_dates)
                
                if min_day == max_day:
                    dd_str = str(min_day)
                else:
                    dd_str = f"{min_day} - {max_day}"
                
                if valid_months:
                    mm_str = month_map.get(valid_months[0], "UNK")

            output_filename = f"BCA {rek_digits} {dd_str} {mm_str}.xlsx"
            output_path = os.path.join(current_folder, output_filename)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_clean.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                
                worksheet = writer.sheets['Sheet1']

                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    
                    for cell in column:
                        try:
                            if cell.value:
                                cell_len = len(str(cell.value))
                                if cell_len > max_length:
                                    max_length = cell_len
                        except:
                            pass
                    
                    adjusted_width = (max_length + 2) 
                    worksheet.column_dimensions[column_letter].width = adjusted_width

            print(f"   [BERHASIL] Disimpan: {output_filename}")

        except Exception as e:
            print(f"   [GAGAL] Error pada file {os.path.basename(input_path)}: {e}")
        
        print("-" * 40)

    print("\nSemua proses selesai!")
    input("Ketik enter untuk keluar")

if __name__ == "__main__":
    convert_csv_to_excel_autofit()