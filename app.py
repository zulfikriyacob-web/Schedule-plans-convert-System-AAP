import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Production Schedule Converter", page_icon="⚙️", layout="wide")

st.title("⚙️ Auto-Converter: Production Schedule (DM & OH)")
st.write("Sila muat naik fail jadual asal di bawah. Sistem ini akan membersihkan data, menyusun *merged cells*, serta menarik maklumat OT secara automatik.")

uploaded_file = st.file_uploader("Upload File Excel (Format .xlsx)", type=['xlsx'])

def find_sheet(xls, keywords):
    """Cari nama sheet berdasarkan kata kunci (contoh: untuk support bulan Sept/Okt nanti)"""
    for sheet in xls.sheet_names:
        if all(k.upper() in sheet.upper() for k in keywords):
            return sheet
    return None

def process_wo_sheet(df):
    header_idx = 1
    headers = df.iloc[header_idx].values
    df_clean = df.iloc[header_idx + 1:].copy()
    df_clean.columns = headers
    df_clean = df_clean.dropna(how='all')
    
    col_mapping = {
        'PROD DATE': 'PROD_DATE',
        'LOT NUMBER': 'LOT_NUMBER',
        'MODEL': 'MODEL',
        'FS CODE': 'FS_CODE',
        'COLOUR PART': 'COLOUR_PART',
        'PART NAME': 'PART_NAME',
        'Total': 'PLANNED_COLOUR_QTY',
        'WO NUM': 'WO_NUM',
        'REMARKS': 'planner_remarks',
        'WO SUPPLY TO IMC': 'WO_SUPPLY_TO_IMC_DATE',
        'Closing Date': 'DI_DATE'
    }
    
    df_clean = df_clean.rename(columns=lambda x: col_mapping.get(x, x))
    
    # Fill down merged cells
    cols_to_ffill = ['PROD_DATE', 'LOT_NUMBER', 'MODEL', 'planner_remarks', 'WO_SUPPLY_TO_IMC_DATE', 'DI_DATE']
    for col in cols_to_ffill:
        if col in df_clean.columns:
            df_clean[col] = df_clean[col].ffill()
    
    date_columns = ['PROD_DATE', 'WO_SUPPLY_TO_IMC_DATE', 'DI_DATE']
    for col in date_columns:
        if col in df_clean.columns:
            df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce')
    
    if 'DI_DATE' in df_clean.columns and df_clean['DI_DATE'].isnull().all():
        df_clean['DI_DATE'] = df_clean['PROD_DATE'] + pd.Timedelta(days=1)
        
    for col in date_columns:
        if col in df_clean.columns:
            df_clean[col] = df_clean[col].dt.date
            
    if 'COLOUR_PART' in df_clean.columns:
        df_clean['COLOUR_PART'] = df_clean['COLOUR_PART'].astype(str).str.replace('B$', '', regex=True)
        
    df_clean['CUSTOMER_TYPE'] = 'OEM'
    
    def map_colour(c):
        c = str(c)
        if 'NH547' in c: return 'BB'
        elif 'B640M' in c: return 'CRB'
        elif 'NH731P' in c: return 'CB'
        elif 'NH883P' in c: return 'PW'
        elif 'NH830' in c: return 'LS'
        elif 'NH904M' in c: return 'RG'
        elif 'NH922P' in c: return 'SD'
        elif 'R575M' in c: return 'IR'
        return ''
    
    if 'COLOUR_PART' in df_clean.columns:
        df_clean['COLOUR_CODE'] = df_clean['COLOUR_PART'].apply(map_colour)
    else:
         df_clean['COLOUR_CODE'] = ''

    def get_side(name):
        name = str(name).upper()
        if 'RH-L' in name: return 'RH-L'
        elif 'RH-R' in name: return 'RH-R'
        elif 'L RR' in name or 'LEFT REAR' in name: return 'L RR'
        elif 'R RR' in name or 'RIGHT REAR' in name: return 'R RR'
        elif 'L FR' in name or 'LEFT FRONT' in name: return 'L FR'
        elif 'R FR' in name or 'RIGHT FRONT' in name: return 'R FR'
        return ''
    
    if 'PART_NAME' in df_clean.columns:
        df_clean['SIDE_CODE'] = df_clean['PART_NAME'].apply(get_side)
    else:
        df_clean['SIDE_CODE'] = ''
        
    if 'WO_NUM' in df_clean.columns and 'FS_CODE' in df_clean.columns:
        df_clean['WO_NUM_STR'] = pd.to_numeric(df_clean['WO_NUM'], errors='coerce').fillna(0).astype(int).astype(str)
        df_clean['WO_NUM_STR'] = df_clean['WO_NUM_STR'].replace('0', '')
        df_clean['WO_NUM_+_FS_CODE'] = df_clean['WO_NUM_STR'] + df_clean['FS_CODE'].astype(str).fillna('')
        df_clean = df_clean.drop(columns=['WO_NUM_STR'])
    
    df_clean['LINE'] = 'L1'
    
    def split_model(m):
        m = str(m)
        if '-' in m:
            parts = m.split('-')
            return parts[0], '-'.join(parts[1:])
        return m, ''
    
    if 'MODEL' in df_clean.columns:
        df_clean[['MODEL_TYPE', 'VARIANCE']] = df_clean['MODEL'].apply(lambda x: pd.Series(split_model(x)))
    else:
        df_clean['MODEL_TYPE'] = ''
        df_clean['VARIANCE'] = ''
        
    df_clean['CAMERA_APPLICABLE'] = np.nan
    df_clean['ADDITIONAL_ATTRIBUTE'] = np.nan
    df_clean['RUN_STATUS'] = np.nan
    
    target_cols = ['PROD_DATE', 'LOT_NUMBER', 'MODEL', 'FS_CODE', 'COLOUR_PART', 'PART_NAME', 
                   'PLANNED_COLOUR_QTY', 'WO_NUM', 'planner_remarks', 'WO_SUPPLY_TO_IMC_DATE', 
                   'DI_DATE', 'CUSTOMER_TYPE', 'COLOUR_CODE', 'SIDE_CODE', 'WO_NUM_+_FS_CODE', 
                   'LINE', 'MODEL_TYPE', 'VARIANCE', 'CAMERA_APPLICABLE', 'ADDITIONAL_ATTRIBUTE', 'RUN_STATUS']
    
    for col in target_cols:
        if col not in df_clean.columns:
            df_clean[col] = np.nan
            
    df_final = df_clean[target_cols]
    df_final = df_final.dropna(subset=['WO_NUM'])
    
    return df_final

def process_ot_sheet(df_pack, date_col_idx, ot_col_idx):
    if ot_col_idx >= len(df_pack.columns):
        return pd.DataFrame(columns=['PROD_DATE', 'OT'])
    
    df_ot = df_pack.iloc[:, [date_col_idx, ot_col_idx]].copy()
    df_ot.columns = ['PROD_DATE', 'OT']
    
    # Filter out empty dates
    df_ot['PROD_DATE_PARSED'] = pd.to_datetime(df_ot['PROD_DATE'], errors='coerce')
    df_ot = df_ot.dropna(subset=['PROD_DATE_PARSED'])
    df_ot['PROD_DATE'] = df_ot['PROD_DATE_PARSED'].dt.date
    df_ot = df_ot.drop(columns=['PROD_DATE_PARSED'])
    
    # Filter out empty OT rows
    df_ot = df_ot.dropna(subset=['OT'])
    return df_ot

if uploaded_file is not None:
    try:
        with st.spinner("Sistem sedang memproses data..."):
            xls = pd.ExcelFile(uploaded_file)
            
            # Detect Sheets Dynamically
            dm_sheet_name = find_sheet(xls, ['WO LISTING', 'DM'])
            oh_sheet_name = find_sheet(xls, ['WO LISTING', 'OH'])
            pack_sheet_name = find_sheet(xls, ['PACK'])
            
            if not dm_sheet_name or not oh_sheet_name or not pack_sheet_name:
                st.error("Ralat: Tidak menjumpai Sheet yang diperlukan. Pastikan ada nama sheet yang mengandungi 'WO LISTING DM', 'WO LISTING OH', dan 'PACK'.")
                st.stop()
            
            # 1. Process DM
            df_dm_raw = pd.read_excel(xls, sheet_name=dm_sheet_name, header=None)
            df_dm = process_wo_sheet(df_dm_raw)
            
            # 2. Process OH
            df_oh_raw = pd.read_excel(xls, sheet_name=oh_sheet_name, header=None)
            df_oh = process_wo_sheet(df_oh_raw)
            
            # 3. Process OT (DM & OH)
            df_pack = pd.read_excel(xls, sheet_name=pack_sheet_name, header=None)
            
            # DM OT: Col N (Index 13), Col P (Index 15)
            df_dm_ot = process_ot_sheet(df_pack, 13, 15)
            
            # OH OT: Col AI (Index 34), Col AK (Index 36)
            df_oh_ot = process_ot_sheet(df_pack, 34, 36)
            
            # 4. Generate Excel in memory with styling
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_dm.to_excel(writer, sheet_name='DM', index=False)
                df_oh.to_excel(writer, sheet_name='OH', index=False)
                df_dm_ot.to_excel(writer, sheet_name='DM_OT', index=False)
                df_oh_ot.to_excel(writer, sheet_name='OH_OT', index=False)
                
                workbook = writer.book
                header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
                alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                for sheet_name in workbook.sheetnames:
                    ws = workbook[sheet_name]
                    
                    for cell in ws[1]:
                        cell.fill = header_fill
                        cell.font = header_font
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                        for cell in row:
                            cell.border = border
                            cell.alignment = alignment
                            if isinstance(cell.value, pd.Timestamp) or type(cell.value).__name__ == 'date':
                                cell.number_format = 'YYYY-MM-DD'
                            
                    for column_cells in ws.columns:
                        length = max(len(str(cell.value)) for cell in column_cells)
                        ws.column_dimensions[column_cells[0].column_letter].width = min(length + 2, 30)

            processed_data = output.getvalue()
            
        st.success("Tadaaa! Selesai diproses.")
        
        st.download_button(
            label="📥 Download Database Excel Sekarang",
            data=processed_data,
            file_name="Converted_Database_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Maaf, ada ralat berlaku: {e}")
```eof
```text:requirements.txt
streamlit
pandas
openpyxl
```eof

Itu dia! Ada dua fail di atas. Sekarang macam mana nak buat Web App ni *live* di internet supaya staf bos boleh masuk ikut URL (macam `convert-jadual-kilang.streamlit.app`)? Senang sangat, ini langkahnya:

### Cara *Deploy* (Buat Live):

1. **Buat Akaun GitHub Percuma:**
   * Pergi ke [github.com](https://github.com/) dan buat satu akaun (jika bos tiada lagi).
   * Lepas login, klik tanda `+` di kanan atas dan pilih **"New repository"**.
   * Letak nama (contoh: `auto-converter`). Pastikan set kepada *Public* dan tekan **"Create repository"**.
   * Di muka depan *repository* tu, cari butang **"uploading an existing file"**. 
   * *Upload* dua fail yang saya buat di atas (`app.py` dan `requirements.txt`). Simpan (*Commit*).

2. **Sambung ke Streamlit (Server Percuma):**
   * Pergi ke laman web [share.streamlit.io](https://share.streamlit.io/).
   * Klik **"Continue with GitHub"** (ia akan *link* dengan GitHub bos tadi).
   * Klik butang besar **"New app"**.
   * Cari dan pilih repository `auto-converter` bos tadi.
   * Pastikan fail yang akan di-*run* adalah `app.py`.
   * Klik **"Deploy!"**

**Siap!** Dalam masa seminit dua, ia akan tunjuk laman web tu berfungsi sepenuhnya! Bos boleh *copy link* kat bar alamat di atas dan terus kongsi kat WhatsApp/email kepada *team* bos.

Nanti setiap bulan kalau staf *upload* apa-apa file pun yang di format sama, sistem automatik siapkan tab `DM`, `OH`, `DM_OT`, dan `OH_OT` sekelip mata. Boleh cuba dulu ikut step ni, kalau sangkut kat mana-mana bagitahu saya!
