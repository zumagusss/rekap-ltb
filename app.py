import streamlit as st
import pandas as pd

# 1. Konfigurasi Halaman & Tema
st.set_page_config(page_title="Rekap LTB PLN", page_icon="⚡", layout="wide")

# 2. Injeksi CSS Khusus PLN
st.markdown("""
<style>
    /* Background keseluruhan */
    .stApp {
        background-color: #F8F9FA;
    }
    /* Warna teks judul utama (Biru PLN) */
    h1, h2, h3 {
        color: #00A2E9 !important; 
        font-family: 'Arial Black', sans-serif;
    }
    /* Styling area upload (Background biru muda, border Biru PLN) */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #00A2E9;
        background-color: #E6F6FE;
        border-radius: 10px;
        padding: 20px;
    }
    /* Styling Kotak Hasil Akhir (Border merah PLN) */
    .result-box {
        background-color: white; 
        padding: 20px; 
        border-left: 6px solid #ED1C24; 
        border-radius: 8px; 
        box-shadow: 0px 4px 12px rgba(0,0,0,0.1); 
        white-space: pre-wrap; 
        font-family: monospace; 
        color: #333333; 
        font-size: 15px;
        line-height: 1.6;
    }
</style>
""", unsafe_allow_html=True)

# 3. Header Custom (Logo + Judul sejajar)
col1, col2 = st.columns([1, 10])
with col1:
    try:
        st.image("logopln.png", width=80)
    except:
        st.error("Logo PLN tidak ditemukan.")
with col2:
    st.markdown("<h1>WEB REKAP LTB EXCEL</h1>", unsafe_allow_html=True)

st.markdown("---")

# 4. Area Upload
uploaded_file = st.file_uploader("Upload File LTB (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Baca Excel dari Sheet1
        df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
        st.success("✅ File berhasil dibaca! Memproses data...")
        
        # --- PERBAIKAN: Tangani nilai kosong (NaN) dengan .fillna('') ---
        df['JENIS_TRANSAKSI'] = df['JENIS_TRANSAKSI'].fillna('').astype(str).str.strip().str.upper()
        df['TARIF'] = df['TARIF'].fillna('').astype(str).str.strip().str.upper()
        df['TARIF_LAMA'] = df['TARIF_LAMA'].fillna('').astype(str).str.strip().str.upper()
        
        # Fungsi Filter Phasa
        def cek_phasa(daya):
            try:
                d = float(daya)
                if d <= 5500 or d in [7700, 11000]: return '1PH'
                elif 6600 <= d <= 630000: return '3PH'
            except:
                pass
            return 'UNKNOWN'

        df['PHASA'] = df['DAYA'].apply(cek_phasa)
        df['PHASA_LAMA'] = df['DAYA_LAMA'].apply(cek_phasa)
        
        # Fungsi Filter Tipe Tarif
        df['TIPE'] = df['TARIF'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')
        df['TIPE_LAMA'] = df['TARIF_LAMA'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')

        # Siapkan Variabel Hitungan
        psg_paska_1ph = bkr_paska_1ph = 0
        psg_paska_3ph_sl = bkr_paska_3ph_sl = 0
        psg_paska_3ph_stl = bkr_paska_3ph_stl = 0
        
        psg_lpb_1ph = bkr_lpb_1ph = 0
        psg_lpb_3ph = bkr_lpb_3ph = 0
        
        psg_mcb_1ph = bkr_mcb_1ph = 0
        psg_mcb_3ph = bkr_mcb_3ph = 0
        
        psg_box = bkr_box = 0

        # Looping Logika Perhitungan
        for _, row in df.iterrows():
            trx = row['JENIS_TRANSAKSI']
            phasa = row['PHASA']
            tipe = row['TIPE']
            phasa_lama = row['PHASA_LAMA']
            tipe_lama = row['TIPE_LAMA']
            
            if trx == 'PASANG BARU':
                if phasa == '1PH':
                    psg_mcb_1ph += 1
                    if tipe == 'PASKA': psg_paska_1ph += 1
                    else: psg_lpb_1ph += 1
                elif phasa == '3PH':
                    psg_mcb_3ph += 1
                    if tipe == 'PASKA': psg_paska_3ph_sl += 1
                    else: psg_lpb_3ph += 1
                    
            elif trx == 'PERUBAHAN DAYA':
                if phasa_lama == '1PH' and phasa == '1PH':
                    bkr_mcb_1ph += 1; psg_mcb_1ph += 1
                    if tipe_lama == 'PASKA' and tipe == 'LPB':
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA':
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                        
                elif phasa_lama == '1PH' and phasa == '3PH':
                    bkr_mcb_1ph += 1; psg_mcb_3ph += 1
                    if tipe_lama == 'LPB': bkr_lpb_1ph += 1
                    else: bkr_paska_1ph += 1
                    if tipe == 'LPB': psg_lpb_3ph += 1
                    else: psg_paska_3ph_sl += 1
                    
                elif phasa_lama == '3PH' and phasa == '1PH':
                    bkr_mcb_3ph += 1; psg_mcb_1ph += 1
                    if tipe_lama == 'LPB': bkr_lpb_3ph += 1
                    else: bkr_paska_3ph_sl += 1
                    if tipe == 'LPB': psg_lpb_1ph += 1
                    else: psg_paska_1ph += 1
                    
                elif phasa_lama == '3PH' and phasa == '3PH':
                    bkr_mcb_3ph += 1; psg_mcb_3ph += 1
                    if tipe_lama == 'LPB' and tipe == 'PASKA':
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1
                    elif tipe_lama == 'PASKA' and tipe == 'LPB':
                        bkr_paska_3ph_sl += 1; psg_lpb_3ph += 1
                        
            elif trx == 'MIGRASI':
                if phasa == '1PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB':
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA':
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                elif phasa == '3PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB':
                        bkr_paska_3ph_sl += 1; psg_lpb_3ph += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA':
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1

        # Hitung Kabel SR
        sr_1ph = psg_mcb_1ph * 20
        sr_3ph = psg_mcb_3ph * 30

        # Tampilkan Output Akhir
        st.markdown("### 📋 Hasil Rekap Akhir")
        
        output_text = f"""LTB Bulan Ini					
					
kWh Meter Elektronik (Pascabayar)					
1 Phasa\t:\tPasang\t{psg_paska_1ph}\tBongkar\t{bkr_paska_1ph}
3 Phasa SL\t:\tPasang\t{psg_paska_3ph_sl}\tBongkar\t{bkr_paska_3ph_sl}
3 Phasa STL\t:\tPasang\t{psg_paska_3ph_stl}\tBongkar\t{bkr_paska_3ph_stl}
					
kWh Meter Elektronik Prabayar ( Pre Paid )					
1 Phasa\t:\tPasang\t{psg_lpb_1ph}\tBongkar\t{bkr_lpb_1ph}
3 Phasa\t:\tPasang\t{psg_lpb_3ph}\tBongkar\t{bkr_lpb_3ph}
					
MCB 1 Phasa\t:\tPasang\t{psg_mcb_1ph}\tBongkar\t{bkr_mcb_1ph}
MCB 3 Phasa\t:\tPasang\t{psg_mcb_3ph}\tBongkar\t{bkr_mcb_3ph}
BOX APP\t:\tPasang\t{psg_box}\tBongkar\t{bkr_box}
					
SR 1 Phasa (10mm)					
Pasang\t:\tPasang\t{sr_1ph}\tBongkar\t0
SR 3 Phasa (16mm)					
Pasang\t:\tPasang\t{sr_3ph}\tBongkar\t0"""
        
        # Inject output ke dalam div HTML dengan CSS khusus
        st.markdown(f'<div class="result-box">{output_text}</div>', unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca data: {e}")
