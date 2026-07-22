import streamlit as st
import pandas as pd

st.set_page_config(page_title="Rekap LTB Otomatis", layout="wide")
st.title("⚡ Web Rekap LTB Excel")

uploaded_file = st.file_uploader("Upload File LTB (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Baca Excel dari Sheet1
        df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
        st.success("✅ File berhasil dibaca! Memproses data...")
        
        # Bersihkan Data
        df['JENIS_TRANSAKSI'] = df['JENIS_TRANSAKSI'].astype(str).str.strip().str.upper()
        df['TARIF'] = df['TARIF'].astype(str).str.strip().str.upper()
        df['TARIF_LAMA'] = df['TARIF_LAMA'].astype(str).str.strip().str.upper()
        
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
        
        psg_box = bkr_box = 0 # Dummy untuk data yang belum ada

        # Looping untuk ngitung sesuai rule
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

        # Tampilkan Output
        st.subheader("📋 Hasil Rekap Akhir")
        
        output_text = f"""LTB Bulan Ini					
					
kWh Meter Elektronik (Pascabayar)					
1 Phasa	:	Pasang	{psg_paska_1ph}	Bongkar	{bkr_paska_1ph}
3 Phasa SL	:	Pasang	{psg_paska_3ph_sl}	Bongkar	{bkr_paska_3ph_sl}
3 Phasa STL	:	Pasang	{psg_paska_3ph_stl}	Bongkar	{bkr_paska_3ph_stl}
					
kWh Meter Elektronik Prabayar ( Pre Paid )					
1 Phasa	:	Pasang	{psg_lpb_1ph}	Bongkar	{bkr_lpb_1ph}
3 Phasa	:	Pasang	{psg_lpb_3ph}	Bongkar	{bkr_lpb_3ph}
					
MCB 1 Phasa	:	Pasang	{psg_mcb_1ph}	Bongkar	{bkr_mcb_1ph}
MCB 3 Phasa	:	Pasang	{psg_mcb_3ph}	Bongkar	{bkr_mcb_3ph}
BOX APP	:	Pasang	{psg_box}	Bongkar	{bkr_box}
					
SR 1 Phasa (10mm)					
Pasang	:	Pasang	{sr_1ph}	Bongkar	0
SR 3 Phasa (16mm)					
Pasang	:	Pasang	{sr_3ph}	Bongkar	0"""
        
        st.code(output_text, language="text")
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca data: {e}")
