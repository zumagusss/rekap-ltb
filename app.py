import streamlit as st
import pandas as pd

# 1. Konfigurasi Halaman & Tema
st.set_page_config(page_title="Rekap LTB PLN", page_icon="⚡", layout="wide")

# 2. Injeksi CSS Khusus PLN & UI Responsif
st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    h1, h2, h3 { color: #00A2E9 !important; font-family: 'Arial Black', sans-serif; }
    
    /* Area Upload */
    div[data-testid="stFileUploader"] {
        border: 2px dashed #00A2E9; background-color: #EBF7FD; border-radius: 10px; padding: 20px;
    }
    
    /* Rekap Text Box */
    .result-box {
        background-color: white; padding: 20px; border-left: 6px solid #ED1C24; 
        border-radius: 8px; box-shadow: 0px 4px 12px rgba(0,0,0,0.1); 
        white-space: pre-wrap; font-family: 'Courier New', Courier, monospace; color: #333333; font-size: 15px; line-height: 1.6;
    }
    
    /* Styling Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 50px; white-space: pre-wrap; background-color: #ffffff; border-radius: 4px; 
        padding-top: 10px; padding-bottom: 10px; border: 1px solid #ddd; font-weight: bold; color: #555;
    }
    .stTabs [aria-selected="true"] { background-color: #00A2E9 !important; color: white !important; border-color: #00A2E9 !important; }
</style>
""", unsafe_allow_html=True)

# Fungsi HTML CSS Grid (Biar 100% Rapi di HP & PC)
def render_cards(data_dict, accent_color):
    html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin-bottom: 15px;">'
    for k, v in data_dict.items():
        html += f'''
        <div style="background: white; border-top: 4px solid {accent_color}; border-radius: 8px; padding: 15px; box-shadow: 0 2px 6px rgba(0,0,0,0.08); text-align: center; display: flex; flex-direction: column; justify-content: center;">
            <span style="font-size: 12px; color: #555; font-weight: 700; margin-bottom: 8px; line-height: 1.3;">{k}</span>
            <span style="font-size: 26px; color: #111; font-weight: 900;">{v}</span>
        </div>
        '''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

# 3. Header Custom (Logo + Judul)
col1, col2 = st.columns([1, 10])
with col1:
    try:
        st.image("logopln.png", width=80)
    except:
        pass
with col2:
    st.markdown("<h1>⚡ WEB REKAP LTB EXCEL</h1>", unsafe_allow_html=True)

st.markdown("---")

# 4. Area Upload
uploaded_file = st.file_uploader("Upload File LTB (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Pengecekan Ekstensi File
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df = pd.read_csv(uploaded_file)
        else:
            try:
                df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
            except ValueError:
                df = pd.read_excel(uploaded_file, sheet_name=0)
                
        st.success(f"✅ File {uploaded_file.name} berhasil dibaca!")
        
        # Bersihkan Data
        df['JENIS_TRANSAKSI'] = df['JENIS_TRANSAKSI'].fillna('').astype(str).str.strip().str.upper()
        df['TARIF'] = df['TARIF'].fillna('').astype(str).str.strip().str.upper()
        df['TARIF_LAMA'] = df['TARIF_LAMA'].fillna('').astype(str).str.strip().str.upper()
        
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
        df['TIPE'] = df['TARIF'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')
        df['TIPE_LAMA'] = df['TARIF_LAMA'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')

        # === GLOBAL VARIABLES ===
        psg_paska_1ph = bkr_paska_1ph = psg_paska_3ph_sl = bkr_paska_3ph_sl = psg_paska_3ph_stl = bkr_paska_3ph_stl = 0
        psg_lpb_1ph = bkr_lpb_1ph = psg_lpb_3ph = bkr_lpb_3ph = 0
        psg_mcb_1ph = bkr_mcb_1ph = psg_mcb_3ph = bkr_mcb_3ph = 0
        psg_box = bkr_box = 0

        # === DETAILED TRACKER ===
        dt_pb = {'PASANG KWH PASKA 3PH SL': 0, 'PASANG KWH LPB 3PH': 0, 'PASANG MCB 3PH': 0, 'PASANG KWH LPB 1PH': 0, 'PASANG KWH PASKA 1PH': 0, 'PASANG MCB 1PH': 0}
        dt_pd_1 = {'BONGKAR MCB 1PH': 0, 'PASANG MCB 1PH': 0}
        dt_pd_1_mig_lpb = {'BONGKAR MCB 1PH': 0, 'PASANG MCB 1PH': 0, 'BONGKAR KWH PASKA 1PH': 0, 'PASANG KWH LPB 1PH': 0}
        dt_pd_1_mig_paska = {'BONGKAR MCB 1PH': 0, 'PASANG MCB 1PH': 0, 'BONGKAR KWH LPB 1PH': 0, 'PASANG KWH PASKA 1PH': 0}
        dt_pd_1_to_3 = {'BONGKAR LPB 1PH': 0, 'BONGKAR PASKA 1PH': 0, 'PASANG LPB 3 PH': 0, 'PASANG PASKA 3 PH': 0, 'BONGKAR MCB 1PH': 0, 'PASANG MCB 3 PH': 0}
        dt_pd_3_to_1 = {'BONGKAR MCB 3PH': 0, 'BONGKAR LPB 3PH': 0, 'BONGKAR PASKA 3PH': 0, 'PASANG MCB 1PH': 0, 'PASANG LPB 1PH': 0, 'PASANG PASKA 1PH': 0}
        dt_pd_3 = {'BONGKAR MCB 3PH': 0, 'PASANG MCB 3PH': 0}
        dt_pd_3_mig_paska = {'BONGKAR MCB 3PH': 0, 'PASANG MCB 3PH': 0, 'BONGKAR KWH LPB 3PH': 0, 'PASANG KWH PASKA 3PH SL': 0}
        dt_mig = {'PASANG KWH LPB 3 FASA': 0, 'BONGKAR KWH PASKA 3 FASA SL': 0, 'PASANG KWH LPB 1PH': 0, 'BONGKAR KWH PASKA 1PH': 0, 'PASANG KWH PASKA 1PH': 0, 'BONGKAR LPB 1PH': 0, 'PASANG KWH PASKA 3PH': 0, 'BONGKAR LPB 3PH': 0}

        # Looping Logika
        for _, row in df.iterrows():
            trx = row['JENIS_TRANSAKSI']
            phasa = row['PHASA']
            tipe = row['TIPE']
            phasa_lama = row['PHASA_LAMA']
            tipe_lama = row['TIPE_LAMA']
            
            if trx == 'PASANG BARU':
                if phasa == '1PH':
                    psg_mcb_1ph += 1; dt_pb['PASANG MCB 1PH'] += 1
                    if tipe == 'PASKA': psg_paska_1ph += 1; dt_pb['PASANG KWH PASKA 1PH'] += 1
                    else: psg_lpb_1ph += 1; dt_pb['PASANG KWH LPB 1PH'] += 1
                elif phasa == '3PH':
                    psg_mcb_3ph += 1; dt_pb['PASANG MCB 3PH'] += 1
                    if tipe == 'PASKA': psg_paska_3ph_sl += 1; dt_pb['PASANG KWH PASKA 3PH SL'] += 1
                    else: psg_lpb_3ph += 1; dt_pb['PASANG KWH LPB 3PH'] += 1
                    
            elif trx == 'PERUBAHAN DAYA':
                if phasa_lama == '1PH' and phasa == '1PH':
                    bkr_mcb_1ph += 1; psg_mcb_1ph += 1
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                        dt_pd_1_mig_lpb['BONGKAR MCB 1PH'] += 1; dt_pd_1_mig_lpb['PASANG MCB 1PH'] += 1
                        dt_pd_1_mig_lpb['BONGKAR KWH PASKA 1PH'] += 1; dt_pd_1_mig_lpb['PASANG KWH LPB 1PH'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                        dt_pd_1_mig_paska['BONGKAR MCB 1PH'] += 1; dt_pd_1_mig_paska['PASANG MCB 1PH'] += 1
                        dt_pd_1_mig_paska['BONGKAR KWH LPB 1PH'] += 1; dt_pd_1_mig_paska['PASANG KWH PASKA 1PH'] += 1
                    else:
                        dt_pd_1['BONGKAR MCB 1PH'] += 1; dt_pd_1['PASANG MCB 1PH'] += 1
                        
                elif phasa_lama == '1PH' and phasa == '3PH':
                    bkr_mcb_1ph += 1; psg_mcb_3ph += 1
                    dt_pd_1_to_3['BONGKAR MCB 1PH'] += 1; dt_pd_1_to_3['PASANG MCB 3 PH'] += 1
                    if tipe_lama == 'LPB': bkr_lpb_1ph += 1; dt_pd_1_to_3['BONGKAR LPB 1PH'] += 1
                    else: bkr_paska_1ph += 1; dt_pd_1_to_3['BONGKAR PASKA 1PH'] += 1
                    if tipe == 'LPB': psg_lpb_3ph += 1; dt_pd_1_to_3['PASANG LPB 3 PH'] += 1
                    else: psg_paska_3ph_sl += 1; dt_pd_1_to_3['PASANG PASKA 3 PH'] += 1
                        
                elif phasa_lama == '3PH' and phasa == '1PH':
                    bkr_mcb_3ph += 1; psg_mcb_1ph += 1
                    dt_pd_3_to_1['BONGKAR MCB 3PH'] += 1; dt_pd_3_to_1['PASANG MCB 1PH'] += 1
                    if tipe_lama == 'LPB': bkr_lpb_3ph += 1; dt_pd_3_to_1['BONGKAR LPB 3PH'] += 1
                    else: bkr_paska_3ph_sl += 1; dt_pd_3_to_1['BONGKAR PASKA 3PH'] += 1
                    if tipe == 'LPB': psg_lpb_1ph += 1; dt_pd_3_to_1['PASANG LPB 1PH'] += 1
                    else: psg_paska_1ph += 1; dt_pd_3_to_1['PASANG PASKA 1PH'] += 1
                        
                elif phasa_lama == '3PH' and phasa == '3PH':
                    bkr_mcb_3ph += 1; psg_mcb_3ph += 1
                    if tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1
                        dt_pd_3_mig_paska['BONGKAR MCB 3PH'] += 1; dt_pd_3_mig_paska['PASANG MCB 3PH'] += 1
                        dt_pd_3_mig_paska['BONGKAR KWH LPB 3PH'] += 1; dt_pd_3_mig_paska['PASANG KWH PASKA 3PH SL'] += 1
                    else: 
                        dt_pd_3['BONGKAR MCB 3PH'] += 1; dt_pd_3['PASANG MCB 3PH'] += 1
                        
            elif trx == 'MIGRASI':
                if phasa == '1PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                        dt_mig['BONGKAR KWH PASKA 1PH'] += 1; dt_mig['PASANG KWH LPB 1PH'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                        dt_mig['BONGKAR LPB 1PH'] += 1; dt_mig['PASANG KWH PASKA 1PH'] += 1
                elif phasa == '3PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_3ph_sl += 1; psg_lpb_3ph += 1
                        dt_mig['BONGKAR KWH PASKA 3 FASA SL'] += 1; dt_mig['PASANG KWH LPB 3 FASA'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1
                        dt_mig['BONGKAR LPB 3PH'] += 1; dt_mig['PASANG KWH PASKA 3PH'] += 1

        sr_1ph = psg_mcb_1ph * 20
        sr_3ph = psg_mcb_3ph * 30
        
        dt_pb['PASANG SR 1PH'] = sr_1ph
        dt_pb['PASANG SR 3PH'] = sr_3ph
        dt_pb['PASANG KWH PASKA 3PH STL'] = 0
        dt_pb['PASANG BOX APP 53 KVA'] = 0

        # --- UI INTERAKTIF TABS ---
        tab1, tab2, tab3 = st.tabs(["📊 Dashboard Visual", "🔍 Detail Rincian", "📝 Rekap Akhir"])
        
        with tab1:
            st.markdown("### Ringkasan Cepat")
            render_cards({
                "TOTAL DATA": len(df),
                "MCB 1PH PASANG": psg_mcb_1ph,
                "LPB 1PH PASANG": psg_lpb_1ph,
                "KABEL SR 1PH (M)": sr_1ph
            }, "#FFCC00") # Warna Kuning PLN
            
            st.markdown("### Grafik Pasang vs Bongkar")
            chart_data = pd.DataFrame({
                'Kategori': ['MCB 1PH', 'MCB 3PH', 'LPB 1PH', 'Paska 1PH'],
                'Pasang': [psg_mcb_1ph, psg_mcb_3ph, psg_lpb_1ph, psg_paska_1ph],
                'Bongkar': [bkr_mcb_1ph, bkr_mcb_3ph, bkr_lpb_1ph, bkr_paska_1ph]
            }).set_index('Kategori')
            st.bar_chart(chart_data, use_container_width=True)

        with tab2:
            st.markdown("### 📋 Detail Rincian Per Kategori")
            
            with st.expander("📌 PASANG BARU", expanded=False):
                render_cards(dt_pb, "#00A2E9")

            with st.expander("📌 PERUBAHAN DAYA", expanded=False):
                st.markdown("**Perubahan Daya 1 Fasa**")
                render_cards(dt_pd_1, "#ED1C24")
                st.markdown("**Perubahan Daya 1 Fasa + Migrasi ke LPB**")
                render_cards(dt_pd_1_mig_lpb, "#00A2E9")
                st.markdown("**Perubahan Daya 1 Fasa + Migrasi ke Paska**")
                render_cards(dt_pd_1_mig_paska, "#FFCC00")
                st.markdown("**Perubahan Daya 1 Fasa ke 3 Fasa**")
                render_cards(dt_pd_1_to_3, "#ED1C24")
                st.markdown("**Perubahan Daya 3 Fasa ke 1 Fasa**")
                render_cards(dt_pd_3_to_1, "#00A2E9")
                st.markdown("**Perubahan Daya 3 Fasa**")
                render_cards(dt_pd_3, "#FFCC00")
                st.markdown("**Perubahan Daya 3 Fasa + Migrasi ke Paska**")
                render_cards(dt_pd_3_mig_paska, "#ED1C24")

            with st.expander("📌 MIGRASI", expanded=False):
                render_cards(dt_mig, "#00A2E9")

        with tab3:
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
            
            st.download_button(
                label="📥 Download Hasil Rekap (.txt)",
                data=output_text, file_name="Rekap_LTB.txt", mime="text/plain", type="primary"
            )
            st.markdown(f'<div class="result-box">{output_text}</div>', unsafe_allow_html=True)
        
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
