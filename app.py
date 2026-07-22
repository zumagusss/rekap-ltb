import streamlit as st
import pandas as pd

# ==========================================
# 1. PAGE CONFIG & UI SETUP (PROFESIONAL)
# ==========================================
st.set_page_config(page_title="PLN - Dashboard LTB", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    /* Reset & Global Background */
    .stApp { background-color: #F4F7F9; }
    
    /* Sembunyikan elemen default Streamlit */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Typography Premium */
    h1, h2, h3, p, span { font-family: 'Inter', 'Segoe UI', sans-serif !important; }
    
    /* Navbar Custom */
    .navbar {
        display: flex; align-items: center; justify-content: space-between;
        background: white; padding: 15px 30px; border-radius: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04); margin-bottom: 30px; border-bottom: 4px solid #00A2E9;
    }
    
    /* Uploader Premium UX (Fix Dark Mode Clash) */
    div[data-testid="stFileUploader"] {
        background-color: white; border-radius: 16px; padding: 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04);
    }
    div[data-testid="stFileUploadDropzone"] {
        background-color: #F8FBFE !important; border: 2px dashed #00A2E9 !important; border-radius: 12px !important;
    }
    div[data-testid="stFileUploadDropzone"] * { color: #1E293B !important; font-weight: 500;}
    
    /* Custom CSS Grid untuk Cards (100% Responsif HP & PC) */
    .grid-container {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 15px; margin-bottom: 20px;
    }
    .kpi-card {
        background: white; padding: 24px; border-radius: 16px; 
        box-shadow: 0 4px 15px rgba(0,0,0,0.03); border: 1px solid #EEF2F6;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        display: flex; flex-direction: column; position: relative; overflow: hidden;
    }
    .kpi-card:hover { transform: translateY(-4px); box-shadow: 0 8px 25px rgba(0,162,233,0.15); }
    .kpi-title { font-size: 12px; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; z-index: 2;}
    .kpi-value { font-size: 32px; color: #0F172A; font-weight: 900; z-index: 2;}
    
    /* Aksen Warna PLN pada Cards */
    .card-blue { border-bottom: 4px solid #00A2E9; }
    .card-yellow { border-bottom: 4px solid #FFCC00; }
    .card-red { border-bottom: 4px solid #ED1C24; }
    
    /* Tampilan Rekap Text (Clean Receipt Style) */
    .result-box {
        background-color: white; padding: 30px; border-radius: 12px; border-left: 6px solid #00A2E9;
        box-shadow: 0 10px 30px rgba(0,0,0,0.05); font-family: 'Courier New', monospace; 
        color: #1E293B; font-size: 15px; line-height: 1.8; white-space: pre-wrap; font-weight: 600;
    }
    
    /* Styling Tabs Ala Pill / Segmented Control */
    .stTabs [data-baseweb="tab-list"] { background-color: white; padding: 5px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.02); gap: 5px; margin-bottom:20px;}
    .stTabs [data-baseweb="tab"] { height: 45px; border-radius: 8px; border: none; font-weight: 600; color: #64748B; padding: 0 20px;}
    .stTabs [aria-selected="true"] { background-color: #00A2E9 !important; color: white !important; box-shadow: 0 4px 10px rgba(0,162,233,0.3); }
</style>
""", unsafe_allow_html=True)

# Fungsi Render Grid HTML
def render_cards(data_dict, color_class):
    html = '<div class="grid-container">'
    for title, value in data_dict.items():
        html += f'''
        <div class="kpi-card {color_class}">
            <div class="kpi-title">{title}</div>
            <div class="kpi-value">{value}</div>
        </div>
        '''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

# ==========================================
# 2. HEADER NAVBAR
# ==========================================
col1, col2 = st.columns([1, 15])
with col1:
    try:
        st.image("logopln.png", width=65)
    except:
        st.markdown("⚡")
with col2:
    st.markdown("""
        <div style="padding-top: 8px;">
            <h1 style="color: #00A2E9 !important; font-size: 26px; font-weight: 900; margin:0;">DASHBOARD LTB <span style="color:#FFCC00;">DATEL</span></h1>
            <p style="color: #64748B !important; font-size: 14px; margin:0; font-weight: 500;">Sistem Rekapitulasi Pasang Baru & Perubahan Daya Otomatis</p>
        </div>
    """, unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# 3. LOGIKA & PEMROSESAN DATA
# ==========================================
uploaded_file = st.file_uploader("Upload Data LTB (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Pengecekan Ekstensi File
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv': df = pd.read_csv(uploaded_file)
        else:
            try: df = pd.read_excel(uploaded_file, sheet_name='Sheet1')
            except ValueError: df = pd.read_excel(uploaded_file, sheet_name=0)
                
        # st.success dihapus diganti alert premium di bawah
        st.markdown(f'<div style="background:#E6FBF0; color:#0A7A3B; padding:12px 20px; border-radius:8px; font-weight:600; margin-bottom:20px; border-left: 4px solid #10B981;">✅ Berhasil memproses {len(df):,} baris data.</div>', unsafe_allow_html=True)
        
        # Bersihkan Data
        for col in ['JENIS_TRANSAKSI', 'TARIF', 'TARIF_LAMA']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip().str.upper()
            else:
                df[col] = ''
        
        def cek_phasa(daya):
            try:
                d = float(daya)
                if d <= 5500 or d in [7700, 11000]: return '1PH'
                elif 6600 <= d <= 630000: return '3PH'
            except: pass
            return 'UNKNOWN'

        df['PHASA'] = df['DAYA'].apply(cek_phasa)
        df['PHASA_LAMA'] = df['DAYA_LAMA'].apply(cek_phasa) if 'DAYA_LAMA' in df.columns else 'UNKNOWN'
        df['TIPE'] = df['TARIF'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')
        df['TIPE_LAMA'] = df['TARIF_LAMA'].apply(lambda x: 'LPB' if 'T' in x else 'PASKA')

        # === GLOBAL VARIABLES ===
        psg_paska_1ph = bkr_paska_1ph = psg_paska_3ph_sl = bkr_paska_3ph_sl = psg_paska_3ph_stl = bkr_paska_3ph_stl = 0
        psg_lpb_1ph = bkr_lpb_1ph = psg_lpb_3ph = bkr_lpb_3ph = 0
        psg_mcb_1ph = bkr_mcb_1ph = psg_mcb_3ph = bkr_mcb_3ph = 0
        psg_box = bkr_box = 0

        # === DETAILED TRACKER ===
        dt_pb = {'KWH PASKA 3PH SL (PASANG)': 0, 'KWH LPB 3PH (PASANG)': 0, 'MCB 3PH (PASANG)': 0, 'KWH LPB 1PH (PASANG)': 0, 'KWH PASKA 1PH (PASANG)': 0, 'MCB 1PH (PASANG)': 0}
        dt_pd_1 = {'MCB 1PH (BONGKAR)': 0, 'MCB 1PH (PASANG)': 0}
        dt_pd_1_mig_lpb = {'MCB 1PH (BONGKAR)': 0, 'MCB 1PH (PASANG)': 0, 'KWH PASKA 1PH (BONGKAR)': 0, 'KWH LPB 1PH (PASANG)': 0}
        dt_pd_1_mig_paska = {'MCB 1PH (BONGKAR)': 0, 'MCB 1PH (PASANG)': 0, 'KWH LPB 1PH (BONGKAR)': 0, 'KWH PASKA 1PH (PASANG)': 0}
        dt_pd_1_to_3 = {'LPB 1PH (BONGKAR)': 0, 'PASKA 1PH (BONGKAR)': 0, 'LPB 3 PH (PASANG)': 0, 'PASKA 3 PH (PASANG)': 0, 'MCB 1PH (BONGKAR)': 0, 'MCB 3 PH (PASANG)': 0}
        dt_pd_3_to_1 = {'MCB 3PH (BONGKAR)': 0, 'LPB 3PH (BONGKAR)': 0, 'PASKA 3PH (BONGKAR)': 0, 'MCB 1PH (PASANG)': 0, 'LPB 1PH (PASANG)': 0, 'PASKA 1PH (PASANG)': 0}
        dt_pd_3 = {'MCB 3PH (BONGKAR)': 0, 'MCB 3PH (PASANG)': 0}
        dt_pd_3_mig_paska = {'MCB 3PH (BONGKAR)': 0, 'MCB 3PH (PASANG)': 0, 'KWH LPB 3PH (BONGKAR)': 0, 'KWH PASKA 3PH SL (PASANG)': 0}
        dt_mig = {'KWH LPB 3 FASA (PASANG)': 0, 'KWH PASKA 3 FASA SL (BONGKAR)': 0, 'KWH LPB 1PH (PASANG)': 0, 'KWH PASKA 1PH (BONGKAR)': 0, 'KWH PASKA 1PH (PASANG)': 0, 'LPB 1PH (BONGKAR)': 0, 'KWH PASKA 3PH (PASANG)': 0, 'LPB 3PH (BONGKAR)': 0}

        # Looping Logika Perhitungan
        for _, row in df.iterrows():
            trx = row['JENIS_TRANSAKSI']
            phasa = row['PHASA']
            tipe = row['TIPE']
            phasa_lama = row['PHASA_LAMA']
            tipe_lama = row['TIPE_LAMA']
            
            if trx == 'PASANG BARU':
                if phasa == '1PH':
                    psg_mcb_1ph += 1; dt_pb['MCB 1PH (PASANG)'] += 1
                    if tipe == 'PASKA': psg_paska_1ph += 1; dt_pb['KWH PASKA 1PH (PASANG)'] += 1
                    else: psg_lpb_1ph += 1; dt_pb['KWH LPB 1PH (PASANG)'] += 1
                elif phasa == '3PH':
                    psg_mcb_3ph += 1; dt_pb['MCB 3PH (PASANG)'] += 1
                    if tipe == 'PASKA': psg_paska_3ph_sl += 1; dt_pb['KWH PASKA 3PH SL (PASANG)'] += 1
                    else: psg_lpb_3ph += 1; dt_pb['KWH LPB 3PH (PASANG)'] += 1
                    
            elif trx == 'PERUBAHAN DAYA':
                if phasa_lama == '1PH' and phasa == '1PH':
                    bkr_mcb_1ph += 1; psg_mcb_1ph += 1
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                        dt_pd_1_mig_lpb['MCB 1PH (BONGKAR)'] += 1; dt_pd_1_mig_lpb['MCB 1PH (PASANG)'] += 1
                        dt_pd_1_mig_lpb['KWH PASKA 1PH (BONGKAR)'] += 1; dt_pd_1_mig_lpb['KWH LPB 1PH (PASANG)'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                        dt_pd_1_mig_paska['MCB 1PH (BONGKAR)'] += 1; dt_pd_1_mig_paska['MCB 1PH (PASANG)'] += 1
                        dt_pd_1_mig_paska['KWH LPB 1PH (BONGKAR)'] += 1; dt_pd_1_mig_paska['KWH PASKA 1PH (PASANG)'] += 1
                    else:
                        dt_pd_1['MCB 1PH (BONGKAR)'] += 1; dt_pd_1['MCB 1PH (PASANG)'] += 1
                        
                elif phasa_lama == '1PH' and phasa == '3PH':
                    bkr_mcb_1ph += 1; psg_mcb_3ph += 1
                    dt_pd_1_to_3['MCB 1PH (BONGKAR)'] += 1; dt_pd_1_to_3['MCB 3 PH (PASANG)'] += 1
                    if tipe_lama == 'LPB': bkr_lpb_1ph += 1; dt_pd_1_to_3['LPB 1PH (BONGKAR)'] += 1
                    else: bkr_paska_1ph += 1; dt_pd_1_to_3['PASKA 1PH (BONGKAR)'] += 1
                    if tipe == 'LPB': psg_lpb_3ph += 1; dt_pd_1_to_3['LPB 3 PH (PASANG)'] += 1
                    else: psg_paska_3ph_sl += 1; dt_pd_1_to_3['PASKA 3 PH (PASANG)'] += 1
                        
                elif phasa_lama == '3PH' and phasa == '1PH':
                    bkr_mcb_3ph += 1; psg_mcb_1ph += 1
                    dt_pd_3_to_1['MCB 3PH (BONGKAR)'] += 1; dt_pd_3_to_1['MCB 1PH (PASANG)'] += 1
                    if tipe_lama == 'LPB': bkr_lpb_3ph += 1; dt_pd_3_to_1['LPB 3PH (BONGKAR)'] += 1
                    else: bkr_paska_3ph_sl += 1; dt_pd_3_to_1['PASKA 3PH (BONGKAR)'] += 1
                    if tipe == 'LPB': psg_lpb_1ph += 1; dt_pd_3_to_1['LPB 1PH (PASANG)'] += 1
                    else: psg_paska_1ph += 1; dt_pd_3_to_1['PASKA 1PH (PASANG)'] += 1
                        
                elif phasa_lama == '3PH' and phasa == '3PH':
                    bkr_mcb_3ph += 1; psg_mcb_3ph += 1
                    if tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1
                        dt_pd_3_mig_paska['MCB 3PH (BONGKAR)'] += 1; dt_pd_3_mig_paska['MCB 3PH (PASANG)'] += 1
                        dt_pd_3_mig_paska['KWH LPB 3PH (BONGKAR)'] += 1; dt_pd_3_mig_paska['KWH PASKA 3PH SL (PASANG)'] += 1
                    else: 
                        dt_pd_3['MCB 3PH (BONGKAR)'] += 1; dt_pd_3['MCB 3PH (PASANG)'] += 1
                        
            elif trx == 'MIGRASI':
                if phasa == '1PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_1ph += 1; psg_lpb_1ph += 1
                        dt_mig['KWH PASKA 1PH (BONGKAR)'] += 1; dt_mig['KWH LPB 1PH (PASANG)'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_1ph += 1; psg_paska_1ph += 1
                        dt_mig['LPB 1PH (BONGKAR)'] += 1; dt_mig['KWH PASKA 1PH (PASANG)'] += 1
                elif phasa == '3PH':
                    if tipe_lama == 'PASKA' and tipe == 'LPB': 
                        bkr_paska_3ph_sl += 1; psg_lpb_3ph += 1
                        dt_mig['KWH PASKA 3 FASA SL (BONGKAR)'] += 1; dt_mig['KWH LPB 3 FASA (PASANG)'] += 1
                    elif tipe_lama == 'LPB' and tipe == 'PASKA': 
                        bkr_lpb_3ph += 1; psg_paska_3ph_sl += 1
                        dt_mig['LPB 3PH (BONGKAR)'] += 1; dt_mig['KWH PASKA 3PH (PASANG)'] += 1

        sr_1ph = psg_mcb_1ph * 20
        sr_3ph = psg_mcb_3ph * 30
        
        dt_pb['SR 1PH (PASANG)'] = f"{sr_1ph:,} m"
        dt_pb['SR 3PH (PASANG)'] = f"{sr_3ph:,} m"
        dt_pb['KWH PASKA 3PH STL (PASANG)'] = 0
        dt_pb['BOX APP 53 KVA (PASANG)'] = 0

        # ==========================================
        # 4. TABS UI RENDER
        # ==========================================
        tab1, tab2, tab3 = st.tabs(["📊 Executive Summary", "🔍 Detail Transaksi", "📝 Format Laporan Text"])
        
        with tab1:
            st.markdown("<h3 style='color:#1E293B !important; font-size:18px; margin: 10px 0 20px 0;'>Highlights Bulan Ini</h3>", unsafe_allow_html=True)
            render_cards({
                "TOTAL BERKAS": f"{len(df):,}",
                "PASANG MCB 1PH": f"{psg_mcb_1ph:,}",
                "PASANG LPB 1PH": f"{psg_lpb_1ph:,}",
                "KABEL SR 1PH": f"{sr_1ph:,} M"
            }, "card-blue")
            
            st.markdown("<h3 style='color:#1E293B !important; font-size:18px; margin: 30px 0 20px 0;'>Grafik Komparasi Material (Pasang vs Bongkar)</h3>", unsafe_allow_html=True)
            chart_data = pd.DataFrame({
                'Kategori': ['MCB 1PH', 'MCB 3PH', 'LPB 1PH', 'Paska 1PH'],
                'Pasang': [psg_mcb_1ph, psg_mcb_3ph, psg_lpb_1ph, psg_paska_1ph],
                'Bongkar': [bkr_mcb_1ph, bkr_mcb_3ph, bkr_lpb_1ph, bkr_paska_1ph]
            }).set_index('Kategori')
            st.bar_chart(chart_data, use_container_width=True, height=350)

        with tab2:
            # Custom CSS biar font expander lebih tebal
            st.markdown("""<style>.streamlit-expanderHeader { font-weight: 700 !important; color: #00A2E9 !important; font-size: 15px !important; }</style>""", unsafe_allow_html=True)
            
            with st.expander("📌 RINCIAN PASANG BARU", expanded=False):
                render_cards(dt_pb, "card-blue")

            with st.expander("📌 RINCIAN PERUBAHAN DAYA", expanded=False):
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>A. Perubahan Daya 1 Fasa (Tetap)</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_1, "card-red")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>B. 1 Fasa + Migrasi ke LPB</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_1_mig_lpb, "card-yellow")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>C. 1 Fasa + Migrasi ke Paska</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_1_mig_paska, "card-blue")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>D. Lintas Fasa (1 Fasa ke 3 Fasa)</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_1_to_3, "card-red")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>E. Lintas Fasa (3 Fasa ke 1 Fasa)</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_3_to_1, "card-yellow")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>F. Perubahan Daya 3 Fasa (Tetap)</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_3, "card-blue")
                st.markdown("<h4 style='color:#334155 !important; font-size:14px; margin-top:10px;'>G. 3 Fasa + Migrasi ke Paska</h4>", unsafe_allow_html=True)
                render_cards(dt_pd_3_mig_paska, "card-red")

            with st.expander("📌 RINCIAN MIGRASI MURNI", expanded=False):
                render_cards(dt_mig, "card-yellow")

        with tab3:
            output_text = f"""LTB Bulan Ini					
					
kWh Meter Elektronik (Pascabayar)					
1 Phasa\t\t:\tPasang\t{psg_paska_1ph}\tBongkar\t{bkr_paska_1ph}
3 Phasa SL\t:\tPasang\t{psg_paska_3ph_sl}\tBongkar\t{bkr_paska_3ph_sl}
3 Phasa STL\t:\tPasang\t{psg_paska_3ph_stl}\tBongkar\t{bkr_paska_3ph_stl}
					
kWh Meter Elektronik Prabayar ( Pre Paid )					
1 Phasa\t\t:\tPasang\t{psg_lpb_1ph}\tBongkar\t{bkr_lpb_1ph}
3 Phasa\t\t:\tPasang\t{psg_lpb_3ph}\tBongkar\t{bkr_lpb_3ph}
					
MCB 1 Phasa\t:\tPasang\t{psg_mcb_1ph}\tBongkar\t{bkr_mcb_1ph}
MCB 3 Phasa\t:\tPasang\t{psg_mcb_3ph}\tBongkar\t{bkr_mcb_3ph}
BOX APP\t\t:\tPasang\t{psg_box}\tBongkar\t{bkr_box}
					
SR 1 Phasa (10mm)					
Pasang\t\t:\tPasang\t{sr_1ph}\tBongkar\t0
SR 3 Phasa (16mm)					
Pasang\t\t:\tPasang\t{sr_3ph}\tBongkar\t0"""
            
            st.markdown(f'<div class="result-box">{output_text}</div>', unsafe_allow_html=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Download Rekap LTB (.txt)",
                data=output_text, file_name="Rekap_LTB_Otomatis.txt", mime="text/plain", type="primary", use_container_width=True
            )
        
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan struktur file Excel: {e}. Pastikan menggunakan template DATEL yang benar.")
