import streamlit as st
import pandas as pd
import os

import sqlite3
try:
    if os.path.exists("afp_database.db"):
        if os.path.getsize("afp_database.db") == 0:
            os.remove("afp_database.db")
        else:
            with sqlite3.connect("afp_database.db") as conn:
                cur = conn.cursor()
                cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='programs'")
                if not cur.fetchone():
                    conn.close()
                    os.remove("afp_database.db")
except Exception:
    pass
import datetime
from database import (
    get_machines,
    get_active_queue,
    get_completed_lots,
    get_lot_details,
    update_lot_sequence,
    update_lot_fields,
    add_new_lot,
    delete_lot,
    get_genel_makine,
    update_genel_makine,
    DB_PATH
)
from pdf_generator import generate_pdf
import migrate
from streamlit_sortables import sort_items
import altair as alt

# Set page config
st.set_page_config(
    page_title="AFP Planlama ve İmalat Yönetim Sistemi",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cloud Mode Authentication
is_cloud = False
try:
    if st.secrets.get("CLOUD_MODE", False):
        is_cloud = True
        if "authenticated" not in st.session_state:
            st.session_state.authenticated = False
            
        if not st.session_state.authenticated:
            st.markdown("<h2 style='text-align: center;'>AFP Planlama (Bulut Erişimi)</h2>", unsafe_allow_html=True)
            password = st.text_input("Giriş Şifresi:", type="password")
            if st.button("Giriş Yap"):
                if password == "1923":
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Hatalı şifre!")
            st.stop()
except FileNotFoundError:
    pass # No secrets file, running locally

# Initialize session state for tracking modified machines
if "modified_machines" not in st.session_state:
    st.session_state.modified_machines = set()

# Custom Premium Styling (CSS)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&display=swap');
    
    /* Global Styles */
    html, body, [class*="css"] {
        font-family: 'Outfit', sans-serif;
    }
    
    /* Main Header with Gradient */
    .header-container {
        background: linear-gradient(135deg, #1f4e79 0%, #2e75b6 100%);
        padding: 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .header-container h1 {
        margin: 0;
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.5px;
    }
    .header-container p {
        margin: 5px 0 0 0;
        opacity: 0.9;
        font-size: 1rem;
    }
    
    /* KPI Card Style */
    .kpi-card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-left: 5px solid #1f4e79;
        margin-bottom: 15px;
    }
    .kpi-title {
        font-size: 0.9rem;
        color: #666;
        font-weight: 600;
        text-transform: uppercase;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1f4e79;
        margin-top: 5px;
    }
    
    /* Section headers */
    .section-title {
        color: #1f4e79;
        font-weight: 600;
        border-bottom: 2px solid #e9ecef;
        padding-bottom: 8px;
        margin-top: 15px;
        margin-bottom: 15px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
/* Make pills (kutucuklar) smaller */
div[data-testid="stPills"] button, div[data-testid="stPills"] label {
    padding: 2px 10px !important;
    font-size: 11px !important;
    min-height: 25px !important;
    line-height: 1.2 !important;
}
</style>
""", unsafe_allow_html=True)


# Main App Header
st.markdown("""
<div class="header-container">
    <h1>⚙️ AFP Planlama & İmalat Yönetim Sistemi</h1>
    <p>Excel karmaşasından uzak, SQLite veritabanı destekli yüksek performanslı üretim takip yazılımı</p>
</div>
""", unsafe_allow_html=True)

# Main Module Navigation
main_module = st.pills("Ana Modül", ["Hammadde", "Soğuk Şekillendirme"], default="Soğuk Şekillendirme", label_visibility="collapsed")
st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)

if main_module == "Hammadde":
    import hammadde
    hammadde.render_hammadde()
    st.stop()
elif main_module is None:
    st.info("Lütfen devam etmek için yukarıdan bir modül seçin.")
    st.stop()

# Tabs
tab_dashboard, tab_simulator, tab_ai_planning, tab_ai_compare, tab_planning, tab_genel_makine, tab_work_order, tab_add_lot, tab_sync = st.tabs([
    "📊 Yönetim Paneli (Dashboard)",
    "🔮 Yönetim Paneli Simülatör",
    "🤖 Otonom Kapasite Planlama",
    "⚖️ AI Karşılaştırma",
    "⚙️ Makine Planlama Kuyruğu",
    "⚙️ Genel Makine Ayarları",
    "📄 İş Emri Kartı Yazdır",
    "➕ Yeni Lot Ekle",
    "🔄 Excel Veri Göçü & Eşleme"
])

# Initialize DB if not exists
if not os.path.exists(DB_PATH):
    st.warning("Veritabanı bulunamadı. Lütfen 'Excel Veri Göçü & Eşleme' sekmesinden verileri aktarın.")

# ==========================================
# TAB 1: DASHBOARD
# ==========================================
with tab_dashboard:
    if os.path.exists(DB_PATH):
        try:
            conn = migrate.sqlite3.connect(DB_PATH)
            # Query stats
            total_all_net_tonaj = pd.read_sql("SELECT SUM(CAST([NET TONAJ] AS REAL)) FROM programs WHERE ([SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = '')", conn).iloc[0,0]
            active_lots = pd.read_sql("SELECT COUNT(*) FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''", conn).iloc[0,0]
            completed_lots = pd.read_sql("SELECT COUNT(*) FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTTİ'", conn).iloc[0,0]
            total_qty = pd.read_sql("SELECT SUM(CAST(ADET AS REAL)) FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ'", conn).iloc[0,0]
            total_kg = pd.read_sql("SELECT SUM(CAST(KG AS REAL)) FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ'", conn).iloc[0,0]
            total_delayed_kg = pd.read_sql("SELECT SUM(CAST([geciken net tonaj] AS REAL)) FROM programs WHERE ([SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = '')", conn).iloc[0,0]
            
            # Query customer-based delayed tonnage with region classification
            query_delayed = """
                SELECT 
                    IFNULL([MÜŞTERİ], 'BİLİNMEYEN') AS Müşteri, 
                    IFNULL([BÖLGE], 'BİLİNMEYEN') AS Bölge,
                    SUM(CAST([geciken net tonaj] AS REAL)) AS [Geciken Tonaj (KG)],
                    COUNT(CASE WHEN CAST([geciken net tonaj] AS REAL) > 0 THEN 1 END) AS [Geciken Lot Sayısı]
                FROM programs 
                WHERE ([SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = '')
                GROUP BY [MÜŞTERİ], [BÖLGE] 
                HAVING [Geciken Tonaj (KG)] > 0 
                ORDER BY [Geciken Tonaj (KG)] DESC
            """
            df_delayed = pd.read_sql_query(query_delayed, conn)
            
            # Query active weights by customer using [NET TONAJ]
            query_cust = """
                SELECT 
                    IFNULL([MÜŞTERİ], 'BİLİNMEYEN') AS Müşteri,
                    SUM(CAST([NET TONAJ] AS REAL)) AS [Toplam Net Tonaj (KG)],
                    COUNT(*) AS [Lot Sayısı]
                FROM programs
                WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''
                GROUP BY [MÜŞTERİ]
                HAVING [Toplam Net Tonaj (KG)] > 0
                ORDER BY [Toplam Net Tonaj (KG)] DESC
            """
            df_cust = pd.read_sql_query(query_cust, conn)
            
            # Query active weights by diameter (anma çapı) using [NET TONAJ]
            query_cap = """
                SELECT 
                    IFNULL([ÇAP], 'BİLİNMEYEN') AS Çap,
                    SUM(CAST([NET TONAJ] AS REAL)) AS [Toplam Net Tonaj (KG)],
                    COUNT(*) AS [Lot Sayısı]
                FROM programs
                WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''
                GROUP BY [ÇAP]
                HAVING [Toplam Net Tonaj (KG)] > 0
                ORDER BY [Toplam Net Tonaj (KG)] DESC
            """
            df_cap = pd.read_sql_query(query_cap, conn)
            
            # Query active weights by standards using [NET TONAJ]
            query_std = """
                SELECT 
                    IFNULL([STANDART], 'BİLİNMEYEN') AS Standart,
                    SUM(CAST([NET TONAJ] AS REAL)) AS [Toplam Net Tonaj (KG)],
                    COUNT(*) AS [Lot Sayısı]
                FROM programs
                WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''
                GROUP BY [STANDART]
                HAVING [Toplam Net Tonaj (KG)] > 0
                ORDER BY [Toplam Net Tonaj (KG)] DESC
            """
            df_std = pd.read_sql_query(query_std, conn)
            
            conn.close()
            
            # Split data by region (İHRACAT vs YURT İÇİ) and explicitly sort descending by tonnage
            df_ihr = df_delayed[df_delayed["Bölge"] == "İHRACAT"].copy()
            if not df_ihr.empty:
                df_ihr = df_ihr.sort_values(by="Geciken Tonaj (KG)", ascending=False)
                df_ihr["Geciken Tonaj (Ton)"] = (df_ihr["Geciken Tonaj (KG)"] / 1000.0).round(2)
                
            df_yur = df_delayed[df_delayed["Bölge"].str.contains("YURT", na=False)].copy()
            if not df_yur.empty:
                df_yur = df_yur.sort_values(by="Geciken Tonaj (KG)", ascending=False)
                df_yur["Geciken Tonaj (Ton)"] = (df_yur["Geciken Tonaj (KG)"] / 1000.0).round(2)
                
            ihr_ton = df_ihr["Geciken Tonaj (Ton)"].sum() if not df_ihr.empty else 0.0
            yur_ton = df_yur["Geciken Tonaj (Ton)"].sum() if not df_yur.empty else 0.0
            
            # Formulate Qty/KG strings
            total_all_net_tonaj_str = f"{total_all_net_tonaj:,.1f} KG" if pd.notnull(total_all_net_tonaj) else "0.0 KG"
            total_qty_str = f"{total_qty:,.0f} Adet" if total_qty else "0 Adet"
            total_kg_str = f"{total_kg:,.1f} KG" if total_kg else "0.0 KG"
            total_delayed_str = f"{total_delayed_kg:,.1f} KG" if total_delayed_kg else "0.0 KG"
            if total_delayed_kg:
                total_delayed_str += f" ({total_delayed_kg/1000.0:,.1f} Ton)"
            
            # KPI Cards row
            kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
            with kpi1:
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #2e75b6;">
                    <div class="kpi-title">Toplam Net Tonaj</div>
                    <div class="kpi-value">{total_all_net_tonaj_str}</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi2:
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #f4b183;">
                    <div class="kpi-title">Aktif Üretim Kuyruğu</div>
                    <div class="kpi-value">{active_lots:,} Lot</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi3:
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #a9d18e;">
                    <div class="kpi-title">Tamamlanan Üretimler</div>
                    <div class="kpi-value">{completed_lots:,} Lot</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi4:
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #ffc000;">
                    <div class="kpi-title">Aktif Planlanan Ağırlık</div>
                    <div class="kpi-value">{total_kg_str}</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi5:
                st.markdown(f"""
                <div class="kpi-card" style="border-left-color: #c00000; border-left-width: 5px;">
                    <div class="kpi-title" style="color: #c00000;">Toplam Geciken Tonaj</div>
                    <div class="kpi-value" style="color: #c00000; font-size: 1.1rem; line-height: 1.8rem; font-weight: 700;">{total_delayed_str}</div>
                    <div style="font-size: 0.75rem; color: #555; margin-top: 5px; font-weight: 600;">
                        ✈️ İhr: {ihr_ton:,.1f} T | 🏠 Yurtiçi: {yur_ton:,.1f} T
                    </div>
                </div>
                """, unsafe_allow_html=True)

            # Termin Ayı Bazlı Grafik
            query_termin = """
                SELECT 
                    [TERMİN TARİHİ],
                    CAST(REPLACE([NET TONAJ], ',', '') AS REAL) AS [NET TONAJ],
                    CAST(REPLACE([KALAN ADET], ',', '') AS REAL) AS [KALAN ADET]
                FROM programs
                WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' OR [SİPARİŞ DURUMU] IS NULL OR [SİPARİŞ DURUMU] = ''
            """
            import sqlite3
            with sqlite3.connect(DB_PATH) as conn_chart:
                df_t = pd.read_sql_query(query_termin, conn_chart)
            df_t['Termin_Date'] = pd.to_datetime(df_t['TERMİN TARİHİ'], format='%d.%m.%Y', errors='coerce')
            df_t = df_t.dropna(subset=['Termin_Date'])
            df_t['Termin Ayı'] = df_t['Termin_Date'].dt.strftime('%Y-%m')
            df_grp = df_t.groupby('Termin Ayı')[['NET TONAJ', 'KALAN ADET']].sum().reset_index()
            
            if not df_grp.empty:
                st.markdown("<h3 class='section-title'>📈 Termin Ayı Bazlı Net Tonaj ve Kalan Adet</h3>", unsafe_allow_html=True)
                import altair as alt
                # Convert Tonaj to Ton for cleaner numbers
                df_grp['TON'] = (df_grp['NET TONAJ'] / 1000).astype(int)
                
                base = alt.Chart(df_grp).encode(x=alt.X('Termin Ayı:N', title='Termin Ayı', axis=alt.Axis(labelAngle=-45, labelFontSize=13, titleFontSize=14)))
                
                # TON (Bar + Label)
                # Koyu lacivert/mavi (Sunum için şık ve tok bir renk)
                bar = base.mark_bar(color='#1f4e79', opacity=0.9).encode(
                    y=alt.Y('TON:Q', title=None, axis=None)
                )
                bar_text = base.mark_text(align='center', baseline='bottom', dy=-5, dx=-15, fontSize=13, color='#1f4e79', fontWeight='bold').encode(
                    y=alt.Y('TON:Q', axis=None),
                    text=alt.Text('TON:Q', format=',.0f')
                )
                
                # KALAN ADET (Line + Label)
                # Canlı bir turuncu/kiremit rengi (Dikkat çekici)
                line = base.mark_line(color='#d9534f', point=alt.OverlayMarkDef(color='#d9534f', size=100), strokeWidth=3).encode(
                    y=alt.Y('KALAN ADET:Q', title=None, axis=None)
                )
                line_text = base.mark_text(align='center', baseline='bottom', dy=-15, dx=15, fontSize=13, color='#d9534f', fontWeight='bold').encode(
                    y=alt.Y('KALAN ADET:Q', axis=None),
                    text=alt.Text('KALAN ADET:Q', format=',.0f')
                )
                
                chart = alt.layer(bar, bar_text, line, line_text).resolve_scale(y='independent').properties(height=400).configure_view(stroke='transparent')
                st.altair_chart(chart, use_container_width=True)

            # ==========================================
            # ÜRETİM TAHMİN SİMÜLATÖRÜ
            # ==========================================

            st.markdown("<hr style='margin:30px 0; border:0; border-top:1px solid #ccc;'>", unsafe_allow_html=True)
            
            # X Kesme Noktası Genel Özet Tablosu
            st.markdown("<h3 class='section-title'>🔍 Tüm Makinelerin X Kesme Noktası Özet Tablosu</h3>", unsafe_allow_html=True)
            
            machines = get_machines()
            cutoff_data = []
            for m in machines:
                q = get_active_queue(m)
                if not q.empty and 'X_CUTOFF' in q.columns:
                    x_rows = q[q['X_CUTOFF'] == 1]
                    if not x_rows.empty:
                        cutoff_idx = x_rows.index[0]
                        cutoff_df = q.iloc[:cutoff_idx+1]
                        sum_k_adet = pd.to_numeric(cutoff_df['KALAN ADET'], errors='coerce').sum()
                        sum_k_kg = pd.to_numeric(cutoff_df['NET TONAJ'], errors='coerce').sum()
                        sum_x_gun = pd.to_numeric(cutoff_df['X GÜNLÜK İŞ'], errors='coerce').sum()
                        
                        cutoff_data.append({
                            "Makine": m,
                            "K. ADET (Kesme Noktası)": sum_k_adet,
                            "K.KG (Kesme Noktası)": sum_k_kg,
                            "X GÜNLÜK İŞ (Kesme Noktası)": sum_x_gun
                        })
            
            if cutoff_data:
                df_cutoff_summary = pd.DataFrame(cutoff_data)
                df_cutoff_summary = df_cutoff_summary.sort_values(by="X GÜNLÜK İŞ (Kesme Noktası)", ascending=True)
                
                col_conf = {
                    "K. ADET (Kesme Noktası)": st.column_config.NumberColumn("K. ADET (Kesme Noktası)", format="%d"),
                    "K.KG (Kesme Noktası)": st.column_config.NumberColumn("K.KG (Kesme Noktası)", format="%.0f"),
                    "X GÜNLÜK İŞ (Kesme Noktası)": st.column_config.ProgressColumn(
                        "X GÜNLÜK İŞ (Kesme Noktası)", 
                        format="%.1f",
                        min_value=0,
                        max_value=float(df_cutoff_summary["X GÜNLÜK İŞ (Kesme Noktası)"].max() or 1)
                    )
                }
                
                st.dataframe(df_cutoff_summary, use_container_width=True, hide_index=True, column_config=col_conf)
                
            else:
                st.info("Herhangi bir makinede X kesme noktası hesaplaması bulunamadı.")
                
            st.markdown("<hr style='margin:30px 0; border:0; border-top:1px solid #ccc;'>", unsafe_allow_html=True)
                
            st.markdown("<hr style='margin:30px 0; border:0; border-top:1px solid #ccc;'>", unsafe_allow_html=True)
            # 2. Bölgesel Müşteri Geciken Net Tonaj Analizi
            st.markdown("<h3 class='section-title'>🌍 Bölgesel Müşteri ve Toplam Geciken Net Tonaj Analizi</h3>", unsafe_allow_html=True)
            
            if not df_delayed.empty:
                df_delayed_chart = df_delayed.copy()
                df_delayed_chart["Geciken Tonaj (Ton)"] = (df_delayed_chart["Geciken Tonaj (KG)"] / 1000.0).round(2)
                # Normalize Region for charting
                df_delayed_chart["Bölge_Ana"] = df_delayed_chart["Bölge"].apply(
                    lambda x: "İHRACAT" if "İHRACAT" in str(x).upper() else ("YURTİÇİ" if "YURT" in str(x).upper() else str(x))
                )
                
                df_total_region = df_delayed_chart.groupby("Bölge_Ana")["Geciken Tonaj (Ton)"].sum().reset_index()
                
                col_donut, col_bar = st.columns([1, 2])
                
                with col_donut:
                    st.markdown("<h4 style='color:#1f4e79;'>🍩 Toplam Geciken (Bölge Bazlı)</h4>", unsafe_allow_html=True)
                    
                    base_donut = alt.Chart(df_total_region).encode(
                        theta=alt.Theta(field="Geciken Tonaj (Ton)", type="quantitative", stack=True),
                        color=alt.Color(field="Bölge_Ana", type="nominal", scale=alt.Scale(domain=['İHRACAT', 'YURTİÇİ'], range=['#2e75b6', '#c00000']), legend=alt.Legend(title="Bölge")),
                        tooltip=['Bölge_Ana', 'Geciken Tonaj (Ton)']
                    )
                    arc = base_donut.mark_arc(innerRadius=40, outerRadius=120)
                    # Adding data labels for donut on the colored slices
                    text = base_donut.mark_text(radius=80, size=18, fontWeight="bold").encode(
                        text=alt.Text('Geciken Tonaj (Ton):Q', format='.1f'),
                        color=alt.value('white')
                    )
                    chart_donut = (arc + text).properties(height=300)
                    
                    st.altair_chart(chart_donut, use_container_width=True)
                    
                with col_bar:
                    st.markdown("<h4 style='color:#1f4e79;'>📊 En Çok Geciken İlk 15 Müşteri (Genel)</h4>", unsafe_allow_html=True)
                    df_top_customers = df_delayed_chart.sort_values("Geciken Tonaj (Ton)", ascending=False).head(15)
                    
                    base_bar = alt.Chart(df_top_customers).encode(
                        x=alt.X('Müşteri:N', sort=df_top_customers['Müşteri'].tolist(), axis=alt.Axis(labelOverlap=False, labelLimit=250)),
                        y=alt.Y('Geciken Tonaj (Ton):Q', axis=alt.Axis(title='Geciken Tonaj (Ton)')),
                        color=alt.Color('Bölge_Ana:N', scale=alt.Scale(domain=['İHRACAT', 'YURTİÇİ'], range=['#2e75b6', '#c00000']), legend=None),
                        tooltip=['Müşteri', 'Bölge', 'Geciken Tonaj (Ton)', 'Geciken Tonaj (KG)', 'Geciken Lot Sayısı']
                    )
                    bars = base_bar.mark_bar()
                    # Adding data labels for bar chart
                    text_bar = base_bar.mark_text(align='center', baseline='bottom', dy=-5, size=15, fontWeight='bold').encode(
                        text=alt.Text('Geciken Tonaj (Ton):Q', format='.1f'),
                        color=alt.value('black')
                    )
                    chart_bar = (bars + text_bar).properties(height=300)
                    
                    st.altair_chart(chart_bar, use_container_width=True)
                    
            else:
                st.info("Geciken sipariş bulunmamaktadır.")
                
            st.markdown("<hr style='margin:30px 0; border:0; border-top:1px solid #ccc;'>", unsafe_allow_html=True)


            
            # 3. Aktif Sipariş Analizleri (Durumu Bitmedi Olanlar)
            st.markdown("<h3 class='section-title'>📈 Aktif Sipariş Ağırlık Analizleri (Geciken & Gecikmeyen Tümü)</h3>", unsafe_allow_html=True)
            
            # 3.1 Müşteri Bazlı
            st.markdown("<h4 style='color:#2e75b6; margin-top:20px;'>👤 Müşteri Bazlı Aktif Sipariş Dağılımı</h4>", unsafe_allow_html=True)
            if not df_cust.empty:
                df_cust["Toplam Net Tonaj (Ton)"] = (df_cust["Toplam Net Tonaj (KG)"] / 1000.0).round(2)
                col_c1, col_c2 = st.columns([2, 1])
                with col_c1:
                    chart_cust = alt.Chart(df_cust.head(10)).mark_bar(color='#2e75b6').encode(
                        x=alt.X('Müşteri:N', sort=alt.EncodingSortField(field='Toplam Net Tonaj (Ton)', order='descending')),
                        y='Toplam Net Tonaj (Ton):Q',
                        tooltip=['Müşteri', 'Toplam Net Tonaj (Ton)', 'Toplam Net Tonaj (KG)', 'Lot Sayısı']
                    ).properties(height=280)
                    st.altair_chart(chart_cust, use_container_width=True)
                with col_c2:
                    st.dataframe(
                        df_cust[["Müşteri", "Toplam Net Tonaj (KG)", "Lot Sayısı"]],
                        use_container_width=True,
                        hide_index=True
                    )
            else:
                st.info("Aktif sipariş bulunmamaktadır.")
                
            st.markdown("<hr style='margin:20px 0; border:0; border-top:1px dashed #eee;'>", unsafe_allow_html=True)
            
            # 3.2 Çap Bazlı
            st.markdown("<h4 style='color:#27ae60; margin-top:10px;'>📏 Çap (Anma Çapı) Bazlı Aktif Sipariş Dağılımı</h4>", unsafe_allow_html=True)
            if not df_cap.empty:
                df_cap["Toplam Net Tonaj (Ton)"] = (df_cap["Toplam Net Tonaj (KG)"] / 1000.0).round(2)
                col_cap1, col_cap2 = st.columns([2, 1])
                with col_cap1:
                    chart_cap = alt.Chart(df_cap.head(10)).mark_bar(color='#27ae60').encode(
                        x=alt.X('Çap:N', sort=alt.EncodingSortField(field='Toplam Net Tonaj (Ton)', order='descending')),
                        y='Toplam Net Tonaj (Ton):Q',
                        tooltip=['Çap', 'Toplam Net Tonaj (Ton)', 'Toplam Net Tonaj (KG)', 'Lot Sayısı']
                    ).properties(height=280)
                    st.altair_chart(chart_cap, use_container_width=True)
                with col_cap2:
                    st.dataframe(
                        df_cap[["Çap", "Toplam Net Tonaj (KG)", "Lot Sayısı"]],
                        use_container_width=True,
                        hide_index=True
                    )
            else:
                st.info("Aktif sipariş bulunmamaktadır.")
                
            st.markdown("<hr style='margin:20px 0; border:0; border-top:1px dashed #eee;'>", unsafe_allow_html=True)
            
            # 3.3 Standart Bazlı
            st.markdown("<h4 style='color:#8e44ad; margin-top:10px;'>📋 Standart Bazlı Aktif Sipariş Dağılımı</h4>", unsafe_allow_html=True)
            if not df_std.empty:
                df_std["Toplam Net Tonaj (Ton)"] = (df_std["Toplam Net Tonaj (KG)"] / 1000.0).round(2)
                col_std1, col_std2 = st.columns([2, 1])
                with col_std1:
                    chart_std = alt.Chart(df_std.head(10)).mark_bar(color='#8e44ad').encode(
                        x=alt.X('Standart:N', sort=alt.EncodingSortField(field='Toplam Net Tonaj (Ton)', order='descending')),
                        y='Toplam Net Tonaj (Ton):Q',
                        tooltip=['Standart', 'Toplam Net Tonaj (Ton)', 'Toplam Net Tonaj (KG)', 'Lot Sayısı']
                    ).properties(height=280)
                    st.altair_chart(chart_std, use_container_width=True)
                with col_std2:
                    st.dataframe(
                        df_std[["Standart", "Toplam Net Tonaj (KG)", "Lot Sayısı"]],
                        use_container_width=True,
                        hide_index=True
                    )
            else:
                st.info("Aktif sipariş bulunmamaktadır.")
                
            
                
        except Exception as e:
            st.error(f"Gösterge paneli yüklenirken hata oluştu: {e}")
    else:
        st.info("Lütfen önce veritabanını oluşturun.")

# ==========================================
# TAB 2: MACHINE PLANNING
# ==========================================

with tab_simulator:
    st.markdown("<h2 class='section-title'>🔮 Yönetim Paneli - Üretim Tahmin Simülatörü</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color:#555;'>Belirttiğiniz gün sayısı boyunca tüm makinelerin kesintisiz çalıştığı varsayılarak; <b>hangi pazara (yurtiçi/ihracat)</b>, <b>hangi müşteriye</b> ve <b>hangi özelliklerde</b> ürün üretileceği hesaplanır.</p>", unsafe_allow_html=True)
            
    sim_days = st.number_input("Kaç Günlük Üretim Hesaplansın?", min_value=0.1, max_value=365.0, value=1.0, step=0.5, key="sim_days_input")
            
    if sim_days > 0:
        sim_kg = 0.0
        sim_adet = 0.0
                
        # Dictionary for aggregations
        simulated_lots = []
        stats_bolge = {}
        stats_musteri = {}
        stats_cap = {}
        stats_kalite = {}
        stats_kaplama = {}
                
        def add_stats(bolge, musteri, cap, kalite, kaplama, kg, adet):
            import pandas as pd
            for d, k in [(stats_bolge, bolge), (stats_musteri, musteri), (stats_cap, cap), (stats_kalite, kalite), (stats_kaplama, kaplama)]:
                key = str(k) if pd.notna(k) and str(k).strip() != "" else "Bilinmiyor"
                if key not in d:
                    d[key] = {"KG": 0.0, "Adet": 0.0}
                d[key]["KG"] += kg
                d[key]["Adet"] += adet
        if not os.path.exists(DB_PATH):
            st.info('Lütfen önce veritabanını oluşturun.')
        else:

            from database import get_db_connection
            import pandas as pd
            conn = get_db_connection()
            query = "SELECT * FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' AND MAKİNE IS NOT NULL AND MAKİNE != '' AND MAKİNE != 'nan'"
            df_sim = pd.read_sql_query(query, conn)
            conn.close()
                
            cols = df_sim.columns.tolist()
            x_col = next((c for c in cols if 'X G' in c.upper() and 'L' in c.upper() and 'K' in c.upper()), 'X GÜNLÜK İŞ')
            kg_col = next((c for c in cols if 'NET TONAJ' in c.upper()), 'NET TONAJ')
            adet_col = next((c for c in cols if 'KALAN ADET' in c.upper()), 'KALAN ADET')
                
            machines_active = df_sim['MAKİNE'].unique()
                
            for m in machines_active:
                m_df = df_sim[df_sim['MAKİNE'] == m].copy()
                m_df['sıra_num'] = pd.to_numeric(m_df['sıra'], errors='coerce').fillna(999999)
                m_df = m_df.sort_values(by='sıra_num', ascending=True).drop(columns=['sıra_num'])
                rem_days = float(sim_days)
                for _, row in m_df.iterrows():
                    if rem_days <= 0:
                        break
                    lot_days = pd.to_numeric(row.get(x_col, 0), errors='coerce')
                    if pd.isna(lot_days) or lot_days <= 0:
                        continue
                        
                    l_kg = pd.to_numeric(row.get(kg_col, 0), errors='coerce')
                    l_adet = pd.to_numeric(row.get(adet_col, 0), errors='coerce')
                        
                    if pd.isna(l_kg): l_kg = 0
                    if pd.isna(l_adet): l_adet = 0
                        
                    b = row.get("BÖLGE", "Bilinmiyor")
                    mus = row.get("MÜŞTERİ", "Bilinmiyor")
                    c = row.get("ÇAP", "Bilinmiyor")
                    kal = row.get("KALİTE", "Bilinmiyor")
                    kap = row.get("KAPLAMATIPI", "Bilinmiyor")
                        
                    if lot_days <= rem_days:
                        sim_kg += l_kg
                        sim_adet += l_adet
                        add_stats(b, mus, c, kal, kap, l_kg, l_adet)
                        simulated_lots.append({
                            "LOT NO": row.get("LOT", ""),
                            "MAKİNE": m,
                            "MÜŞTERİ": mus,
                            "STANDART": row.get("STANDART", ""),
                            "HAMMADDE": row.get("HAMMADDE", ""),
                            "AÇIKLAMA": row.get("AÇIKLAMA", ""),
                            "ÇAP": c,
                            "KALİTE": kal,
                            "KAPLAMATIPI": kap,
                            "KG": l_kg,
                            "ADET": l_adet,
                            "DURUM": "TAMAMLANDI"
                        })
                        rem_days -= lot_days
                    else:
                        frac = rem_days / lot_days
                        sim_kg += l_kg * frac
                        sim_adet += l_adet * frac
                        add_stats(b, mus, c, kal, kap, l_kg * frac, l_adet * frac)
                        simulated_lots.append({
                            "LOT NO": row.get("LOT", ""),
                            "MAKİNE": m,
                            "MÜŞTERİ": mus,
                            "STANDART": row.get("STANDART", ""),
                            "HAMMADDE": row.get("HAMMADDE", ""),
                            "AÇIKLAMA": row.get("AÇIKLAMA", ""),
                            "ÇAP": c,
                            "KALİTE": kal,
                            "KAPLAMATIPI": kap,
                            "KG": l_kg * frac,
                            "ADET": l_adet * frac,
                            "DURUM": f"KISMİ (%{(frac*100):.1f})"
                        })
                        rem_days = 0
                
            st.markdown("<hr>", unsafe_allow_html=True)
            col_sim1, col_sim2 = st.columns(2)
            with col_sim1:
                st.info(f"⚖️ **{sim_days} Günlük Tahmini Ağırlık:** {sim_kg:,.1f} KG ({(sim_kg/1000):,.1f} Ton)")
            with col_sim2:
                st.success(f"📦 **{sim_days} Günlük Tahmini Adet:** {sim_adet:,.0f} Adet")
                    
            import altair as alt
                
            def make_donut(d_dict, title):
                if not d_dict: return None
                df = pd.DataFrame([{"Kategori": k, "KG": v["KG"]} for k, v in d_dict.items()])
                df = df[df["KG"] > 0]
                if df.empty: return None
                    
                df = df.sort_values("KG", ascending=False)
                df["Yüzde"] = (df["KG"] / df["KG"].sum()) * 100
                df["Etiket"] = df["Kategori"] + " | " + df["KG"].map(lambda x: f"{x:,.0f} KG") + " (%" + df["Yüzde"].round(1).astype(str) + ")"
                    
                chart = alt.Chart(df).mark_arc(innerRadius=50).encode(
                    theta=alt.Theta(field="KG", type="quantitative"),
                    color=alt.Color(field="Etiket", type="nominal", sort=alt.EncodingSortField(field="KG", order="descending"), legend=alt.Legend(title=title, labelLimit=500)),
                    order=alt.Order(field="KG", type="quantitative", sort="descending"),
                    tooltip=["Kategori", alt.Tooltip("KG:Q", format=",.1f"), alt.Tooltip("Yüzde:Q", format=".1f")]
                ).properties(title=title, height=350)
                return chart
                    
            def make_bar(d_dict, title, top_n=10):
                if not d_dict: return None
                df = pd.DataFrame([{"Kategori": k, "KG": v["KG"]} for k, v in d_dict.items()])
                df = df.sort_values("KG", ascending=False).head(top_n)
                if df.empty: return None
                    
                chart = alt.Chart(df).mark_bar().encode(
                    x=alt.X("KG:Q", title="Tonaj (KG)"),
                    y=alt.Y("Kategori:N", sort="-x", title="", axis=alt.Axis(labelOverlap=False, labelLimit=250)),
                    color=alt.Color("Kategori:N", legend=None),
                    tooltip=["Kategori", alt.Tooltip("KG:Q", format=",.1f")]
                ).properties(title=title, height=450)
                return chart

            st.markdown("<h3 class='section-title'>📊 Kırılım Analizleri</h3>", unsafe_allow_html=True)
                
            st.markdown("#### 🌍 Yurtiçi / İhracat")
            c1, c2 = st.columns([2, 1])
            with c1:
                ch_bolge = make_donut(stats_bolge, "Bölge Bazlı Tonaj Dağılımı")
                if ch_bolge: st.altair_chart(ch_bolge, use_container_width=True)
            with c2:
                df_bolge = pd.DataFrame([{"Bölge": k, "Ağırlık (KG)": round(v["KG"], 1), "Adet": int(v["Adet"])} for k,v in stats_bolge.items()])
                st.dataframe(df_bolge, hide_index=True)
                    
            st.markdown("#### 🏢 Müşteri Dağılımı")
            ch_mus = make_bar(stats_musteri, "En Çok Üretim Yapılacak İlk 10 Müşteri")
            if ch_mus: st.altair_chart(ch_mus, use_container_width=True)
                
            st.markdown("#### ⚙️ Ürün Özellikleri")
            uc1, uc2, uc3 = st.columns(3)
            with uc1:
                ch_cap = make_donut(stats_cap, "Çap (Anma Çapı) Dağılımı")
                if ch_cap: st.altair_chart(ch_cap, use_container_width=True)
            with uc2:
                ch_kalite = make_donut(stats_kalite, "Kalite Dağılımı")
                if ch_kalite: st.altair_chart(ch_kalite, use_container_width=True)
            with uc3:
                ch_kaplama = make_donut(stats_kaplama, "Kaplama Tipi Dağılımı")
                if ch_kaplama: st.altair_chart(ch_kaplama, use_container_width=True)
            
            if simulated_lots:
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("#### 📋 Simülasyon Kapsamına Giren Ürünlerin Listesi", unsafe_allow_html=True)
                st.markdown("<p style='font-size:13px; color:#555;'>Aşağıdaki tablo, belirlediğiniz gün sayısı boyunca makinelerde işleneceği öngörülen lotların detaylı listesini gösterir.</p>", unsafe_allow_html=True)
                df_sim_list = pd.DataFrame(simulated_lots)
                df_sim_list['KG'] = df_sim_list['KG'].apply(lambda x: f"{x:,.1f}")
                df_sim_list['ADET'] = df_sim_list['ADET'].apply(lambda x: f"{x:,.0f}")
                st.dataframe(df_sim_list, use_container_width=True, hide_index=True)

with tab_ai_planning:
    st.markdown("<h2 class='section-title'>🤖 Otonom Kapasite Planlama (Yapay Zeka Destekli)</h2>", unsafe_allow_html=True)
    st.info("Bu modül, makine kapasite kullanımlarını ve belirlenen kıstasları dikkate alarak üretim kuyruğunu algoritmik olarak optimize etmek için tasarlanmaktadır. Kıstaslarınızı belirledikçe bu panel şekillenecektir.")
    
    st.markdown("### 🎯 Planlama Kıstasları ve Kısıtlar")
    col_ai1, col_ai2 = st.columns(2)
    with col_ai1:
        st.write("**Önceliklendirme Hedefleri**")
        opt_goal = st.selectbox("Optimizasyon Hedefi", ["Gecikmeleri Minimize Et (Termin Odaklı)", "Kalıp/Ayar Değişimini Minimize Et (Verimlilik Odaklı)", "Karma Hibrit Model (Dengeli)"])
        
        st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
        st.write("**Kural 1: Kesinleşmiş Plan Kısıtı (X Noktası)**")
        
        machine_list = []
        if os.path.exists(DB_PATH):
            conn_temp = migrate.sqlite3.connect(DB_PATH)
            machines_df = pd.read_sql("SELECT DISTINCT MAKİNE FROM programs WHERE MAKİNE IS NOT NULL AND MAKİNE != ''", conn_temp)
            conn_temp.close()
            machine_list = sorted(machines_df['MAKİNE'].tolist())
            machine_list = [m for m in machine_list if str(m).strip().upper() not in ['B', 'N', 'V'] and 'ÜRETİLEMİYOR' not in str(m).strip().upper()]
            
        if "locked_machines_key" not in st.session_state:
            st.session_state.locked_machines_key = machine_list
            
        col_b1, col_b2, _ = st.columns([2, 2, 5])
        if col_b1.button("🔒 Tümünü Kilitle", use_container_width=True):
            st.session_state.locked_machines_key = machine_list
            st.rerun()
        if col_b2.button("🔓 Tümünü Aç", use_container_width=True):
            st.session_state.locked_machines_key = []
            st.rerun()

        locked_machines = st.multiselect(
            "X İşaretine Kadar Planı Kilitlenecek Makineler",
            options=machine_list,
            key="locked_machines_key",
            help="Sadece burada seçili olan makinelerde 'X_CUTOFF' satırına kadar olan sıralama korunur. Seçilmeyen makineler baştan sona yeniden planlanır."
        )
        
        if len(locked_machines) > 0:
            st.info(f"Seçili {len(locked_machines)} makinede 'X' noktasına kadar olan sıralama korunacaktır.")
        else:
            st.warning("Dikkat: Kilit hiçbir makinede aktif değil. Algoritma tüm makinelerin kuyruklarını baştan sona yeniden dizecektir.")
            
    with col_ai2:
        st.write("**Kural 2: Satış Departmanı Acil Talepleri**")
        st.write("Satış ekibinin öncelikli üretilmesini talep ettiği LOT numaralarını buraya ekleyebilirsiniz. Eklediğiniz LOT'lar siz silene kadar burada kalır.")
        
        import json
        import os
        URGENT_LOTS_FILE = 'urgent_lots.json'
        
        # Load existing urgent lots
        if os.path.exists(URGENT_LOTS_FILE):
            with open(URGENT_LOTS_FILE, 'r', encoding='utf-8') as f:
                saved_urgent_lots = json.load(f)
        else:
            saved_urgent_lots = []
            
        # Add new lots
        with st.form("add_urgent_lot_form", clear_on_submit=True):
            new_lots_input = st.text_input("Yeni Öncelikli LOT Ekle (Birden fazlaysa virgülle ayırın)", placeholder="Örn: 2605366, 2605367")
            submitted = st.form_submit_button("➕ Ekle")
            if submitted and new_lots_input.strip():
                new_lots = [l.strip() for l in new_lots_input.split(',') if l.strip()]
                for nl in new_lots:
                    if nl not in saved_urgent_lots:
                        saved_urgent_lots.append(nl)
                with open(URGENT_LOTS_FILE, 'w', encoding='utf-8') as f:
                    json.dump(saved_urgent_lots, f)
                st.rerun()
                
        # Display and manage saved lots
        if saved_urgent_lots:
            st.write("📌 **Aktif Öncelikli LOT'lar:**")
            lots_to_remove = []
            
            # Query DB for statuses

            if not os.path.exists(DB_PATH):
                st.info('Lütfen önce veritabanını oluşturun.')
                st.stop()
            conn_temp2 = sqlite3.connect(DB_PATH)
            for lot in saved_urgent_lots:
                # Need to match lot strictly. Often lot has .0 or similar, but let's do a simple LIKE or exact match
                df_status = pd.read_sql(f"SELECT [SİPARİŞ DURUMU] FROM programs WHERE LOT LIKE '{lot}%'", conn_temp2)
                
                is_finished = False
                if not df_status.empty:
                    # If all matching records are BİTTİ (or similar completed status)
                    statuses = df_status['SİPARİŞ DURUMU'].astype(str).str.strip().str.upper().tolist()
                    if all(s == 'BİTTİ' for s in statuses):
                        is_finished = True
                
                col_lot1, col_lot2 = st.columns([4, 1])
                with col_lot1:
                    if is_finished:
                        st.success(f"✅ {lot} (Üretimi Bitti - Silebilirsiniz)")
                    else:
                        st.info(f"⏳ {lot} (Kuyrukta / Üretimde)")
                with col_lot2:
                    if st.button("🗑️ Sil", key=f"del_lot_{lot}"):
                        lots_to_remove.append(lot)
            conn_temp2.close()
            
            if lots_to_remove:
                for l in lots_to_remove:
                    saved_urgent_lots.remove(l)
                with open(URGENT_LOTS_FILE, 'w', encoding='utf-8') as f:
                    json.dump(saved_urgent_lots, f)
                st.rerun()
                
        urgent_lots_input = ",".join(saved_urgent_lots) # Pass to autonomous planner
        
        st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)
        st.write("**Kapasite Kısıtları**")
        st.write("*(Makinelerin çap, hız ve vardiya kapasiteleri veritabanından dinamik alınacaktır)*")
        
    st.markdown("### ⚙️ Optimizasyon Kontrolü")
    target_machine_options = ["Tüm Makineler"] + machine_list
    target_machine_selection = st.selectbox("Hangi Makine İçin Otonom Planlama Yapılsın?", options=target_machine_options, help="Sadece seçili makinenin otonom sıralaması (ai_sıra) baştan hesaplanır. Diğer makineler etkilenmez.")
    
    if st.button("🚀 Otonom Planlamayı Başlat", type="primary", key="btn_run_ai"):
        with st.spinner("Yapay Zeka planlama motoru çalışıyor..."):
            import sys
            import importlib
            if 'ai_planner' in sys.modules:
                importlib.reload(sys.modules['ai_planner'])
            from ai_planner import run_autonomous_planning
            success, msg = run_autonomous_planning(DB_PATH, locked_machines, urgent_lots_input, target_machine=target_machine_selection)
            if success:
                st.success(msg + " Otonom sonuçlar orijinal planınızı HİÇ BOZMADI. '⚖️ AI Karşılaştırma' sekmesine geçerek eski ve yeni halini yan yana görebilirsiniz.")
            else:
                st.error(msg)

# ==========================================
# TAB 3.5: AI KARŞILAŞTIRMA
# ==========================================
with tab_ai_compare:
    from datetime import datetime
    st.header("⚖️ Manuel vs Otonom Plan Karşılaştırması & Analiz")
    if os.path.exists(DB_PATH):
        try:
            conn = migrate.sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute("PRAGMA table_info(programs)")
            cols = [c[1] for c in cursor.fetchall()]
            
            if "ai_sıra" not in cols:
                st.info("Henüz hiçbir otonom planlama çalıştırılmamış. Lütfen 'Otonom Kapasite Planlama' sekmesinden bir planlama başlatın.")
            else:
                # 1. Fetch data
                query = "SELECT LOT, MAKİNE, [sıra], ai_sıra, [SİPARİŞ DURUMU], STANDART, ÇAP, BOY, [TERMİN TARİHİ], X_CUTOFF, [X GÜNLÜK İŞ], [KALAN ADET], [NET TONAJ], BÖLGE, MÜŞTERİ FROM programs WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ'"
                df_all = pd.read_sql(query, conn)
                
                # Convert numeric columns
                df_all['X GÜNLÜK İŞ'] = pd.to_numeric(df_all['X GÜNLÜK İŞ'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
                df_all['KALAN ADET'] = pd.to_numeric(df_all['KALAN ADET'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                df_all['NET TONAJ'] = pd.to_numeric(df_all['NET TONAJ'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
                
                # Process Termin Date
                df_all['Termin_Tarihi_DT'] = pd.to_datetime(df_all['TERMİN TARİHİ'], format='%d.%m.%Y', errors='coerce')
                
                start_date = pd.Timestamp(datetime.now().date())
                
                # 2. Calculate Cumulative Finish Dates per Machine
                def calc_finish_dates(df, sort_col):
                    df_sorted = df.copy()
                    if sort_col in df_sorted.columns:
                        df_sorted[f'{sort_col}_num'] = pd.to_numeric(df_sorted[sort_col], errors='coerce').fillna(999999)
                        df_sorted = df_sorted.sort_values(f'{sort_col}_num')
                        df_sorted = df_sorted.drop(columns=[f'{sort_col}_num'])
                    
                    df_sorted['Cum_Days'] = df_sorted['X GÜNLÜK İŞ'].cumsum()
                    df_sorted['Tahmini Bitiş'] = pd.to_datetime(df_sorted['Cum_Days'].fillna(0).astype(int).apply(lambda x: start_date + pd.offsets.BDay(x)))
                    df_sorted['Gecikme'] = df_sorted['Tahmini Bitiş'] > df_sorted['Termin_Tarihi_DT']
                    return df_sorted

                # Process all machines
                df_manual_all = pd.DataFrame()
                df_ai_all = pd.DataFrame()
                
                machines = df_all['MAKİNE'].dropna().unique()
                machines = [m for m in machines if str(m).strip().upper() not in ['B', 'N', 'V'] and 'ÜRETİLEMİYOR' not in str(m).strip().upper()]
                for mac in machines:
                    df_m = df_all[df_all['MAKİNE'] == mac]
                    if not df_m.empty:
                        # Manual
                        df_m_man = calc_finish_dates(df_m[df_m['sıra'].notna()], 'sıra')
                        df_manual_all = pd.concat([df_manual_all, df_m_man])
                        
                        # AI
                        df_m_ai = df_m[df_m['ai_sıra'].notna()]
                        if not df_m_ai.empty:
                            df_m_ai_calc = calc_finish_dates(df_m_ai, 'ai_sıra')
                            df_ai_all = pd.concat([df_ai_all, df_m_ai_calc])
                            
                # 3. Overall Dashboard
                st.markdown("### 🏭 Fabrika Geneli Performans (Tüm Makineler)")
                
                if not df_ai_all.empty:
                    # Gecikme KPI
                    man_gecikme_tonaj = df_manual_all[df_manual_all['Gecikme']]['NET TONAJ'].sum()
                    ai_gecikme_tonaj = df_ai_all[df_ai_all['Gecikme']]['NET TONAJ'].sum()
                    man_gecikme_adet = df_manual_all[df_manual_all['Gecikme']]['KALAN ADET'].sum()
                    ai_gecikme_adet = df_ai_all[df_ai_all['Gecikme']]['KALAN ADET'].sum()
                    
                    kpi1, kpi2 = st.columns(2)
                    with kpi1:
                        st.metric("Geciken Toplam Net Tonaj (AI Önerisi)", f"{ai_gecikme_tonaj:,.0f} kg", f"{ai_gecikme_tonaj - man_gecikme_tonaj:,.0f} kg (Eski Plana Göre Fark)", delta_color="inverse")
                    with kpi2:
                        st.metric("Geciken Toplam Adet (AI Önerisi)", f"{ai_gecikme_adet:,.0f} adet", f"{ai_gecikme_adet - man_gecikme_adet:,.0f} adet (Eski Plana Göre Fark)", delta_color="inverse")
                        
                    # İhracat/Yurtiçi by Month
                    st.markdown("#### 🌍 Bölge ve Müşteri Analizi (Otonom Plan)")
                    col_b, col_m = st.columns(2)
                    
                    # Formatting month for grouping
                    df_ai_all['Ay'] = df_ai_all['Tahmini Bitiş'].dt.strftime('%Y-%m')
                    
                    import altair as alt
                    with col_b:
                        st.markdown("**Bölgeye Göre Aylık Tonaj (kg)**")
                        b_group = df_ai_all.groupby(['Ay', 'BÖLGE'])['NET TONAJ'].sum().reset_index()
                        if not b_group.empty:
                            b_group['NET TONAJ'] = b_group['NET TONAJ'].astype(int)
                            b_chart = alt.Chart(b_group).mark_bar().encode(
                                x=alt.X('Ay:N', title=None),
                                xOffset='BÖLGE:N',
                                y=alt.Y('NET TONAJ:Q', title=None, axis=None),
                                color=alt.Color('BÖLGE:N', legend=alt.Legend(title="Bölge", orient="bottom"))
                            )
                            b_text = b_chart.mark_text(
                                align='center',
                                baseline='bottom',
                                dy=-5,
                                fontSize=11,
                            ).encode(
                                text=alt.Text('NET TONAJ:Q', format=',d')
                            )
                            st.altair_chart((b_chart + b_text).configure_view(stroke='transparent'), use_container_width=True)
                        
                    with col_m:
                        st.markdown("**En Çok Üretim Yapılacak Müşteriler (İlk 5)**")
                        c_group = df_ai_all.groupby('MÜŞTERİ')['NET TONAJ'].sum().sort_values(ascending=False).head(5).reset_index()
                        if not c_group.empty:
                            c_group['NET TONAJ'] = c_group['NET TONAJ'].astype(int)
                            c_chart = alt.Chart(c_group).mark_bar().encode(
                                x=alt.X('MÜŞTERİ:N', sort='-y', title=None, axis=alt.Axis(labelAngle=-45)),
                                y=alt.Y('NET TONAJ:Q', title=None, axis=None),
                                color=alt.Color('MÜŞTERİ:N', legend=None)
                            )
                            c_text = c_chart.mark_text(
                                align='center',
                                baseline='bottom',
                                dy=-5,
                                fontSize=11,
                            ).encode(text=alt.Text('NET TONAJ:Q', format=',d'))
                            st.altair_chart((c_chart + c_text).configure_view(stroke='transparent'), use_container_width=True)
                        
                st.markdown("<hr>", unsafe_allow_html=True)
                
                # 4. Machine Level View
                st.markdown("### ⚙️ Makine Bazında Kuyruk ve Gecikme İncelemesi")
                
                if len(machines) > 0:
                    # Prepare dynamic pill options with delay reduction indicators
                    pill_options = []
                    mac_mapping = {} # To map the pill text back to the actual machine name
                    
                    if not df_manual_all.empty and not df_ai_all.empty:
                        for m in sorted(machines):
                            man_delay = df_manual_all[(df_manual_all['MAKİNE'] == m) & df_manual_all['Gecikme']]['NET TONAJ'].sum()
                            ai_delay = df_ai_all[(df_ai_all['MAKİNE'] == m) & df_ai_all['Gecikme']]['NET TONAJ'].sum()
                            diff = man_delay - ai_delay
                            
                            if diff > 0:
                                label = f"🟢 {m} (↓ {diff/1000:.1f} Ton)"
                            elif diff < 0:
                                label = f"🔴 {m} (↑ {abs(diff)/1000:.1f} Ton)"
                            else:
                                label = f"⚪ {m}"
                                
                            pill_options.append(label)
                            mac_mapping[label] = m
                    else:
                        pill_options = sorted(machines)
                        mac_mapping = {m: m for m in machines}
                        
                    selected_pill = st.pills("İncelemek istediğiniz makineyi seçin:", options=pill_options, default=pill_options[0] if len(pill_options)>0 else None, key="ai_comp_mac")
                    selected_mac = mac_mapping.get(selected_pill) if selected_pill else None
                    
                    
                    if not df_ai_all.empty:
                        df_manual = df_manual_all[df_manual_all['MAKİNE'] == selected_mac].reset_index(drop=True)
                        df_ai = df_ai_all[df_ai_all['MAKİNE'] == selected_mac].reset_index(drop=True)
                        
                        if df_ai.empty:
                            st.warning("Bu makine için yapay zeka sonucu bulunamadı. Lütfen planlamayı tekrar başlatın.")
                        else:
                            # Machine-Specific Dashboard
                            m_man_gecikme = df_manual[df_manual['Gecikme']]['NET TONAJ'].sum()
                            m_ai_gecikme = df_ai[df_ai['Gecikme']]['NET TONAJ'].sum()
                            
                            st.markdown(f"#### {selected_mac} Makinesi Performans Farkı")
                            kpi1, kpi2 = st.columns(2)
                            with kpi1:
                                st.metric("Bu Makinede Geciken Tonaj (AI)", f"{m_ai_gecikme:,.0f} kg", f"{m_ai_gecikme - m_man_gecikme:,.0f} kg (Eski Plana Göre Fark)", delta_color="inverse")
                            
                            # Format Date for display
                            df_manual['Tahmini Bitiş'] = df_manual['Tahmini Bitiş'].dt.strftime('%d.%m.%Y')
                            df_ai['Tahmini Bitiş'] = df_ai['Tahmini Bitiş'].dt.strftime('%d.%m.%Y')
                            
                            def style_manual(row):
                                cutoff = df_manual[df_manual['X_CUTOFF'] == 1]
                                cutoff_idx = cutoff.index[0] if not cutoff.empty else -1
                                
                                style = [''] * len(row)
                                if cutoff_idx >= 0 and row.name <= cutoff_idx:
                                    style = ['background-color: #333333; color: white;'] * len(row)
                                
                                lot_val = str(row.get('LOT', '')).split('.')[0].strip()
                                if lot_val in saved_urgent_lots:
                                    style = ['background-color: #9b59b6; color: white; font-weight: bold;'] * len(row)
                                elif row.get('Gecikme', False):
                                    if not (cutoff_idx >= 0 and row.name <= cutoff_idx):
                                        style = ['color: #dc3545; font-weight: bold;'] * len(row)
                                return style

                            def style_ai(row):
                                original_row = df_ai.loc[row.name]
                                cutoff = df_ai[df_ai['X_CUTOFF'] == 1]
                                cutoff_idx = cutoff.index[0] if not cutoff.empty else -1
                                
                                style = [''] * len(row)
                                if cutoff_idx >= 0 and row.name <= cutoff_idx:
                                    style = ['background-color: #333333; color: white;'] * len(row)
                                else:
                                    sira = original_row['sıra']
                                    ai_sira = original_row['ai_sıra']
                                    
                                    if pd.notna(sira) and pd.notna(ai_sira):
                                        try:
                                            sira_val = float(sira)
                                            ai_sira_val = float(ai_sira)
                                            if ai_sira_val < sira_val:
                                                style = ['background-color: rgba(40, 167, 69, 0.3); color: black;'] * len(row)
                                            elif ai_sira_val > sira_val:
                                                style = ['background-color: rgba(220, 53, 69, 0.3); color: black;'] * len(row)
                                        except (ValueError, TypeError):
                                            pass
                                            
                                lot_val = str(row.get('LOT', '')).split('.')[0].strip()
                                if lot_val in saved_urgent_lots:
                                    style = ['background-color: #9b59b6; color: white; font-weight: bold;'] * len(row)
                                elif row.get('Gecikme', False):
                                    style = [s + ' color: #dc3545; font-weight: bold;' if not s else s for s in style]
                                return style
                                
                            col1, col2 = st.columns(2)
                            display_cols = ['LOT', 'MÜŞTERİ', 'STANDART', 'ÇAP', 'BOY', 'TERMİN TARİHİ', 'Tahmini Bitiş', 'Gecikme']
                            with col1:
                                st.subheader("👨‍💻 Sizin Planınız (Manuel)")
                                st.dataframe(df_manual[display_cols].style.apply(style_manual, axis=1), use_container_width=True, hide_index=True)
                            with col2:
                                st.subheader("🤖 Yapay Zeka (Öneri)")
                                st.dataframe(df_ai[display_cols].style.apply(style_ai, axis=1), use_container_width=True, hide_index=True)
                else:
                    st.info("Aktif sipariş bulunan makine yok.")
            conn.close()
        except Exception as e:
            st.error(f"Hata oluştu: {e}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("Veritabanı yok.")

with tab_planning:
    if os.path.exists(DB_PATH):
        machines = get_machines()
        
        if "selected_machine" not in st.session_state:
            st.session_state.selected_machine = machines[0] if machines else None
            
        st.markdown("<h4 style='color:#1f4e79;'>⚙️ İncelemek ve Planlamak İstediğiniz Makineyi Seçin</h4>", unsafe_allow_html=True)
        
        b_machines = [m for m in machines if m.startswith('B')]
        v_machines = [m for m in machines if m.startswith('V')]
        n_machines = [m for m in machines if m.startswith('N')]
        other_machines = [m for m in machines if not (m.startswith('B') or m.startswith('V') or m.startswith('N'))]
        
        st.markdown("""
        <style>
        /* Butonları tek satıra sığdırmak için ekstra minimal CSS */
        .stButton button {
            padding: 0px 0px !important;
            min-height: 24px !important;
            height: 26px !important;
        }
        .stButton button p {
            font-size: 10px !important;
            font-weight: 700 !important;
            margin: 0 !important;
        }
        </style>
        """, unsafe_allow_html=True)
        
        def render_machine_buttons(mach_list, label):
            if not mach_list: return
            st.markdown(f"<p style='margin-bottom:2px; margin-top:8px; font-size:13px; font-weight:700; color:#1f4e79;'>{label}</p>", unsafe_allow_html=True)
            cols = st.columns(len(mach_list))
            for j, m in enumerate(mach_list):
                with cols[j]:
                    is_active = (st.session_state.get('selected_machine') == m)
                    if st.button(m, key=f"btn_mach_{m}", use_container_width=True, type="primary" if is_active else "secondary"):
                        st.session_state.selected_machine = m
                        st.rerun()
                            
        with st.container(border=True):
            render_machine_buttons(b_machines, "🟦 B Serisi Makineler")
            render_machine_buttons(v_machines, "🟩 V Serisi Makineler")
            render_machine_buttons(n_machines, "🟧 N Serisi Makineler")
            render_machine_buttons(other_machines, "⬛ Diğer Makineler")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        selected_machine = st.session_state.get('selected_machine')
        
        col_m_select, col_k_adet, col_k_kg, col_x_gun = st.columns([1.5, 1, 1, 1])
        with col_m_select:
            if selected_machine:
                st.markdown(f"<h3 class='section-title' style='margin-top:0;'>{selected_machine} Makinesi Aktif Kuyruğu</h3>", unsafe_allow_html=True)
            
        if selected_machine:
            
            # Load active queue
            queue_df = get_active_queue(selected_machine)
            
            if not queue_df.empty:
                # Find X Cutoff
                cutoff_idx = -1
                if 'X_CUTOFF' in queue_df.columns:
                    x_rows = queue_df[queue_df['X_CUTOFF'] == 1]
                    if not x_rows.empty:
                        cutoff_idx = x_rows.index[0]
                        
                sum_k_adet = 0
                sum_k_kg = 0
                sum_x_gun = 0
                
                if cutoff_idx >= 0:
                    cutoff_df = queue_df.iloc[:cutoff_idx+1]
                    sum_k_adet = pd.to_numeric(cutoff_df['KALAN ADET'], errors='coerce').sum()
                    sum_k_kg = pd.to_numeric(cutoff_df['NET TONAJ'], errors='coerce').sum()
                    sum_x_gun = pd.to_numeric(cutoff_df['X GÜNLÜK İŞ'], errors='coerce').sum()
                    
                with col_k_adet:
                    if cutoff_idx >= 0:
                        st.metric("K. ADET (Kesme Noktası)", f"{sum_k_adet:,.0f}")
                with col_k_kg:
                    if cutoff_idx >= 0:
                        st.metric("K.KG (Kesme Noktası)", f"{sum_k_kg:,.0f}")
                with col_x_gun:
                    if cutoff_idx >= 0:
                        st.metric("X GÜNLÜK İŞ (Kesme Noktası)", f"{sum_x_gun:,.1f}")

                # Define Excel column mapping (pull values from SQLite table columns)
                COLUMN_MAPPING = {
                    "LOT NO": "LOT",
                    "SİP.TARİH": "SİP.TARİH",
                    "MÜŞTERİ": "MÜŞTERİ",
                    "ÇAP": "ÇAP",
                    "BOY": "BOY",
                    "DİŞADIM": "DİŞADIM",
                    "STANDART": "STANDART",
                    "KALİTE": "KALİTE",
                    "MARKA": "MARKA",
                    "PASO": "PASO",
                    "NET GRAM": "NET GRAM",
                    "BR.GR.": "BRÜT GRAM",
                    "ADET": "ADET",
                    "KG": "KG",
                    "TERMİN": "TERMİN TARİHİ",
                    "HAMMADDE": "HAMMADDE",
                    "MATERIAL": "MATERIAL",
                    "TEMELURUN": "TEMELURUN",
                    "AÇIKLAMA": "AÇIKLAMA",
                    "KAPLAMATIPI": "KAPLAMATIPI",
                    "KAPLAMASTANDART": "KAPLAMASTANDART",
                    "DOCTYPE": "DOCTYPE",
                    "DOCNUM": "DOCNUM",
                    "ITEMNUM": "ITEMNUM",
                    "CUSTOMER": "CUSTOMER",
                    "NOT": "NOT",
                    "DRAWNUM": "DRAWNUM",
                    "SİVRİLTME": "SİVRİLTME",
                    "DURUM": "SİPARİŞ DURUMU",
                    "MAKİNA": "MAKİNE",
                    "K. ADET": "KALAN ADET",
                    "K.KG": "NET TONAJ",
                    "X GÜNLÜK İŞ": "X GÜNLÜK İŞ",
                    "ÜGGT": "ÜGGT",
                    "ÜBKZ": "ÜBKZ"
                }
                
                # Construct display dataframe with exact Excel columns and order
                df_display = pd.DataFrame()
                for target_col, src_col in COLUMN_MAPPING.items():
                    if src_col in queue_df.columns:
                        df_display[target_col] = queue_df[src_col]
                    else:
                        df_display[target_col] = None
                
                # Render styled table at the very top, full width
                def highlight_cutoff(row):
                    style = [''] * len(row)
                    if cutoff_idx >= 0 and row.name <= cutoff_idx:
                        style = ['background-color: #333333; color: white;'] * len(row)
                        
                    lot_val = str(row.get('LOT', '')).split('.')[0].strip()
                    if lot_val in saved_urgent_lots:
                        style = ['background-color: #9b59b6; color: white; font-weight: bold;'] * len(row)
                        
                    return style
                    
                styled_df = df_display.style.apply(highlight_cutoff, axis=1)
                st.dataframe(styled_df, use_container_width=True, hide_index=True)
                
                st.markdown("<hr style='margin:20px 0;'>", unsafe_allow_html=True)
                
                # Render controls and drag-and-drop section side-by-side below the table
                if not is_cloud:
                    col_drag, col_actions = st.columns([2, 1])
                    
                    with col_drag:
                        with st.expander("↕️ Sürükle-Bırak ile Sıralamayı Değiştir (Drag & Drop)", expanded=True):
                            st.markdown("<p style='font-size:0.9rem; color:#666;'>Ögeleri sürükleyip bırakarak yeni üretim sırasını belirleyin ve ardından aşağıdaki kaydet butonuna basın.</p>", unsafe_allow_html=True)
                            
                            # Format items for draggable list
                            sort_list = []
                            for idx, (_, row) in enumerate(queue_df.iterrows(), 1):
                                sort_list.append(f"📦 Sıra {idx}: LOT {row['LOT']} - {row.get('STANDART', '')} ({row.get('ÇAP', '')}x{row.get('BOY', '')}) - {row.get('MÜŞTERİ', '')}")
                            
                            # Render sortables widget
                            import hashlib
                            list_hash = hashlib.md5(str(sort_list).encode('utf-8')).hexdigest()
                            sorted_items = sort_items(sort_list, direction="vertical", key=f"sort_{selected_machine}_{list_hash}")
                            
                            if st.button("💾 Yeni Sıralamayı Kaydet", use_container_width=True, type="primary"):
                                try:
                                    # Extract lot numbers in the new order
                                    new_lots = []
                                    for item in sorted_items:
                                        parts = item.split("LOT ")
                                        lot_no = parts[1].split(" - ")[0].strip()
                                        new_lots.append(lot_no)
                                        
                                    # Save to database
                                    update_lot_sequence(new_lots)
                                    st.session_state.modified_machines.add(selected_machine)
                                    st.success("Yeni üretim sıralaması veritabanına başarıyla kaydedildi!")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Sıralama kaydedilirken hata oluştu: {e}")
                    
                    with col_actions:
                        st.markdown("<div style='background-color:#f8f9fa; padding:15px; border-radius:8px; border:1px solid #ddd;'>", unsafe_allow_html=True)
                        st.subheader("Planı Düzenle")
                        
                        # Target lot selection
                        lot_list = queue_df["LOT"].tolist()
                        target_lot = st.selectbox("İşlem Yapılacak Lot Seçin:", lot_list)
                        
                        # Ordering Buttons
                        idx = lot_list.index(target_lot) if target_lot in lot_list else -1
                        
                        btn_up, btn_down = st.columns(2)
                        with btn_up:
                            if st.button("⬆️ Yukarı Taşı", use_container_width=True) and idx > 0:
                                lot_list[idx], lot_list[idx-1] = lot_list[idx-1], lot_list[idx]
                                update_lot_sequence(lot_list)
                                st.session_state.modified_machines.add(selected_machine)
                                st.success(f"{target_lot} sırası yukarı alındı.")
                                st.rerun()
                                
                        with btn_down:
                            if st.button("⬇️ Aşağı Taşı", use_container_width=True) and idx < len(lot_list) - 1:
                                lot_list[idx], lot_list[idx+1] = lot_list[idx+1], lot_list[idx]
                                update_lot_sequence(lot_list)
                                st.session_state.modified_machines.add(selected_machine)
                                st.success(f"{target_lot} sırası aşağı alındı.")
                                st.rerun()
                                
                        if st.button("🔝 En Üste Taşı", use_container_width=True) and idx > 0:
                            lot_list.insert(0, lot_list.pop(idx))
                            update_lot_sequence(lot_list)
                            st.session_state.modified_machines.add(selected_machine)
                            st.success(f"{target_lot} en üste taşındı.")
                            st.rerun()
                            
                        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
                        
                        # Status / Machine Change ACTIONS
                        new_status = st.selectbox("Durum Değiştir:", ["BİTMEDİ", "BİTTİ"])
                        if st.button("Durumu Güncelle", use_container_width=True):
                            update_lot_fields(target_lot, {"SİPARİŞ DURUMU": new_status})
                            st.session_state.modified_machines.add(selected_machine)
                            st.success(f"{target_lot} durumu '{new_status}' olarak güncellendi.")
                            st.rerun()
                            
                        new_machine = st.selectbox("Makinayı Değiştir:", machines, index=machines.index(selected_machine))
                        if st.button("Makineye Gönder", use_container_width=True) and new_machine != selected_machine:
                            # Append to the end of the target machine queue
                            target_queue = get_active_queue(new_machine)
                            new_sira = float(len(target_queue) + 1)
                            update_lot_fields(target_lot, {"MAKİNE": new_machine, "sıra": new_sira})
                            st.session_state.modified_machines.add(selected_machine)
                            st.session_state.modified_machines.add(new_machine)
                            st.success(f"{target_lot} Loti {new_machine} makinesine aktarıldı.")
                            st.rerun()
                            
                        st.markdown("</div>", unsafe_allow_html=True)

            else:
                st.info(f"{selected_machine} makinesine atanmış aktif lot bulunmamaktadır.")
    else:
        st.info("Lütfen veritabanını oluşturun.")

# ==========================================
# TAB: GENEL MAKINE AYARLARI
# ==========================================
with tab_genel_makine:
    if os.path.exists(DB_PATH):
        st.markdown("<h3 class='section-title'>⚙️ Genel Makine Parametreleri</h3>", unsafe_allow_html=True)
        st.info("Bu ekrandan genel makine parametrelerini (hedeflenen verimlilik, çalışma süreleri vb.) düzenleyebilirsiniz. Değiştirmek istediğiniz hücreye çift tıklayın.")
        
        df_genel = get_genel_makine()
        
        if not df_genel.empty:
            # "2021-2025 ORTALAMALARI" ve sonrasındaki gereksiz tabloyu kesip at
            first_col = df_genel.columns[0]
            ort_idx = df_genel.index[df_genel[first_col].astype(str).str.contains("2021", na=False)].tolist()
            if ort_idx:
                df_genel = df_genel.loc[:ort_idx[0]-1]
            
            # Tamamen boş satırları (MAKİNE adı boş olanları) temizle
            df_genel = df_genel.dropna(subset=[first_col]).reset_index(drop=True)
            
            import pandas as pd
            st.markdown("### 📊 Otomatik Üretim ve Makine Analizi")
            
            # Dinamik Sütun Tespiti (Boşluk vb. hatalarından korunmak için)
            geciken_col = next((c for c in df_genel.columns if "GECİKEN KG" in str(c).upper()), None)
            verim_col = next((c for c in df_genel.columns if "HEDEFLENEN" in str(c).upper() and "VER" in str(c).upper()), None)
            is_gunu_col = next((c for c in df_genel.columns if "GÜNÜ" in str(c).upper()), None)
            makine_col = next((c for c in df_genel.columns if c.upper().strip() in ["MAKİNE", "MAKINE", "BMAKİNE"]), "MAKİNE")
            
            if geciken_col and verim_col:
                geciken_kg = pd.to_numeric(df_genel[geciken_col], errors='coerce').fillna(0)
                hedef_verim = pd.to_numeric(df_genel[verim_col], errors='coerce').fillna(0)
                
                toplam_geciken_kg = geciken_kg.sum()
                ortalama_verim = hedef_verim[hedef_verim > 0].mean() if len(hedef_verim[hedef_verim > 0]) > 0 else 0
                geciken_makineler = df_genel[geciken_kg > 0][makine_col].dropna().astype(str).tolist()
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Genel Ortalama Verimlilik", f"%{ortalama_verim:.1f}")
                col2.metric("Toplam Geciken Üretim", f"{toplam_geciken_kg:,.0f} KG", delta=f"{len(geciken_makineler)} Makinede" if len(geciken_makineler)>0 else "Yok", delta_color="inverse")
                col3.metric("Kritik Makine Sayısı", f"{len(geciken_makineler)}")
                
                with st.expander("💡 Sistem Yorumu ve Kapasite Önerileri", expanded=True):
                    if len(geciken_makineler) > 0:
                        st.error(f"⚠️ **Gecikme Uyarısı:** {', '.join(geciken_makineler)} makinesinde/makinelerinde toplam **{toplam_geciken_kg:,.0f} KG** geciken üretim görünüyor. Darboğazı (bottleneck) aşmak için bu makinelerde mesai saatlerini (Çalışma DK) artırmayı veya 'Hedeflenen Verimlilik' bazlı bakım yapmayı değerlendirebilirsiniz.")
                    else:
                        st.success("✅ **Harika!** Şu anda sisteme yansıyan hiçbir geciken sipariş görünmüyor. Üretim planınız hedeflenen kapasite sınırları içinde tıkır tıkır işliyor.")
                    
                    if is_gunu_col:
                        is_gunu = pd.to_numeric(df_genel[is_gunu_col], errors='coerce').fillna(0)
                        temp_df = df_genel.copy()
                        temp_df["IS_GUNU_NUM"] = is_gunu
                        yogun_makineler = temp_df.sort_values(by="IS_GUNU_NUM", ascending=False)
                        if not yogun_makineler.empty and yogun_makineler.iloc[0]["IS_GUNU_NUM"] > 0:
                            top_m = yogun_makineler.iloc[0]
                            st.info(f"🔥 **En Yoğun Makine:** Mevcut iş yüküne göre bitirme süresi en uzun makineniz **{top_m[makine_col]}** (Yaklaşık **{top_m['IS_GUNU_NUM']:.1f} iş günü** sürecek). Eğer mümkünse, bu makinedeki bazı iş emirlerini kapasitesi boş olan muadil makinelere kaydırarak üretim sürecini hızlandırabilirsiniz.")

            st.markdown("---")
            
            formula_cols = [
                "SİPARİŞ ADEDİ", "ADET", "KG", "Ortalama Gramaj", "Ortalama lot büyüklü", 
                "%100 verim ile günlük üretim", "hedeflenen verim ile günlük üretim", 
                "GECİKEN ADET", "GECİKEN KG"
            ]
            
            bar_cols = [c for c in df_genel.columns if any(x in str(c).replace("İ", "I").replace("ı", "i").lower() for x in ["geciken", "iş günü", "is gunu"])]
            for bc in bar_cols:
                # Kesin olarak numeric/float yapalım ki ProgressColumn hata vermesin
                df_genel[bc] = pd.to_numeric(df_genel[bc], errors='coerce').fillna(0).astype(float)
                bc_lower = str(bc).replace("İ", "I").replace("ı", "i").lower()
                if "geciken" in bc_lower and "kg" in bc_lower:
                    df_genel[bc] = df_genel[bc] / 1000.0 # Görselde Ton olarak göster
                
            col_config = {}
            for col in df_genel.columns:
                if col in bar_cols:
                    max_v = float(df_genel[col].max())
                    if max_v <= 0:
                        max_v = 1.0 # Eğer hepsi 0 ise patlamaması için 1 veriyoruz
                        
                    col_lower = str(col).replace("İ", "I").replace("ı", "i").lower()
                    if "iş günü" in col_lower or "günlük iş" in col_lower or "is gunu" in col_lower:
                        fmt = "%.1f"
                        lbl = col
                    elif "geciken" in col_lower and "adet" in col_lower:
                        fmt = "%d"
                        lbl = col
                    elif "geciken" in col_lower and "kg" in col_lower:
                        fmt = "%d"
                        lbl = "GECİKEN (TON)"
                    else:
                        fmt = "%f"
                        lbl = col
                        
                    col_config[col] = st.column_config.ProgressColumn(
                        label=lbl, 
                        format=fmt,
                        min_value=0.0,
                        max_value=max_v
                    )
                else:
                    is_disabled = (col in formula_cols)
                    if pd.api.types.is_numeric_dtype(df_genel[col]):
                        col_config[col] = st.column_config.NumberColumn(col, disabled=is_disabled, format="%d")
                    else:
                        col_config[col] = st.column_config.Column(col, disabled=is_disabled)
                    
            def highlight_cells(row):
                styles = [''] * len(row)
                for i, col in enumerate(df_genel.columns):
                    col_str = str(col).replace("İ", "I").replace("ı", "i").lower()
                    if col in ["HEDEFLENEN \nVERİMLİLİK", "DEVİR/DK", "ÇALIŞMA DK"]:
                        styles[i] = 'background-color: #2b4b7c; color: white;' 
                    elif "günlük üretim" in col_str:
                        styles[i] = 'background-color: #2e603b; color: white;' 
                return styles
            
            styled_df = df_genel.style.apply(highlight_cells, axis=1)
            
            edited_df = st.data_editor(
                styled_df,
                use_container_width=True,
                hide_index=True,
                column_config=col_config,
                num_rows="fixed",
                key="genel_makine_editor"
            )
            
            if st.button("💾 Genel Makine Değişikliklerini Kaydet", type="primary", use_container_width=True):
                import pandas as pd
                
                # Geciken KG verisi görselleştirme için Ton'a dönüştürülmüştü,
                # Veritabanına kaydederken KG cinsine geri çeviriyoruz.
                gec_kg_cols = [c for c in edited_df.columns if "geciken" in str(c).replace("İ", "I").replace("ı", "i").lower() and "kg" in str(c).replace("İ", "I").replace("ı", "i").lower()]
                for c in gec_kg_cols:
                    edited_df[c] = pd.to_numeric(edited_df[c], errors="coerce").fillna(0) * 1000.0
                
                # Arayüzde anlık görmek için Python tarafında temel formül hesaplamaları
                if "DEVİR/DK" in edited_df.columns and "ÇALIŞMA DK" in edited_df.columns:
                    devir = pd.to_numeric(edited_df["DEVİR/DK"], errors="coerce").fillna(0)
                    calisma = pd.to_numeric(edited_df["ÇALIŞMA DK"], errors="coerce").fillna(0)
                    
                    if "%100 verim ile günlük üretim" in edited_df.columns:
                        edited_df["%100 verim ile günlük üretim"] = devir * calisma
                        
                    hedef_col = next((c for c in edited_df.columns if "HEDEFLENEN" in c.upper() and "VER" in c.upper()), None)
                    if hedef_col and "hedeflenen verim ile günlük üretim" in edited_df.columns:
                        hedef = pd.to_numeric(edited_df[hedef_col], errors="coerce").fillna(0)
                        edited_df["hedeflenen verim ile günlük üretim"] = (devir * calisma * hedef) / 100

                update_genel_makine(edited_df)
                st.session_state.modified_machines.add("GENEL MAKİNE")
                st.success("Değişiklikler veritabanına kaydedildi! Excel'e aktarmak için 'Excel Veri Göçü & Eşleme' sekmesini kullanın.")
                st.rerun()
        else:
            st.warning("GENEL MAKİNE tablosunda veri bulunamadı. Lütfen 'Excel Veri Göçü & Eşleme' sekmesinden aktarım işlemini tekrarlayın.")
    else:
        st.info("Lütfen veritabanını oluşturun.")

# ==========================================
# TAB 3: PRINT WORK ORDER
# ==========================================
with tab_work_order:
    if os.path.exists(DB_PATH):
        st.markdown("<h3 class='section-title'>Lot No ile İş Emri Kartı Arama & Basma</h3>", unsafe_allow_html=True)
        
        search_col, print_col = st.columns([1, 2])
        with search_col:
            search_lot = st.text_input("Aramak İstediğiniz LOT NO girin:")
            
        if search_lot:
            lot_details = get_lot_details(search_lot)
            
            if lot_details:
                with print_col:
                    st.success(f"LOT {search_lot} bulundu!")
                    
                    # Generate PDF Action
                    pdf_dir = os.path.join(os.path.dirname(__file__), "output_pdfs")
                    os.makedirs(pdf_dir, exist_ok=True)
                    pdf_filename = f"Is_Emri_{search_lot}.pdf"
                    pdf_path = os.path.join(pdf_dir, pdf_filename)
                    
                    # Render button
                    if st.button("📄 İş Emri PDF'i Üret"):
                        try:
                            generate_pdf(lot_details, pdf_path)
                            st.success(f"İş Emri PDF başarıyla üretildi!")
                            
                            # Download link
                            with open(pdf_path, "rb") as f:
                                st.download_button(
                                    label="📥 Üretilen PDF'i İndir",
                                    data=f,
                                    file_name=pdf_filename,
                                    mime="application/pdf",
                                    use_container_width=True
                                )
                        except Exception as e:
                            st.error(f"PDF üretilirken hata oluştu: {e}")
                            
                # Display Lot Data in structured cards
                st.markdown("<h4 style='color:#1f4e79; margin-top:15px;'>Lot Kart Detayları</h4>", unsafe_allow_html=True)
                
                # Layout layout
                card1, card2, card3 = st.columns(3)
                with card1:
                    st.info(f"**Temel Bilgiler**\n\n* **LOT:** {lot_details.get('LOT')}\n* **MAKİNE:** {lot_details.get('MAKİNE')}\n* **MÜŞTERİ:** {lot_details.get('MÜŞTERİ')}\n* **SİPARİŞ DURUMU:** {lot_details.get('SİPARİŞ DURUMU')}")
                with card2:
                    st.info(f"**Teknik Özellikler**\n\n* **STANDART:** {lot_details.get('STANDART')}\n* **KALİTE:** {lot_details.get('KALİTE')}\n* **ÇAP / BOY / ADIM:** {lot_details.get('ÇAP')} x {lot_details.get('BOY')} / {lot_details.get('DİŞADIM')}\n* **MARKA:** {lot_details.get('MARKA')}")
                with card3:
                    st.info(f"**Miktar & Hammadde**\n\n* **ADET / KG:** {lot_details.get('ADET')} / {lot_details.get('KG')} KG\n* **HAMMADDE:** {lot_details.get('HAMMADDE')}\n* **KAPLAMA:** {lot_details.get('KAPLAMATIPI')} ({lot_details.get('KAPLAMASTANDART')})")
                    
            else:
                st.error(f"LOT NO '{search_lot}' veritabanında bulunamadı. Lütfen doğru yazdığınızdan emin olun.")
    else:
        st.info("Lütfen veritabanını oluşturun.")

# ==========================================
# TAB 4: ADD NEW LOT
# ==========================================
with tab_add_lot:
    if os.path.exists(DB_PATH):
        st.markdown("<h3 class='section-title'>Sisteme Yeni Lot Kaydı Ekle</h3>", unsafe_allow_html=True)
        
        with st.form("new_lot_form"):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                lot = st.text_input("LOT NO *", help="Benzersiz lot numarası girin.")
                machine = st.selectbox("MAKİNE *", get_machines())
                customer = st.text_input("MÜŞTERİ")
                status = st.selectbox("SİPARİŞ DURUMU", ["BİTMEDİ", "BİTTİ"])
                
            with col2:
                standart = st.text_input("STANDART (Örn: DIN912)")
                kalite = st.text_input("KALİTE (Örn: 8.8)")
                cap = st.text_input("ÇAP (Örn: M10)")
                boy = st.text_input("BOY (Örn: 110)")
                diadim = st.text_input("DİŞ ADIMI")
                marka = st.text_input("MARKA")
                
            with col3:
                adet = st.text_input("PLANLANAN ADET")
                kg = st.text_input("PLANLANAN KG")
                net_gram = st.text_input("NET GRAM")
                brut_gram = st.text_input("BRÜT GRAM")
                hammadde = st.text_input("HAMMADDE")
                kaplama = st.text_input("KAPLAMA TİPİ")
                
            desc = st.text_area("AÇIKLAMA / NOT")
            
            submitted = st.form_submit_button("💾 Lotu Veritabanına Kaydet", use_container_width=True)
            
            if submitted:
                if not lot or not machine:
                    st.error("LOT NO ve MAKİNE alanları zorunludur!")
                else:
                    # Check duplicate
                    if get_lot_details(lot):
                        st.error(f"Hata: {lot} numaralı Lot zaten kayıtlı!")
                    else:
                        # Construct fields
                        new_data = {
                            "LOT": lot,
                            "MAKİNE": machine,
                            "MÜŞTERİ": customer,
                            "SİPARİŞ DURUMU": status,
                            "STANDART": standart,
                            "KALİTE": kalite,
                            "ÇAP": cap,
                            "BOY": boy,
                            "DİŞADIM": diadim,
                            "MARKA": marka,
                            "ADET": adet,
                            "KG": kg,
                            "NET GRAM": net_gram,
                            "BRÜT GRAM": brut_gram,
                            "HAMMADDE": hammadde,
                            "KAPLAMATIPI": kaplama,
                            "AÇIKLAMA": desc,
                            "sıra": float(len(get_active_queue(machine)) + 1)
                        }
                        
                        try:
                            add_new_lot(new_data)
                            st.session_state.modified_machines.add(machine)
                            st.success(f"{lot} numaralı Lot başarıyla kaydedildi!")
                        except Exception as e:
                            st.error(f"Kayıt sırasında hata oluştu: {e}")
    else:
        st.info("Lütfen veritabanını oluşturun.")

# ==========================================
# TAB 5: SYNC DATA
# ==========================================
with tab_sync:
    st.markdown("<h3 class='section-title'>Excel Veri Eşitleme (Göç) Paneli</h3>", unsafe_allow_html=True)
    st.info(f"**Kaynak Dosya:** `{migrate.ORIGINAL_EXCEL_PATH}`")
    
    st.markdown("""
    Bu panel, orijinal Excel planlama dosyanızdaki verileri okur, temizler ve yerel SQLite veritabanına aktarır.
    
    **⚠️ Dikkat:** Excel'den yeni aktarım yapmak mevcut SQLite veritabanınızı **sıfırlayacaktır**. Excel'den aktarım yapmadan önce veritabanınızı yedeklemek isteyebilirsiniz. Orijinal Excel dosyanıza **kesinlikle yazma yapılmaz**, dosyanız zarar görmez.
    """)
    
    # Run Migration Button
    if st.button("🔄 Excel'den Verileri SQLite Veritabanına Aktar (Sıfırla & Yükle)", type="primary", use_container_width=True):
        with st.spinner("Excel dosyası okunuyor ve veriler dönüştürülüyor... (Bu işlem birkaç saniye sürebilir)"):
            try:
                # Capture standard output to display in streamlit
                import sys
                from io import StringIO
                
                old_stdout = sys.stdout
                sys.stdout = mystdout = StringIO()
                
                migrate.run_migration()
                
                sys.stdout = old_stdout
                output_log = mystdout.getvalue()
                
                st.text_area("İşlem Logları:", output_log, height=250)
                st.success("Aktarım başarıyla tamamlandı! Lütfen sekmeleri yenileyin.")
            except Exception as e:
                st.error(f"Göç sırasında hata meydana geldi: {e}")
                
    st.markdown("<hr style='margin:30px 0; border:0; border-top:1px solid #ccc;'>", unsafe_allow_html=True)
    
    st.markdown("<h4 style='color:#1f4e79;'>💾 Değişiklikleri Orijinal Excel Dosyasına Kaydet</h4>", unsafe_allow_html=True)
    
    # Track and Display Changes
    modified_list = sorted(list(st.session_state.modified_machines)) if "modified_machines" in st.session_state else []
    
    if modified_list:
        st.warning(f"📋 **Aktif Değişiklik Yapılan Makineler ({len(modified_list)} adet):**\n" + ", ".join([f"`{m}`" for m in modified_list]))
    else:
        st.info("ℹ️ **Makinelerde henüz yeni bir sıralama veya değişiklik algılanmadı.**")
        
    st.markdown(f"""
    * **Hedef Dosya:** `{os.path.basename(migrate.ORIGINAL_EXCEL_PATH)}` (Dosya bilgisayarınızda açık olsa dahi güncellenebilir).
    * **Yedekleme:** İşlemden önce otomatik olarak `.bak` yedeği oluşturulur.
    """)
    
    # Option to only sync modified machines
    only_modified = st.checkbox("Sadece değişiklik yapılan makine sekmelerini kaydet (Çok Hızlı)", value=True, help="Eğer işaretli ise sadece yukarıda listelenen sekmeler güncellenir. İşaretli değilse tüm makine sekmeleri taranıp güncellenir.")
    
    # Sync SQLite to Excel Button
    if st.button("💾 Değişiklikleri Excel Dosyasına Kaydet (Geri Yaz)", use_container_width=True):
        with st.spinner("Değişiklikler Excel dosyasına yazılıyor ve formüller güncelleniyor..."):
            try:
                import sys
                # Clear python's import cache for win32com and pywintypes to force new search path resolution
                for m in ["pywintypes", "win32com", "win32com.client", "win32api"]:
                    sys.modules.pop(m, None)
                    
                import database
                import importlib
                importlib.reload(database)
                
                # Determine sheets to update
                sync_list = modified_list if (only_modified and modified_list) else None
                
                database.sync_db_to_excel(migrate.ORIGINAL_EXCEL_PATH, modified_machines=sync_list)
                
                # Success message with details
                if sync_list:
                    st.success(f"Tebrikler! Değişiklikler başarıyla kaydedildi. Güncellenen Makine Sayfaları: {', '.join([f'`{m}`' for m in sync_list])}. Orijinal dosyanın yedeği `.bak` uzantısıyla saklanmıştır.")
                else:
                    st.success(f"Tebrikler! Tüm makine sekmeleri başarıyla güncellendi ve `{os.path.basename(migrate.ORIGINAL_EXCEL_PATH)}` dosyasına kaydedildi.")
                
                # Clear modified machines list after successful sync
                st.session_state.modified_machines = set()
                st.rerun()
            except PermissionError:
                st.error("Hata: Excel dosyasına erişim engellendi! Lütfen dosyanın Excel programında kapalı olduğundan emin olun ve tekrar deneyin.")
            except Exception as e:
                st.error(f"Excel'e geri yazılırken bir hata oluştu: {e}")


# ==========================================
# SIDEBAR EMAIL REPORT
# ==========================================
with st.sidebar:
    st.markdown("<hr style='margin-top:50px;'>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center; color:#1f4e79;'>📧 Yönetici Raporu</h3>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:12px; text-align:center; color:#666;'>Tüm sekmelerdeki verileri (Simülatör, Makine Ayarları, Kuyruk ve Genel Durum) derleyip Outlook üzerinden yöneticinize gönderin.</p>", unsafe_allow_html=True)
    
    if st.button("Tüm Bilgileri Mail At", use_container_width=True, type="primary"):
        with st.spinner("Rapor derleniyor..."):
            try:
                import pandas as pd
                from database import get_db_connection, get_genel_makine
                
                # Fetch Machine Queue Summary
                conn = get_db_connection()
                q_summary = pd.read_sql_query("""
                    SELECT MAKİNE, 
                           SUM(CAST([NET TONAJ] AS REAL)) as Toplam_Tonaj_KG,
                           SUM(CAST([KALAN ADET] AS REAL)) as Toplam_Adet,
                           SUM(CAST([X GÜNLÜK İŞ] AS REAL)) as Toplam_Sure_Gun
                    FROM programs 
                    WHERE [SİPARİŞ DURUMU] = 'BİTMEDİ' AND MAKİNE IS NOT NULL AND MAKİNE != ''
                    GROUP BY MAKİNE
                    ORDER BY Toplam_Sure_Gun DESC
                """, conn)
                conn.close()
                
                # Fetch Machine Settings
                df_genel = get_genel_makine()
                if not df_genel.empty:
                    first_col = df_genel.columns[0]
                    ort_idx = df_genel.index[df_genel[first_col].astype(str).str.contains("2021", na=False)].tolist()
                    if ort_idx:
                        df_genel = df_genel.loc[:ort_idx[0]-1]
                    df_genel = df_genel.dropna(subset=[first_col]).reset_index(drop=True)
                
                # Prepare HTML segments
                html_cutoff = df_cutoff_summary.to_html(index=False, justify='center', border=1, classes='styled-table', float_format=lambda x: f'{x:,.1f}') if 'df_cutoff_summary' in globals() and not df_cutoff_summary.empty else "<p>Veri yok</p>"
                html_ihr = df_ihr[["Müşteri", "Geciken Tonaj (KG)", "Geciken Tonaj (Ton)", "Geciken Lot Sayısı"]].to_html(index=False, justify='center', border=1, classes='styled-table', float_format=lambda x: f'{x:,.1f}') if 'df_ihr' in globals() and not df_ihr.empty else ""
                html_yur = df_yur[["Müşteri", "Geciken Tonaj (KG)", "Geciken Tonaj (Ton)", "Geciken Lot Sayısı"]].to_html(index=False, justify='center', border=1, classes='styled-table', float_format=lambda x: f'{x:,.1f}') if 'df_yur' in globals() and not df_yur.empty else ""
                
                # Machine Queue HTML
                if not q_summary.empty:
                    q_summary['Toplam_Tonaj_KG'] = q_summary['Toplam_Tonaj_KG'].apply(lambda x: f"{x:,.0f}")
                    q_summary['Toplam_Adet'] = q_summary['Toplam_Adet'].apply(lambda x: f"{x:,.0f}")
                    q_summary['Toplam_Sure_Gun'] = q_summary['Toplam_Sure_Gun'].apply(lambda x: f"{x:,.1f}")
                    html_queue = q_summary.to_html(index=False, justify='center', border=1, classes='styled-table', float_format=lambda x: f'{x:,.1f}')
                else:
                    html_queue = "<p>Aktif kuyruk verisi bulunamadı.</p>"
                    
                # Machine Settings HTML
                if not df_genel.empty:
                    html_settings = df_genel.to_html(index=False, justify='center', border=1, classes='styled-table', float_format=lambda x: f'{x:,.1f}')
                else:
                    html_settings = "<p>Makine ayarları bulunamadı.</p>"
                
                # Simulator HTML (from globals)
                sim_days_val = globals().get('sim_days', 10.0)
                sim_kg_val = globals().get('sim_kg', 0.0)
                sim_adet_val = globals().get('sim_adet', 0.0)
                
                # Create HTML Body
                html_body = f"""
                <html>
                <head>
                <style>
                    body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; }}
                    h1 {{ color: #1f4e79; border-bottom: 2px solid #1f4e79; padding-bottom: 5px; }}
                    h2 {{ color: #2e75b6; margin-top: 25px; padding-bottom:3px; border-bottom: 1px solid #ccc; }}
                    h3 {{ color: #444; }}
                    .styled-table {{ border-collapse: collapse; margin-bottom: 20px; width: 95%; font-size: 13px; }}
                    .styled-table th {{ background-color: #1f4e79; color: white; padding: 8px; text-align: center; border: 1px solid #ddd; }}
                    .styled-table td {{ padding: 6px; text-align: center; border: 1px solid #ddd; }}
                    .styled-table tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    .kpi-container {{ padding: 15px; background-color: #f8f9fa; border-left: 5px solid #ffc000; margin-bottom: 20px; font-size: 14px; }}
                    .sim-container {{ padding: 15px; background-color: #e8f4f8; border-left: 5px solid #2980b9; margin-bottom: 20px; font-size: 14px; }}
                </style>
                </head>
                <body>
                    <h1>AFP Kurumsal Yönetim Raporu</h1>
                    
                    <div class="kpi-container">
                        <p><b>📊 1. GENEL DURUM (DASHBOARD) ÖZETİ:</b></p>
                        <ul>
                            <li><b>Sistemdeki Toplam Kayıtlı Lot:</b> {globals().get('total_lots', 0)}</li>
                            <li><b>Aktif Üretim Kuyruğundaki Lot Sayısı:</b> {globals().get('active_lots', 0)}</li>
                            <li><b>Aktif Planlanan Toplam Ağırlık:</b> {globals().get('total_kg_str', '0')}</li>
                            <li><b style="color:red;">Toplam Geciken Tonaj:</b> {globals().get('total_delayed_str', '0')}</li>
                        </ul>
                    </div>
                    
                    <div class="sim-container">
                        <p><b>🔮 2. SİMÜLATÖR TAHMİNİ ({sim_days_val} GÜNLÜK):</b></p>
                        <ul>
                            <li><b>Üretilecek Tahmini Tonaj:</b> {sim_kg_val:,.1f} KG ({(sim_kg_val/1000):,.1f} Ton)</li>
                            <li><b>Üretilecek Tahmini Adet:</b> {sim_adet_val:,.0f} Adet</li>
                        </ul>
                    </div>
                    
                    <h2>3. Makine Bazlı Aktif İş Yükü Özeti (Planlama Kuyruğu)</h2>
                    <p>Sistemde "BİTMEDİ" statüsünde olan işlerin makinelere göre toplam yük dağılımı:</p>
                    {html_queue}
                    
                    <h2>4. Makineler 'X Kesme Noktası' Aciliyet Özeti</h2>
                    {html_cutoff}
                    
                    <h2>5. Genel Makine Kapasite ve Hız Ayarları</h2>
                    {html_settings}
                """
                
                if html_ihr:
                    html_body += f"<h2>6. Geciken İhracat (Yurtdışı) Müşterileri</h2>{html_ihr}"
                if html_yur:
                    html_body += f"<h2>7. Geciken Yurtiçi Müşterileri</h2>{html_yur}"
                    
                html_body += """
                    <br><hr>
                    <p><i>Bu kurumsal rapor, AFP Planlama & Yönetim Sistemi tarafından otomatik olarak derlenmiştir.</i></p>
                </body>
                </html>
                """
                
                # Send via Outlook
                import win32com.client as win32
                import pythoncom
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.Subject = 'AFP Planlama ve Üretim - Kurumsal Yönetim Raporu'
                mail.HTMLBody = html_body
                mail.Display(False)
                
                st.success("✅ Outlook rapor taslağı başarıyla oluşturuldu! Lütfen Outlook ekranını kontrol edin.")
            except Exception as e:
                st.error(f"Outlook açılamadı veya taslak oluşturulamadı. Hata: {e}")
            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
