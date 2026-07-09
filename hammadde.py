import streamlit as st
import json
import os
import pandas as pd
from database import get_db_connection

JSON_PATH = os.path.join(os.path.dirname(__file__), "hammadde_lots.json")

def load_lots():
    if os.path.exists(JSON_PATH):
        try:
            with open(JSON_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_lots(lots):
    with open(JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(lots, f, ensure_ascii=False, indent=4)

def check_lot_status(lot):
    conn = get_db_connection()
    q = "SELECT [SİPARİŞ DURUMU] FROM programs WHERE LOT = ?"
    df = pd.read_sql(q, conn, params=(lot,))
    conn.close()
    if not df.empty:
        status = df.iloc[0]['SİPARİŞ DURUMU']
        if pd.isna(status):
            return "BİLİNMİYOR"
        return str(status).strip().upper()
    return "BULUNAMADI"

def render_hammadde():
    st.title("📦 Hammadde Yönetimi")
    st.markdown("### 🔍 Lot Takip Listesi")
    
    lots = load_lots()
    
    col1, col2 = st.columns([3, 1])
    with col1:
        new_lot = st.text_input("Takip edilecek yeni LOT girin:", placeholder="Örn: 2604431")
    with col2:
        st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
        if st.button("➕ Ekle", use_container_width=True):
            if new_lot and new_lot not in lots:
                lots.append(new_lot)
                save_lots(lots)
                st.rerun()
            elif new_lot in lots:
                st.warning("Bu Lot zaten takip ediliyor!")
                
    st.markdown("<hr>", unsafe_allow_html=True)
    
    if not lots:
        st.info("Henüz takip edilen lot bulunmuyor.")
    else:
        for lot in lots:
            status = check_lot_status(lot)
            
            # Create a nice layout for each lot
            card_col, btn_col = st.columns([4, 1])
            with card_col:
                if status == "BİTTİ":
                    st.success(f"✅ **LOT:** {lot} - **Durum:** BİTTİ (Üretime Hazır)")
                elif status == "BİTMEDİ":
                    st.warning(f"⏳ **LOT:** {lot} - **Durum:** BİTMEDİ (Bekleniyor)")
                elif status == "BULUNAMADI":
                    st.error(f"❌ **LOT:** {lot} - **Durum:** Veritabanında Bulunamadı")
                else:
                    st.info(f"ℹ️ **LOT:** {lot} - **Durum:** {status}")
                    
            with btn_col:
                if st.button("🗑️ Sil", key=f"del_{lot}", use_container_width=True):
                    lots.remove(lot)
                    save_lots(lots)
                    st.rerun()
