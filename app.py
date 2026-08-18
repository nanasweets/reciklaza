import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Spajanje reciklaže", layout="wide")
st.title("🔄 Provera kolona u Excel fajlu")

uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    xl = pd.ExcelFile(uploaded_file)
    
    st.subheader("📋 Dostupni sheetovi:")
    st.write(xl.sheet_names)
    
    # Prikaži kolone iz 'reciklaža'
    if "reciklaža" in xl.sheet_names:
        df_rec = pd.read_excel(uploaded_file, sheet_name="reciklaža", header=0, nrows=0)
        st.subheader("📋 Kolone u sheetu 'reciklaža':")
        st.write(list(df_rec.columns))
        
        # Prikaži prvih 5 redova
        df_rec_data = pd.read_excel(uploaded_file, sheet_name="reciklaža", header=0, nrows=5)
        st.dataframe(df_rec_data, use_container_width=True)
    else:
        st.error("Sheet 'reciklaža' nije pronađen!")
    
    # Prikaži kolone iz 'reklamacije ukupno'
    if "reklamacije ukupno" in xl.sheet_names:
        df_rek = pd.read_excel(uploaded_file, sheet_name="reklamacije ukupno", header=0, nrows=0)
        st.subheader("📋 Kolone u sheetu 'reklamacije ukupno':")
        st.write(list(df_rek.columns))
        
        # Prikaži prvih 5 redova
        df_rek_data = pd.read_excel(uploaded_file, sheet_name="reklamacije ukupno", header=0, nrows=5)
        st.dataframe(df_rek_data, use_container_width=True)
        
        # Pokaži sve kolone sa brojevima
        st.subheader("🔍 Pronađi kolonu sa tvojim brojevima (260101241...)")
        st.info("""
        **Pogledaj gornju tabelu i pronađi kolonu koja sadrži brojeve poput:**
        - 260101241
        - 241202280
        - 241200669
        
        **Zatim mi reci TAČAN NAZIV te kolone!**
        """)
    else:
        st.error("Sheet 'reklamacije ukupno' nije pronađen!")
