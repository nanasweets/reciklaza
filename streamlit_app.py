import streamlit as st
import pandas as pd

st.title("🔄 TEST APLIKACIJA - Reciklaža")

uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, header=0)
        st.success(f"✅ Učitano: {len(df)} redova")
        st.dataframe(df.head(10))
        st.write("Dostupne kolone:", list(df.columns))
    except Exception as e:
        st.error(f"Greška: {e}")
