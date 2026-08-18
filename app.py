import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import re
import plotly.express as px
import plotly.graph_objects as go

# ============================
# KONFIGURACIJA
# ============================

st.set_page_config(page_title="Izveštaj o reciklaži", layout="wide")

SHEET_RECIKLAZA = "reciklaža"
SHEET_REKLAMACIJE = "reklamacije ukupno"

MAPPING = {
    "reklamacija otvorena od strane": "Reklamacija: Created By",
    "model reklamacije": "RMA Model",
    "klasifikacija": "Klasifikacija",
    "način razduženja kupca": "Način razduženja kupca",
    "serijski broj": "Serijski broj uređaja",
}

KOLONE_OUTPUT = [
    "Broj reklamacije", "ID uređaja", "Naziv artikla", "Brend",
    "Robna grupa", "Količina", "Klasifikacija reciklaže", "Datum obrade",
    "Paleta", "Nabavna cena", "ID ulaza", "Skladišna lokacija",
    "Pozicija u rafu", "Klasifikacija štete", "prevoznik",
    "broj pošiljke", "vrednost fakture",
    "reklamacija otvorena od strane", "model reklamacije",
    "klasifikacija", "način razduženja kupca", "serijski broj",
]

HEADER_BG = "1F4E79"
HEADER_FONT = "FFFFFF"
LOOKUP_BG = "D9E1F2"
LOOKUP_FONT = "1F4E79"

# ============================
# FUNKCIJE ZA OBRADU PODATAKA
# ============================

def cisti_broj(vrednost):
    if pd.isna(vrednost):
        return ""
    tekst = str(vrednost).strip()
    cifre = re.sub(r'[^0-9]', '', tekst)
    return cifre

def pronadji_kolonu_broja(df):
    for kol in df.columns:
        kol_lower = kol.lower()
        if "broj" in kol_lower and "reklamacije" in kol_lower:
            return kol
        if "reklamacija" in kol_lower and ("broj" in kol_lower or "id" in kol_lower):
            return kol
    for kol in df.columns:
        if "broj" in kol.lower() or "id" in kol.lower():
            return kol
    return df.columns[0]

def ucitaj_podatke(fajl):
    xl = pd.ExcelFile(fajl)
    if SHEET_RECIKLAZA not in xl.sheet_names:
        st.error(f"Sheet '{SHEET_RECIKLAZA}' nije pronađen!")
        return None, None
    if SHEET_REKLAMACIJE not in xl.sheet_names:
        st.error(f"Sheet '{SHEET_REKLAMACIJE}' nije pronađen!")
        return None, None
    reciklaza = pd.read_excel(fajl, sheet_name=SHEET_RECIKLAZA, header=0)
    reklamacije = pd.read_excel(fajl, sheet_name=SHEET_REKLAMACIJE, header=0)
    return reciklaza, reklamacije

def pripremi_lookup(reklamacije):
    kolona_broja = pronadji_kolonu_broja(reklamacije)
    rek = reklamacije.copy()
    rek["_kljuc"] = rek[kolona_broja].apply(cisti_broj)
    rek = rek[rek["_kljuc"] != ""]
    return rek.set_index("_kljuc")

def spoji(reciklaza, lookup):
    df = reciklaza.copy()
    df["_kljuc"] = df["Broj reklamacije"].apply(cisti_broj)
    
    for out_kol, src_kol in MAPPING.items():
        if src_kol not in lookup.columns:
            df[out_kol] = None
            continue
        vrednosti = []
        for kljuc in df["_kljuc"]:
            if kljuc == "":
                vrednosti.append(None)
            elif kljuc in lookup.index:
                vrednosti.append(lookup.at[kljuc, src_kol])
            else:
                vrednosti.append(None)
        df[out_kol] = vrednosti
    
    df = df.drop(columns=["_kljuc"])
    extras = [c for c in df.columns if c not in KOLONE_OUTPUT]
    finalne = [c for c in KOLONE_OUTPUT if c in df.columns] + extras
    return df[finalne]

# ============================
# POMOĆNA FUNKCIJA ZA EXCEL
# ============================

def excel_compatible(value):
    """Pretvara vrednost u oblik koji Excel može da prihvati."""
    if pd.isna(value) or value is None:
        return None
    if isinstance(value, (int, float, str, bool)):
        return value
    if isinstance(value, (pd.Timestamp, datetime)):
        return value
    if isinstance(value, pd.Timedelta):
        return None  # ili možda vrednost u sekundama
    # Svi ostali tipovi -> pretvori u string
    return str(value)

def formatiraj_i_sacuvaj(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "zbirno"
    
    lookup_kolone = set(MAPPING.keys())
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # Header
    for col_idx, kol in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=kol)
        if kol in lookup_kolone:
            cell.fill = PatternFill("solid", fgColor=LOOKUP_BG)
            cell.font = Font(bold=True, color=LOOKUP_FONT, name="Arial", size=10)
        else:
            cell.fill = PatternFill("solid", fgColor=HEADER_BG)
            cell.font = Font(bold=True, color=HEADER_FONT, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 36
    
    # Podaci
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            # Koristi pomoćnu funkciju
            cell.value = excel_compatible(value)
            
            col_name = df.columns[col_idx - 1]
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if col_name in lookup_kolone:
                cell.fill = PatternFill("solid", fgColor="EEF3FB")
            elif row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F5F8FD")
    
    # Širine kolona
    for col_idx, kol in enumerate(df.columns, start=1):
        max_len = len(str(kol))
        for row_idx in range(2, min(len(df) + 2, 52)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, min(len(str(val)), 40))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3
    
    ws.freeze_panes = "A2"
    
    # Legenda
    legenda_row = len(df) + 3
    ws.cell(row=legenda_row, column=1, value="Legenda:").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=legenda_row + 1, column=1, value="■ Tamno plava = ručno uneti podaci").font = Font(name="Arial", size=9, color=HEADER_BG)
    ws.cell(row=legenda_row + 2, column=1, value="■ Svetlo plava = automatski povučeno").font = Font(name="Arial", size=9, color=LOOKUP_FONT)
    ws.cell(row=legenda_row + 4, column=1, value=f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}").font = Font(italic=True, name="Arial", size=9, color="888888")
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================
# FUNKCIJE ZA IZVEŠTAJE
# ============================

def prikazi_kpi(df):
    st.subheader("📊 Ključni pokazatelji")
    
    col1, col2, col3, col4 = st.columns(4)
    
    ukupno_artikala = df["Količina"].sum() if "Količina" in df.columns else 0
    ukupna_vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    broj_brendova = df["Brend"].nunique() if "Brend" in df.columns else 0
    broj_grupa = df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0
    
    col1.metric("📦 Ukupno artikala", f"{ukupno_artikala:,}")
    col2.metric("💰 Ukupna vrednost", f"{ukupna_vrednost:,.2f} RSD")
    col3.metric("🏷️ Broj brendova", broj_brendova)
    col4.metric("📂 Robnih grupa", broj_grupa)
    
    if "Datum obrade" in df.columns:
        df["Datum obrade"] = pd.to_datetime(df["Datum obrade"], errors="coerce")
        min_datum = df["Datum obrade"].min()
        max_datum = df["Datum obrade"].max()
        st.caption(f"📅 Period: {min_datum.strftime('%d.%m.%Y') if pd.notna(min_datum) else 'N/A'} - {max_datum.strftime('%d.%m.%Y') if pd.notna(max_datum) else 'N/A'}")

def prikazi_klasifikaciju_reciklaze(df):
    st.subheader("📊 Klasifikacija reciklaže")
    
    if "Klasifikacija reciklaže" in df.columns:
        klasifikacija = df["Klasifikacija reciklaže"].value_counts().reset_index()
        klasifikacija.columns = ["Klasifikacija", "Broj artikala"]
        
        fig = px.pie(klasifikacija, values="Broj artikala", names="Klasifikacija", 
                     title="Raspored po klasifikaciji reciklaže",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(klasifikacija, use_container_width=True)

def prikazi_klasifikaciju_stete(df):
    st.subheader("📊 Klasifikacija štete")
    
    if "Klasifikacija štete" in df.columns:
        stete = df["Klasifikacija štete"].value_counts().reset_index()
        stete.columns = ["Klasifikacija štete", "Broj artikala"]
        
        fig = px.bar(stete, x="Klasifikacija štete", y="Broj artikala",
                     title="Najčešće klasifikacije štete",
                     color="Broj artikala", color_continuous_scale="Blues")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)
        
        st.dataframe(stete, use_container_width=True)

def prikazi_top_brendove(df):
    st.subheader("🏆 Top 10 brendova")
    
    if "Brend" in df.columns:
        brendovi = df["Brend"].value_counts().head(10).reset_index()
        brendovi.columns = ["Brend", "Broj artikala"]
        
        fig = px.bar(brendovi, x="Brend", y="Broj artikala",
                     title="Top 10 brendova po broju artikala na reciklaži",
                     color="Broj artikala", color_continuous_scale="Blues")
        st.plotly_chart(fig, use_container_width=True)
        
        # Vrednost po brendovima
        if "Nabavna cena" in df.columns and "Količina" in df.columns:
            df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
            vrednost_brend = df.groupby("Brend")["Ukupna vrednost"].sum().sort_values(ascending=False).head(10).reset_index()
            vrednost_brend.columns = ["Brend", "Ukupna vrednost (RSD)"]
            
            fig2 = px.bar(vrednost_brend, x="Brend", y="Ukupna vrednost (RSD)",
                          title="Top 10 brendova po ukupnoj vrednosti",
                          color="Ukupna vrednost (RSD)", color_continuous_scale="Greens")
            st.plotly_chart(fig2, use_container_width=True)

def prikazi_robne_grupe(df):
    st.subheader("📂 Top 10 robnih grupa")
    
    if "Robna grupa" in df.columns:
        grupe = df["Robna grupa"].value_counts().head(10).reset_index()
        grupe.columns = ["Robna grupa", "Broj artikala"]
        
        fig = px.bar(grupe, x="Robna grupa", y="Broj artikala",
                     title="Top 10 robnih grupa",
                     color="Broj artikala", color_continuous_scale="Reds")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

def prikazi_vremenski_trend(df):
    st.subheader("📈 Vremenski trend")
    
    if "Datum obrade" in df.columns:
        df["Datum obrade"] = pd.to_datetime(df["Datum obrade"], errors="coerce")
        df["Mesec"] = df["Datum obrade"].dt.to_period("M")
        
        meseci = df.groupby("Mesec").size().reset_index()
        meseci.columns = ["Mesec", "Broj artikala"]
        meseci["Mesec"] = meseci["Mesec"].astype(str)
        
        fig = px.line(meseci, x="Mesec", y="Broj artikala",
                      title="Broj artikala po mesecima",
                      markers=True)
        st.plotly_chart(fig, use_container_width=True)

def prikazi_prevoznike(df):
    st.subheader("🚚 Analiza prevoznika")
    
    if "prevoznik" in df.columns:
        prevoznici = df["prevoznik"].value_counts().reset_index()
        prevoznici.columns = ["Prevoznik", "Broj artikala"]
        
        fig = px.pie(prevoznici, values="Broj artikala", names="Prevoznik",
                     title="Raspored po prevoznicima")
        st.plotly_chart(fig, use_container_width=True)
        
        # Štete po prevozniku
        if "Klasifikacija štete" in df.columns:
            stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
            stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
            
            fig2 = px.bar(stete_prevoz, x="Prevoznik", y="Broj", color="Klasifikacija štete",
                          title="Klasifikacije štete po prevoznicima",
                          barmode="group")
            st.plotly_chart(fig2, use_container_width=True)

def prikazi_detaljnu_tabelu(df):
    st.subheader("📋 Detaljna tabela sa filterima")
    
    # Filteri
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if "Brend" in df.columns:
            brendovi = ["Svi"] + sorted(df["Brend"].dropna().unique().tolist())
            izabrani_brend = st.selectbox("Filter po brendu", brendovi)
    
    with col2:
        if "Robna grupa" in df.columns:
            grupe = ["Sve"] + sorted(df["Robna grupa"].dropna().unique().tolist())
            izabrana_grupa = st.selectbox("Filter po robnoj grupi", grupe)
    
    with col3:
        if "Klasifikacija reciklaže" in df.columns:
            klas = ["Sve"] + sorted(df["Klasifikacija reciklaže"].dropna().unique().tolist())
            izabrana_klas = st.selectbox("Filter po klasifikaciji reciklaže", klas)
    
    # Aplikacija filtera
    df_filtered = df.copy()
    if "Brend" in df.columns and izabrani_brend != "Svi":
        df_filtered = df_filtered[df_filtered["Brend"] == izabrani_brend]
    if "Robna grupa" in df.columns and izabrana_grupa != "Sve":
        df_filtered = df_filtered[df_filtered["Robna grupa"] == izabrana_grupa]
    if "Klasifikacija reciklaže" in df.columns and izabrana_klas != "Sve":
        df_filtered = df_filtered[df_filtered["Klasifikacija reciklaže"] == izabrana_klas]
    
    st.dataframe(df_filtered, use_container_width=True)
    st.caption(f"Prikazano {len(df_filtered)} od {len(df)} redova")

# ============================
# STREAMLIT UI
# ============================

st.title("🔄 Izveštaj o reciklaži")
st.markdown("---")

uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Učitavam podatke..."):
        reciklaza, reklamacije = ucitaj_podatke(uploaded_file)
    
    if reciklaza is not None and reklamacije is not None:
        with st.spinner("Spajam podatke..."):
            lookup = pripremi_lookup(reklamacije)
            df = spoji(reciklaza, lookup)
        
        st.success(f"✅ Učitano: {len(reciklaza)} redova u 'reciklaža', {len(reklamacije)} redova u 'reklamacije ukupno'")
        st.info(f"📊 Spojeno: {len(df)} redova")
        
        # ======================
        # IZVEŠTAJI
        # ======================
        
        # 1. KPI
        prikazi_kpi(df)
        st.markdown("---")
        
        # 2. Klasifikacija reciklaže
        prikazi_klasifikaciju_reciklaze(df)
        st.markdown("---")
        
        # 3. Klasifikacija štete
        prikazi_klasifikaciju_stete(df)
        st.markdown("---")
        
        # 4. Top brendovi
        prikazi_top_brendove(df)
        st.markdown("---")
        
        # 5. Robne grupe
        prikazi_robne_grupe(df)
        st.markdown("---")
        
        # 6. Vremenski trend
        prikazi_vremenski_trend(df)
        st.markdown("---")
        
        # 7. Prevoznici
        prikazi_prevoznike(df)
        st.markdown("---")
        
        # 8. Detaljna tabela
        prikazi_detaljnu_tabelu(df)
        st.markdown("---")
        
        # 9. Preuzimanje
        st.subheader("📥 Preuzimanje")
        
        excel_data = formatiraj_i_sacuvaj(df)
        st.download_button(
            label="📥 Preuzmi Excel fajl sa izveštajem",
            data=excel_data,
            file_name=f"Izvestaj_o_reciklazi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
