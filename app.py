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
# POMOĆNE FUNKCIJE ZA STIL
# ============================

def excel_compatible(value):
    if pd.isna(value) or value is None:
        return None
    if isinstance(value, (int, float, str, bool)):
        return value
    if isinstance(value, (pd.Timestamp, datetime)):
        return value
    return str(value)

def primeni_stil_header(ws, max_col, lookup_kolone=None):
    """Primenjuje stil na header red (prvi red)."""
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=1, column=col_idx)
        kol = ws.cell(row=1, column=col_idx).value
        
        if lookup_kolone and kol in lookup_kolone:
            cell.fill = PatternFill("solid", fgColor=LOOKUP_BG)
            cell.font = Font(bold=True, color=LOOKUP_FONT, name="Arial", size=10)
        else:
            cell.fill = PatternFill("solid", fgColor=HEADER_BG)
            cell.font = Font(bold=True, color=HEADER_FONT, name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 36

def primeni_stil_podaci(ws, max_col, max_row, lookup_kolone=None):
    """Primenjuje stil na podatke (redovi 2 do max_row)."""
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for row_idx in range(2, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            col_name = ws.cell(row=1, column=col_idx).value
            
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            
            if lookup_kolone and col_name in lookup_kolone:
                cell.fill = PatternFill("solid", fgColor="EEF3FB")
            elif row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F5F8FD")

def prilagodi_sirine(ws, df):
    """Prilagođava širine kolona na osnovu sadržaja."""
    for col_idx, kol in enumerate(df.columns, start=1):
        max_len = len(str(kol))
        for row_idx in range(2, min(len(df) + 2, 52)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, min(len(str(val)), 40))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3

def dodaj_legendu(ws, legenda_row, lookup_kolone=None):
    """Dodaje legendu na dno lista."""
    ws.cell(row=legenda_row, column=1, value="Legenda:").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=legenda_row + 1, column=1, value="■ Tamno plava = ručno uneti podaci").font = Font(name="Arial", size=9, color=HEADER_BG)
    ws.cell(row=legenda_row + 2, column=1, value="■ Svetlo plava = automatski povučeno").font = Font(name="Arial", size=9, color=LOOKUP_FONT)
    ws.cell(row=legenda_row + 4, column=1, value=f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}").font = Font(italic=True, name="Arial", size=9, color="888888")

# ============================
# FUNKCIJE ZA UPIS SHEET-OVA
# ============================

def upisi_zbirno(ws, df):
    """Sheet: zbirno – svi podaci."""
    lookup_kolone = set(MAPPING.keys())
    
    # Header
    for col_idx, kol in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=kol)
    primeni_stil_header(ws, len(df.columns), lookup_kolone)
    
    # Podaci
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx).value = excel_compatible(value)
    primeni_stil_podaci(ws, len(df.columns), len(df) + 1, lookup_kolone)
    prilagodi_sirine(ws, df)
    
    # Legenda
    dodaj_legendu(ws, len(df) + 3, lookup_kolone)
    
    # Zamrzni header
    ws.freeze_panes = "A2"

def upisi_kpi(ws, df):
    """Sheet: KPI – ključni pokazatelji."""
    # Izračunavanje
    ukupno_artikala = df["Količina"].sum() if "Količina" in df.columns else 0
    ukupna_vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    broj_brendova = df["Brend"].nunique() if "Brend" in df.columns else 0
    broj_grupa = df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0
    
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        min_datum = df_datum["Datum obrade"].min()
        max_datum = df_datum["Datum obrade"].max()
        period = f"{min_datum.strftime('%d.%m.%Y') if pd.notna(min_datum) else 'N/A'} - {max_datum.strftime('%d.%m.%Y') if pd.notna(max_datum) else 'N/A'}"
    else:
        period = "N/A"
    
    # Upis
    podaci = [
        ["Ključni pokazatelj", "Vrednost"],
        ["Ukupan broj artikala", ukupno_artikala],
        ["Ukupna vrednost (RSD)", f"{ukupna_vrednost:,.2f}"],
        ["Broj brendova", broj_brendova],
        ["Broj robnih grupa", broj_grupa],
        ["Period", period],
        ["Broj redova", len(df)],
    ]
    
    for row_idx, row in enumerate(podaci, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    # Stil
    for col_idx in [1, 2]:
        ws.column_dimensions[get_column_letter(col_idx)].width = 25
    
    # Bold za prvu kolonu
    for row_idx in range(1, len(podaci) + 1):
        ws.cell(row=row_idx, column=1).font = Font(bold=True, name="Arial", size=11)
        ws.cell(row=row_idx, column=2).font = Font(name="Arial", size=11)
    
    # Naslov
    ws.cell(row=1, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    ws.cell(row=1, column=2).font = Font(bold=True, size=12, color=HEADER_BG)

def upisi_grupisano(ws, df, kolona, naslov):
    """Sheet: grupisanje po koloni (npr. klasifikacija reciklaže ili štete)."""
    if kolona not in df.columns:
        ws.cell(row=1, column=1, value=f"Kolona '{kolona}' ne postoji")
        return
    
    grupisan = df[kolona].value_counts().reset_index()
    grupisan.columns = [kolona, "Broj artikala"]
    
    # Header
    ws.cell(row=1, column=1, value=kolona)
    ws.cell(row=1, column=2, value="Broj artikala")
    primeni_stil_header(ws, 2)
    
    # Podaci
    for row_idx, row in enumerate(grupisan.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
    primeni_stil_podaci(ws, 2, len(grupisan) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20

def upisi_top_brendove(ws, df):
    """Sheet: Top 10 brendova po broju i vrednosti."""
    if "Brend" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Brend' ne postoji")
        return
    
    # Po broju
    brendovi_broj = df["Brend"].value_counts().head(10).reset_index()
    brendovi_broj.columns = ["Brend", "Broj artikala"]
    
    # Po vrednosti
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        brendovi_vred = df.groupby("Brend")["Ukupna vrednost"].sum().sort_values(ascending=False).head(10).reset_index()
        brendovi_vred.columns = ["Brend", "Ukupna vrednost (RSD)"]
    else:
        brendovi_vred = pd.DataFrame(columns=["Brend", "Ukupna vrednost (RSD)"])
    
    # Upis - po broju
    ws.cell(row=1, column=1, value="Top 10 brendova po broju artikala")
    ws.cell(row=2, column=1, value="Brend")
    ws.cell(row=2, column=2, value="Broj artikala")
    
    for idx, row in enumerate(brendovi_broj.itertuples(index=False), start=3):
        ws.cell(row=idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=idx, column=2).value = excel_compatible(row[1])
    
    # Upis - po vrednosti
    ws.cell(row=1, column=4, value="Top 10 brendova po ukupnoj vrednosti")
    ws.cell(row=2, column=4, value="Brend")
    ws.cell(row=2, column=5, value="Ukupna vrednost (RSD)")
    
    for idx, row in enumerate(brendovi_vred.itertuples(index=False), start=3):
        ws.cell(row=idx, column=4).value = excel_compatible(row[0])
        ws.cell(row=idx, column=5).value = excel_compatible(row[1])
    
    # Stil
    for col in [1, 2, 4, 5]:
        ws.column_dimensions[get_column_letter(col)].width = 25
    
    # Header stil
    for row in [1, 2]:
        for col in [1, 2, 4, 5]:
            cell = ws.cell(row=row, column=col)
            if row == 1:
                cell.font = Font(bold=True, size=12, color=HEADER_BG)
            else:
                cell.font = Font(bold=True, size=11, color=HEADER_BG)

def upisi_top_grupe(ws, df):
    """Sheet: Top 10 robnih grupa."""
    if "Robna grupa" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Robna grupa' ne postoji")
        return
    
    grupe = df["Robna grupa"].value_counts().head(10).reset_index()
    grupe.columns = ["Robna grupa", "Broj artikala"]
    
    ws.cell(row=1, column=1, value="Robna grupa")
    ws.cell(row=1, column=2, value="Broj artikala")
    primeni_stil_header(ws, 2)
    
    for row_idx, row in enumerate(grupe.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
    primeni_stil_podaci(ws, 2, len(grupe) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20

def upisi_vremenski_trend(ws, df):
    """Sheet: Vremenski trend – po mesecima."""
    if "Datum obrade" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Datum obrade' ne postoji")
        return
    
    df_datum = df.copy()
    df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
    df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
    
    # Grupisanje po mesecima
    trend = df_datum.groupby("Mesec").size().reset_index()
    trend.columns = ["Mesec", "Broj artikala"]
    trend["Mesec"] = trend["Mesec"].astype(str)
    
    # Vrednost po mesecima
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df_datum["Ukupna vrednost"] = df_datum["Nabavna cena"] * df_datum["Količina"]
        vrednost_po_mesecu = df_datum.groupby("Mesec")["Ukupna vrednost"].sum().reset_index()
        vrednost_po_mesecu.columns = ["Mesec", "Ukupna vrednost (RSD)"]
        vrednost_po_mesecu["Mesec"] = vrednost_po_mesecu["Mesec"].astype(str)
        trend = trend.merge(vrednost_po_mesecu, on="Mesec", how="left")
    
    # Upis
    ws.cell(row=1, column=1, value="Mesec")
    ws.cell(row=1, column=2, value="Broj artikala")
    if "Ukupna vrednost (RSD)" in trend.columns:
        ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
        primeni_stil_header(ws, 3)
    else:
        primeni_stil_header(ws, 2)
    
    for row_idx, row in enumerate(trend.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        if len(row) > 2:
            ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    
    primeni_stil_podaci(ws, len(trend.columns), len(trend) + 1)
    
    for col_idx in range(1, len(trend.columns) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

def upisi_prevoznike(ws, df):
    """Sheet: Po prevozniku."""
    if "prevoznik" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'prevoznik' ne postoji")
        return
    
    # Po prevozniku
    prevoznici = df["prevoznik"].value_counts().reset_index()
    prevoznici.columns = ["Prevoznik", "Broj artikala"]
    
    # Štete po prevozniku
    if "Klasifikacija štete" in df.columns:
        stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
        stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
    
    # Upis
    ws.cell(row=1, column=1, value="Prevoznik")
    ws.cell(row=1, column=2, value="Broj artikala")
    ws.cell(row=1, column=4, value="Štete po prevozniku")
    ws.cell(row=2, column=4, value="Prevoznik")
    ws.cell(row=2, column=5, value="Klasifikacija štete")
    ws.cell(row=2, column=6, value="Broj")
    
    primeni_stil_header(ws, 2)
    for col in [4, 5, 6]:
        ws.cell(row=1, column=col).font = Font(bold=True, size=11, color=HEADER_BG)
        ws.cell(row=2, column=col).font = Font(bold=True, size=10, color=HEADER_BG)
    
    # Podaci - prevoznici
    for row_idx, row in enumerate(prevoznici.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
    primeni_stil_podaci(ws, 2, len(prevoznici) + 1)
    
    # Podaci - štete
    if "Klasifikacija štete" in df.columns:
        for row_idx, row in enumerate(stete_prevoz.itertuples(index=False), start=3):
            ws.cell(row=row_idx, column=4).value = excel_compatible(row[0])
            ws.cell(row=row_idx, column=5).value = excel_compatible(row[1])
            ws.cell(row=row_idx, column=6).value = excel_compatible(row[2])
        primeni_stil_podaci(ws, 6, len(stete_prevoz) + 3)
    
    for col in [1, 2, 4, 5, 6]:
        ws.column_dimensions[get_column_letter(col)].width = 25

# ============================
# GLAVNA FUNKCIJA ZA EXCEL
# ============================

def formatiraj_i_sacuvaj(df):
    """
    Pravi Excel fajl sa više sheet-ova:
    1. zbirno – svi podaci
    2. KPI – ključni pokazatelji
    3. Po klasifikaciji – klasifikacija reciklaže
    4. Po šteti – klasifikacija štete
    5. Top brendovi – top 10 brendova
    6. Top grupe – top 10 robnih grupa
    7. Vremenski trend – mesečni pregled
    8. Po prevozniku – analiza prevoznika
    """
    wb = Workbook()
    
    # 1. zbirno
    ws = wb.active
    ws.title = "zbirno"
    upisi_zbirno(ws, df)
    
    # 2. KPI
    ws = wb.create_sheet("KPI")
    upisi_kpi(ws, df)
    
    # 3. Po klasifikaciji
    ws = wb.create_sheet("Po klasifikaciji")
    upisi_grupisano(ws, df, "Klasifikacija reciklaže", "Klasifikacija reciklaže")
    
    # 4. Po šteti
    ws = wb.create_sheet("Po šteti")
    upisi_grupisano(ws, df, "Klasifikacija štete", "Klasifikacija štete")
    
    # 5. Top brendovi
    ws = wb.create_sheet("Top brendovi")
    upisi_top_brendove(ws, df)
    
    # 6. Top grupe
    ws = wb.create_sheet("Top grupe")
    upisi_top_grupe(ws, df)
    
    # 7. Vremenski trend
    ws = wb.create_sheet("Vremenski trend")
    upisi_vremenski_trend(ws, df)
    
    # 8. Po prevozniku
    ws = wb.create_sheet("Po prevozniku")
    upisi_prevoznike(ws, df)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================
# FUNKCIJE ZA VIZUELNE IZVEŠTAJE (STREAMLIT)
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
        
        if "Klasifikacija štete" in df.columns:
            stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
            stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
            
            fig2 = px.bar(stete_prevoz, x="Prevoznik", y="Broj", color="Klasifikacija štete",
                          title="Klasifikacije štete po prevoznicima",
                          barmode="group")
            st.plotly_chart(fig2, use_container_width=True)

def prikazi_detaljnu_tabelu(df):
    st.subheader("📋 Detaljna tabela sa filterima")
    
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
# GLAVNI DEO – STREAMLIT UI
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
        
        prikazi_kpi(df)
        st.markdown("---")
        
        prikazi_klasifikaciju_reciklaze(df)
        st.markdown("---")
        
        prikazi_klasifikaciju_stete(df)
        st.markdown("---")
        
        prikazi_top_brendove(df)
        st.markdown("---")
        
        prikazi_robne_grupe(df)
        st.markdown("---")
        
        prikazi_vremenski_trend(df)
        st.markdown("---")
        
        prikazi_prevoznike(df)
        st.markdown("---")
        
        prikazi_detaljnu_tabelu(df)
        st.markdown("---")
        
        # ======================
        # PREUZIMANJE
        # ======================
        
        st.subheader("📥 Preuzimanje izveštaja")
        
        col1, col2 = st.columns([2, 1])
        
        with col1:
            st.info("""
            📋 **Excel izveštaj sadrži 8 sheet-ova:**
            1. **zbirno** – svi spojeni podaci
            2. **KPI** – ključni pokazatelji
            3. **Po klasifikaciji** – analiza po vrstama reciklaže
            4. **Po šteti** – analiza po vrstama štete
            5. **Top brendovi** – top 10 brendova
            6. **Top grupe** – top 10 robnih grupa
            7. **Vremenski trend** – mesečni pregled
            8. **Po prevozniku** – analiza prevoznika
            """)
        
        with col2:
            excel_data = formatiraj_i_sacuvaj(df)
            st.download_button(
                label="📥 Preuzmi Excel izveštaj",
                data=excel_data,
                file_name=f"Izvestaj_o_reciklazi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.caption("📁 Veličina fajla: ~{:.1f} KB".format(len(excel_data) / 1024))
