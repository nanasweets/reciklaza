import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import re
import plotly.express as px
import plotly.graph_objects as go
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm, inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import matplotlib.pyplot as plt
import os

# ============================
# KONFIGURACIJA
# ============================

st.set_page_config(
    page_title="Izveštaj o reciklaži",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Tamna tema
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

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
# POMOĆNE FUNKCIJE ZA EXCEL
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
    for col_idx, kol in enumerate(df.columns, start=1):
        max_len = len(str(kol))
        for row_idx in range(2, min(len(df) + 2, 52)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, min(len(str(val)), 40))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3

def dodaj_legendu(ws, legenda_row, lookup_kolone=None):
    ws.cell(row=legenda_row, column=1, value="Legenda:").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=legenda_row + 1, column=1, value="■ Tamno plava = ručno uneti podaci").font = Font(name="Arial", size=9, color=HEADER_BG)
    ws.cell(row=legenda_row + 2, column=1, value="■ Svetlo plava = automatski povučeno").font = Font(name="Arial", size=9, color=LOOKUP_FONT)
    ws.cell(row=legenda_row + 4, column=1, value=f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}").font = Font(italic=True, name="Arial", size=9, color="888888")

# ============================
# FUNKCIJE ZA UPIS SHEET-OVA
# ============================

def upisi_zbirno(ws, df):
    lookup_kolone = set(MAPPING.keys())
    
    for col_idx, kol in enumerate(df.columns, start=1):
        ws.cell(row=1, column=col_idx, value=kol)
    primeni_stil_header(ws, len(df.columns), lookup_kolone)
    
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx).value = excel_compatible(value)
    primeni_stil_podaci(ws, len(df.columns), len(df) + 1, lookup_kolone)
    prilagodi_sirine(ws, df)
    dodaj_legendu(ws, len(df) + 3, lookup_kolone)
    ws.freeze_panes = "A2"

def upisi_kpi(ws, df):
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
    
    for col_idx in [1, 2]:
        ws.column_dimensions[get_column_letter(col_idx)].width = 25
    
    for row_idx in range(1, len(podaci) + 1):
        ws.cell(row=row_idx, column=1).font = Font(bold=True, name="Arial", size=11)
        ws.cell(row=row_idx, column=2).font = Font(name="Arial", size=11)
    
    ws.cell(row=1, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    ws.cell(row=1, column=2).font = Font(bold=True, size=12, color=HEADER_BG)

def upisi_grupisano(ws, df, kolona, naslov):
    if kolona not in df.columns:
        ws.cell(row=1, column=1, value=f"Kolona '{kolona}' ne postoji")
        return
    
    grupisan = df[kolona].value_counts().reset_index()
    grupisan.columns = [kolona, "Broj artikala"]
    
    ws.cell(row=1, column=1, value=kolona)
    ws.cell(row=1, column=2, value="Broj artikala")
    primeni_stil_header(ws, 2)
    
    for row_idx, row in enumerate(grupisan.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
    primeni_stil_podaci(ws, 2, len(grupisan) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20

def upisi_top_brendove(ws, df):
    if "Brend" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Brend' ne postoji")
        return
    
    brendovi_broj = df["Brend"].value_counts().head(10).reset_index()
    brendovi_broj.columns = ["Brend", "Broj artikala"]
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        brendovi_vred = df.groupby("Brend")["Ukupna vrednost"].sum().sort_values(ascending=False).head(10).reset_index()
        brendovi_vred.columns = ["Brend", "Ukupna vrednost (RSD)"]
    else:
        brendovi_vred = pd.DataFrame(columns=["Brend", "Ukupna vrednost (RSD)"])
    
    ws.cell(row=1, column=1, value="Top 10 brendova po broju artikala")
    ws.cell(row=2, column=1, value="Brend")
    ws.cell(row=2, column=2, value="Broj artikala")
    
    for idx, row in enumerate(brendovi_broj.itertuples(index=False), start=3):
        ws.cell(row=idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=idx, column=2).value = excel_compatible(row[1])
    
    ws.cell(row=1, column=4, value="Top 10 brendova po ukupnoj vrednosti")
    ws.cell(row=2, column=4, value="Brend")
    ws.cell(row=2, column=5, value="Ukupna vrednost (RSD)")
    
    for idx, row in enumerate(brendovi_vred.itertuples(index=False), start=3):
        ws.cell(row=idx, column=4).value = excel_compatible(row[0])
        ws.cell(row=idx, column=5).value = excel_compatible(row[1])
    
    for col in [1, 2, 4, 5]:
        ws.column_dimensions[get_column_letter(col)].width = 25
    
    for row in [1, 2]:
        for col in [1, 2, 4, 5]:
            cell = ws.cell(row=row, column=col)
            if row == 1:
                cell.font = Font(bold=True, size=12, color=HEADER_BG)
            else:
                cell.font = Font(bold=True, size=11, color=HEADER_BG)

def upisi_top_grupe(ws, df):
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
    if "Datum obrade" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Datum obrade' ne postoji")
        return
    
    df_datum = df.copy()
    df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
    df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
    
    trend = df_datum.groupby("Mesec").size().reset_index()
    trend.columns = ["Mesec", "Broj artikala"]
    trend["Mesec"] = trend["Mesec"].astype(str)
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df_datum["Ukupna vrednost"] = df_datum["Nabavna cena"] * df_datum["Količina"]
        vrednost_po_mesecu = df_datum.groupby("Mesec")["Ukupna vrednost"].sum().reset_index()
        vrednost_po_mesecu.columns = ["Mesec", "Ukupna vrednost (RSD)"]
        vrednost_po_mesecu["Mesec"] = vrednost_po_mesecu["Mesec"].astype(str)
        trend = trend.merge(vrednost_po_mesecu, on="Mesec", how="left")
    
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
    if "prevoznik" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'prevoznik' ne postoji")
        return
    
    prevoznici = df["prevoznik"].value_counts().reset_index()
    prevoznici.columns = ["Prevoznik", "Broj artikala"]
    
    if "Klasifikacija štete" in df.columns:
        stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
        stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
    
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
    
    for row_idx, row in enumerate(prevoznici.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
    primeni_stil_podaci(ws, 2, len(prevoznici) + 1)
    
    if "Klasifikacija štete" in df.columns:
        for row_idx, row in enumerate(stete_prevoz.itertuples(index=False), start=3):
            ws.cell(row=row_idx, column=4).value = excel_compatible(row[0])
            ws.cell(row=row_idx, column=5).value = excel_compatible(row[1])
            ws.cell(row=row_idx, column=6).value = excel_compatible(row[2])
        primeni_stil_podaci(ws, 6, len(stete_prevoz) + 3)
    
    for col in [1, 2, 4, 5, 6]:
        ws.column_dimensions[get_column_letter(col)].width = 25

def upisi_metrike(ws, df):
    """Dodatni sheet sa naprednim metrikama"""
    ws.cell(row=1, column=1, value="Napredne metrike")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color=HEADER_BG)
    
    row = 3
    
    # 1. Pareto analiza (80/20)
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        pareto = df.sort_values("Ukupna vrednost", ascending=False)
        pareto["Kumulativni procenat"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
        
        ws.cell(row=row, column=1, value="PARETO ANALIZA (80/20)")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1
        
        headers = ["Redni broj", "Broj reklamacije", "Naziv artikla", "Vrednost (RSD)", "Kumulativni %", "Datum obrade", "Klasifikacija štete"]
        for col_idx, h in enumerate(headers, start=1):
            ws.cell(row=row, column=col_idx, value=h)
        primeni_stil_header(ws, len(headers))
        row += 1
        
        for idx, (_, r) in enumerate(pareto.head(20).iterrows(), 1):
            ws.cell(row=row, column=1).value = idx
            ws.cell(row=row, column=2).value = excel_compatible(r["Broj reklamacije"])
            ws.cell(row=row, column=3).value = excel_compatible(r["Naziv artikla"])
            ws.cell(row=row, column=4).value = excel_compatible(r["Ukupna vrednost"])
            ws.cell(row=row, column=5).value = excel_compatible(r["Kumulativni procenat"])
            ws.cell(row=row, column=6).value = excel_compatible(r["Datum obrade"]) if "Datum obrade" in r else None
            ws.cell(row=row, column=7).value = excel_compatible(r["Klasifikacija štete"]) if "Klasifikacija štete" in r else None
            row += 1
        
        granica = pareto[pareto["Kumulativni procenat"] <= 80].shape[0]
        ws.cell(row=row + 1, column=1, value=f"📊 {granica} artikala čini 80% ukupne vrednosti")
        ws.cell(row=row + 1, column=1).font = Font(bold=True, size=11, color="0066CC")
        row += 3
    
    # 2. Starost artikala
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        danas = datetime.now()
        df_datum["Starost (dani)"] = (danas - df_datum["Datum obrade"]).dt.days
        
        ws.cell(row=row, column=1, value="STAROST ARTIKALA")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1
        
        ws.cell(row=row, column=1, value="Prosečna starost (dani)")
        ws.cell(row=row, column=2, value=excel_compatible(df_datum["Starost (dani)"].mean().round(0)))
        row += 1
        
        ws.cell(row=row, column=1, value="Najstariji artikal (dani)")
        ws.cell(row=row, column=2, value=excel_compatible(df_datum["Starost (dani)"].max()))
        row += 1
        
        ws.cell(row=row, column=1, value="Najmlađi artikal (dani)")
        ws.cell(row=row, column=2, value=excel_compatible(df_datum["Starost (dani)"].min()))
        row += 2
        
        ws.cell(row=row, column=1, value="Top 10 najstarijih artikala")
        ws.cell(row=row, column=1).font = Font(bold=True, size=11)
        row += 1
        
        headers = ["Naziv artikla", "Starost (dani)", "Broj reklamacije"]
        for col_idx, h in enumerate(headers, start=1):
            ws.cell(row=row, column=col_idx, value=h)
        primeni_stil_header(ws, len(headers))
        row += 1
        
        najstariji = df_datum.nlargest(10, "Starost (dani)")[["Naziv artikla", "Starost (dani)", "Broj reklamacije"]]
        for _, r in najstariji.iterrows():
            ws.cell(row=row, column=1).value = excel_compatible(r["Naziv artikla"])
            ws.cell(row=row, column=2).value = excel_compatible(r["Starost (dani)"])
            ws.cell(row=row, column=3).value = excel_compatible(r["Broj reklamacije"])
            row += 1
        row += 2
    
    # 3. Top 10 najskupljih
    if "Nabavna cena" in df.columns:
        ws.cell(row=row, column=1, value="TOP 10 NAJSKUPLJIH ARTIKALA")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1
        
        headers = ["Broj reklamacije", "Naziv artikla", "Nabavna cena (RSD)", "Brend", "Datum obrade", "Klasifikacija reciklaže", "Klasifikacija štete"]
        for col_idx, h in enumerate(headers, start=1):
            ws.cell(row=row, column=col_idx, value=h)
        primeni_stil_header(ws, len(headers))
        row += 1
        
        najskuplji = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend", "Datum obrade", "Klasifikacija reciklaže", "Klasifikacija štete"]]
        for _, r in najskuplji.iterrows():
            ws.cell(row=row, column=1).value = excel_compatible(r["Broj reklamacije"])
            ws.cell(row=row, column=2).value = excel_compatible(r["Naziv artikla"])
            ws.cell(row=row, column=3).value = excel_compatible(r["Nabavna cena"])
            ws.cell(row=row, column=4).value = excel_compatible(r["Brend"])
            ws.cell(row=row, column=5).value = excel_compatible(r["Datum obrade"]) if "Datum obrade" in r else None
            ws.cell(row=row, column=6).value = excel_compatible(r["Klasifikacija reciklaže"]) if "Klasifikacija reciklaže" in r else None
            ws.cell(row=row, column=7).value = excel_compatible(r["Klasifikacija štete"]) if "Klasifikacija štete" in r else None
            row += 1
        row += 2
    
    # 4. Detekcija duplikata
    if "Serijski broj" in df.columns:
        duplikati = df[df["Serijski broj"].duplicated(keep=False)]
        ws.cell(row=row, column=1, value="DETEKCIJA DUPLIKATA")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1
        
        if len(duplikati) > 0:
            ws.cell(row=row, column=1, value=f"⚠️ Pronađeno {len(duplikati)} duplih serijskih brojeva")
            ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="CC0000")
            row += 1
            
            headers = ["Serijski broj", "Naziv artikla", "Broj pojavljivanja", "Brojevi reklamacija"]
            for col_idx, h in enumerate(headers, start=1):
                ws.cell(row=row, column=col_idx, value=h)
            primeni_stil_header(ws, len(headers))
            row += 1
            
            for serijski, group in duplikati.groupby("Serijski broj"):
                ws.cell(row=row, column=1).value = excel_compatible(serijski)
                ws.cell(row=row, column=2).value = excel_compatible(group["Naziv artikla"].iloc[0])
                ws.cell(row=row, column=3).value = len(group)
                brojevi = ", ".join([str(b) for b in group["Broj reklamacije"].tolist()])
                ws.cell(row=row, column=4).value = brojevi
                row += 1
        else:
            ws.cell(row=row, column=1, value="✅ Nema duplih serijskih brojeva")
            ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="008800")
        row += 2
    
    # 5. Prosečna vrednost po brendu
    if "Brend" in df.columns and "Nabavna cena" in df.columns:
        ws.cell(row=row, column=1, value="PROSEČNA VREDNOST PO BRENDU")
        ws.cell(row=row, column=1).font = Font(bold=True, size=12)
        row += 1
        
        prosek_brend = df.groupby("Brend")["Nabavna cena"].mean().sort_values(ascending=False).head(10).reset_index()
        prosek_brend.columns = ["Brend", "Prosečna cena (RSD)"]
        
        ws.cell(row=row, column=1, value="Brend")
        ws.cell(row=row, column=2, value="Prosečna cena (RSD)")
        primeni_stil_header(ws, 2)
        row += 1
        
        for _, r in prosek_brend.iterrows():
            ws.cell(row=row, column=1).value = excel_compatible(r["Brend"])
            ws.cell(row=row, column=2).value = excel_compatible(r["Prosečna cena (RSD)"])
            row += 1
        row += 2
    
    # Širine kolona
    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = 30

# ============================
# GLAVNA FUNKCIJA ZA EXCEL
# ============================

def formatiraj_i_sacuvaj(df):
    wb = Workbook()
    
    ws = wb.active
    ws.title = "zbirno"
    upisi_zbirno(ws, df)
    
    ws = wb.create_sheet("KPI")
    upisi_kpi(ws, df)
    
    ws = wb.create_sheet("Po klasifikaciji")
    upisi_grupisano(ws, df, "Klasifikacija reciklaže", "Klasifikacija reciklaže")
    
    ws = wb.create_sheet("Po šteti")
    upisi_grupisano(ws, df, "Klasifikacija štete", "Klasifikacija štete")
    
    ws = wb.create_sheet("Top brendovi")
    upisi_top_brendove(ws, df)
    
    ws = wb.create_sheet("Top grupe")
    upisi_top_grupe(ws, df)
    
    ws = wb.create_sheet("Vremenski trend")
    upisi_vremenski_trend(ws, df)
    
    ws = wb.create_sheet("Po prevozniku")
    upisi_prevoznike(ws, df)
    
    ws = wb.create_sheet("Napredne metrike")
    upisi_metrike(ws, df)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================
# PDF GENERATOR (sa srpskim slovima i grafikonima)
# ============================

def generisi_pdf(df):
    """Generiše profesionalan PDF izveštaj sa srpskim slovima i grafikonima"""
    
    # Registracija fonta za srpska slova
    font_registriran = False
    font_name = 'Helvetica'
    
    # Pokušaj da učita DejaVu font iz root foldera
    try:
        dir_path = os.path.dirname(__file__)
        font_path = os.path.join(dir_path, 'DejaVuSans.ttf')
        font_bold_path = os.path.join(dir_path, 'DejaVuSans-Bold.ttf')
        
        if os.path.exists(font_path) and os.path.exists(font_bold_path):
            pdfmetrics.registerFont(TTFont('DejaVu', font_path))
            pdfmetrics.registerFont(TTFont('DejaVu-Bold', font_bold_path))
            font_name = 'DejaVu'
            font_registriran = True
    except:
        pass
    
    # Ako nema DejaVu, probaj sistemski font
    if not font_registriran:
        try:
            pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
            pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
            font_name = 'DejaVu'
            font_registriran = True
        except:
            pass
    
    # Priprema podataka
    ukupno_artikala = df["Količina"].sum() if "Količina" in df.columns else 0
    ukupna_vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    broj_brendova = df["Brend"].nunique() if "Brend" in df.columns else 0
    broj_grupa = df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0
    
    # Kreiranje PDF-a
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=A4, 
        rightMargin=1.5*cm, 
        leftMargin=1.5*cm, 
        topMargin=1.5*cm, 
        bottomMargin=1.5*cm
    )
    
    styles = getSampleStyleSheet()
    
    # Stilovi sa srpskim fontom
    naslov_style = ParagraphStyle(
        'Naslov',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=24,
        textColor=colors.HexColor('#1F4E79'),
        alignment=1,
        spaceAfter=20,
        spaceBefore=30
    )
    
    naslov_sekcije_style = ParagraphStyle(
        'NaslovSekcije',
        parent=styles['Heading3'],
        fontName=font_name,
        fontSize=14,
        textColor=colors.HexColor('#2E75B6'),
        spaceAfter=8,
        spaceBefore=12
    )
    
    normal_style = ParagraphStyle(
        'Normal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=10,
        spaceAfter=4
    )
    
    bold_style = ParagraphStyle(
        'Bold',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=10,
        textColor=colors.HexColor('#1F4E79'),
        spaceAfter=4
    )
    
    elements = []
    
    # ============================================================
    # NASLOVNA STRANA
    # ============================================================
    
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("Tehnomanija", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph("IZVEŠTAJ O RECIKLAŽI", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
    elements.append(Spacer(1, 2*cm))
    
    # KPI
    kpi_data = [
        ["Ukupno artikala", f"{ukupno_artikala:,}"],
        ["Ukupna vrednost", f"{ukupna_vrednost:,.2f} RSD"],
        ["Broj brendova", str(broj_brendova)],
        ["Broj robnih grupa", str(broj_grupa)]
    ]
    kpi_table = Table(kpi_data, colWidths=[7*cm, 5*cm])
    kpi_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 11),
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (0, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#CCCCCC')),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (1, 0), (1, -1), colors.HexColor('#F5F8FD')),
    ]))
    elements.append(kpi_table)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 1. KLASIFIKACIJA RECIKLAŽE
    # ============================================================
    
    elements.append(Paragraph("1. Klasifikacija reciklaže", naslov_sekcije_style))
    
    if "Klasifikacija reciklaže" in df.columns:
        klas_data = df["Klasifikacija reciklaže"].value_counts().head(10).reset_index()
        klas_data.columns = ["Klasifikacija", "Broj artikala"]
        
        table_data = [["Klasifikacija", "Broj"]]
        for _, row in klas_data.iterrows():
            table_data.append([str(row["Klasifikacija"])[:50], str(row["Broj artikala"])])
        
        klas_table = Table(table_data, colWidths=[10*cm, 4*cm])
        klas_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(klas_table)
    
    elements.append(Spacer(1, 0.5*cm))
    
    # Grafikon
    if "Klasifikacija reciklaže" in df.columns:
        fig, ax = plt.subplots(figsize=(8, 5))
        klas_data = df["Klasifikacija reciklaže"].value_counts().head(8)
        colors_pie = plt.cm.Blues_r([i/len(klas_data) for i in range(len(klas_data))])
        wedges, texts, autotexts = ax.pie(
            klas_data.values, 
            labels=klas_data.index, 
            autopct='%1.0f%%',
            colors=colors_pie,
            startangle=90
        )
        ax.set_title('Raspored po klasifikaciji reciklaže', fontsize=14, fontweight='bold')
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=10*cm)
        elements.append(img)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 2. KLASIFIKACIJA ŠTETE
    # ============================================================
    
    elements.append(Paragraph("2. Klasifikacija štete", naslov_sekcije_style))
    
    if "Klasifikacija štete" in df.columns:
        steta_data = df["Klasifikacija štete"].value_counts().head(10).reset_index()
        steta_data.columns = ["Klasifikacija štete", "Broj artikala"]
        
        table_data = [["Klasifikacija štete", "Broj"]]
        for _, row in steta_data.iterrows():
            table_data.append([str(row["Klasifikacija štete"])[:50], str(row["Broj artikala"])])
        
        steta_table = Table(table_data, colWidths=[10*cm, 4*cm])
        steta_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(steta_table)
    
    elements.append(Spacer(1, 0.5*cm))
    
    # Grafikon
    if "Klasifikacija štete" in df.columns:
        fig, ax = plt.subplots(figsize=(8, 5))
        steta_data = df["Klasifikacija štete"].value_counts().head(8)
        ax.barh(steta_data.index, steta_data.values, color=plt.cm.Reds_r([i/len(steta_data) for i in range(len(steta_data))]))
        ax.set_xlabel('Broj artikala')
        ax.set_title('Najčešće klasifikacije štete', fontsize=14, fontweight='bold')
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=10*cm)
        elements.append(img)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 3. TOP 10 BRENDOVA
    # ============================================================
    
    elements.append(Paragraph("3. Top 10 brendova", naslov_sekcije_style))
    
    if "Brend" in df.columns:
        brend_data = df["Brend"].value_counts().head(10).reset_index()
        brend_data.columns = ["Brend", "Broj artikala"]
        
        table_data = [["Brend", "Broj artikala"]]
        for _, row in brend_data.iterrows():
            table_data.append([str(row["Brend"]), str(row["Broj artikala"])])
        
        brend_table = Table(table_data, colWidths=[10*cm, 4*cm])
        brend_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(brend_table)
    
    elements.append(Spacer(1, 0.5*cm))
    
    # Grafikon
    if "Brend" in df.columns:
        fig, ax = plt.subplots(figsize=(8, 5))
        brend_data = df["Brend"].value_counts().head(10)
        ax.bar(brend_data.index, brend_data.values, color=plt.cm.Blues([0.3 + 0.7*i/len(brend_data) for i in range(len(brend_data))]))
        ax.set_ylabel('Broj artikala')
        ax.set_title('Top 10 brendova po broju artikala', fontsize=14, fontweight='bold')
        plt.xticks(rotation=45, ha='right')
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=8*cm)
        elements.append(img)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 4. VREMENSKI TREND
    # ============================================================
    
    elements.append(Paragraph("4. Vremenski trend", naslov_sekcije_style))
    
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
        trend_data = df_datum.groupby("Mesec").size().reset_index()
        trend_data.columns = ["Mesec", "Broj artikala"]
        trend_data["Mesec"] = trend_data["Mesec"].astype(str)
        
        table_data = [["Mesec", "Broj artikala"]]
        for _, row in trend_data.iterrows():
            table_data.append([str(row["Mesec"]), str(row["Broj artikala"])])
        
        trend_table = Table(table_data, colWidths=[8*cm, 6*cm])
        trend_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(trend_table)
        
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(trend_data["Mesec"], trend_data["Broj artikala"], marker='o', linewidth=2, color='#1F4E79')
        ax.fill_between(trend_data["Mesec"], trend_data["Broj artikala"], alpha=0.3, color='#1F4E79')
        ax.set_ylabel('Broj artikala')
        ax.set_title('Broj artikala po mesecima', fontsize=14, fontweight='bold')
        ax.grid(True, linestyle='--', alpha=0.5)
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=6*cm)
        elements.append(img)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 5. PARETO ANALIZA
    # ============================================================
    
    elements.append(Paragraph("5. Pareto analiza (80/20)", naslov_sekcije_style))
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        pareto = df.sort_values("Ukupna vrednost", ascending=False)
        pareto["Kumulativni procenat"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
        
        table_data = [["Broj reklamacije", "Naziv artikla", "Vrednost (RSD)", "Kum.%"]]
        for _, row in pareto.head(15).iterrows():
            table_data.append([
                str(row["Broj reklamacije"])[:15],
                str(row["Naziv artikla"])[:30],
                f"{row['Ukupna vrednost']:,.0f}",
                f"{row['Kumulativni procenat']}%"
            ])
        
        pareto_table = Table(table_data, colWidths=[3*cm, 6*cm, 3*cm, 2.5*cm])
        pareto_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        elements.append(pareto_table)
        
        elements.append(Spacer(1, 0.3*cm))
        
        granica = pareto[pareto["Kumulativni procenat"] <= 80].shape[0]
        elements.append(Paragraph(f"📊 {granica} artikala čini 80% ukupne vrednosti", bold_style))
        
        # Grafikon
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.bar(range(20), pareto["Ukupna vrednost"].head(20), color=plt.cm.Blues([0.3 + 0.7*i/20 for i in range(20)]))
        ax.axhline(y=pareto["Ukupna vrednost"].sum() * 0.8, color='red', linestyle='--', linewidth=2, label='80% granica')
        ax.set_xlabel('Redni broj artikla')
        ax.set_ylabel('Vrednost (RSD)')
        ax.set_title('Pareto analiza - Top 20 artikala', fontsize=14, fontweight='bold')
        ax.legend()
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=8*cm)
        elements.append(img)
    
    elements.append(PageBreak())
    
    # ============================================================
    # 6. TOP 10 NAJSKUPLJIH
    # ============================================================
    
    elements.append(Paragraph("6. Top 10 najskupljih artikala", naslov_sekcije_style))
    
    if "Nabavna cena" in df.columns:
        najskuplji = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend"]]
        
        table_data = [["Broj reklamacije", "Naziv artikla", "Cena (RSD)", "Brend"]]
        for _, row in najskuplji.iterrows():
            table_data.append([
                str(row["Broj reklamacije"])[:15],
                str(row["Naziv artikla"])[:35],
                f"{row['Nabavna cena']:,.2f}",
                str(row["Brend"])
            ])
        
        najskuplji_table = Table(table_data, colWidths=[3*cm, 6*cm, 3*cm, 3*cm])
        najskuplji_table.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
        ]))
        elements.append(najskuplji_table)
        
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.barh(najskuplji["Naziv artikla"].str[:20], najskuplji["Nabavna cena"], color=plt.cm.Greens([0.3 + 0.7*i/10 for i in range(10)]))
        ax.set_xlabel('Cena (RSD)')
        ax.set_title('Top 10 najskupljih artikala', fontsize=14, fontweight='bold')
        
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        
        img = Image(img_buffer, width=15*cm, height=8*cm)
        elements.append(img)
    
    # ============================================================
    # ZAVRŠNA STRANA
    # ============================================================
    
    elements.append(PageBreak())
    elements.append(Spacer(1, 5*cm))
    elements.append(Paragraph("Izveštaj generisan automatski", normal_style))
    elements.append(Paragraph(f"Tehnomanija © {datetime.now().year}", normal_style))
    
    # Generisanje PDF-a
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ============================
# EMAIL SLANJE
# ============================

def posalji_email(primalac, excel_data, pdf_data=None, dodatne_adrese=None):
    """Šalje email sa Excel i opciono PDF prilogom"""
    try:
        # Konfiguracija iz secrets
        sender = st.secrets["email"]["sender"]
        password = st.secrets["email"]["password"]
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        
        # Kreiranje poruke
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = primalac
        msg["Subject"] = f"Izveštaj o reciklaži – {datetime.now().strftime('%d.%m.%Y')}"
        
        # Dodatne adrese (CC)
        if dodatne_adrese:
            msg["CC"] = ", ".join(dodatne_adrese)
        
        # Telo poruke
        body = f"""
Poštovani,

U prilogu vam dostavljamo izveštaj o reciklaži sa stanjem od {datetime.now().strftime('%d.%m.%Y')}.

Izveštaj sadrži:
- Kompletnu tabelu sa svim podacima
- Ključne pokazatelje (KPI)
- Analize po klasifikacijama, brendovima i robnim grupama
- Vremenski trend i analizu prevoznika
- Napredne metrike (Pareto, starost, detekcija duplikata)

Izveštaj je generisan automatski.

S poštovanjem,
Tehnomanija
        """
        msg.attach(MIMEText(body, "plain"))
        
        # Excel prilog
        excel_part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        excel_part.set_payload(excel_data)
        encoders.encode_base64(excel_part)
        excel_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.xlsx')
        msg.attach(excel_part)
        
        # PDF prilog (ako postoji)
        if pdf_data:
            pdf_part = MIMEBase('application', 'pdf')
            pdf_part.set_payload(pdf_data)
            encoders.encode_base64(pdf_part)
            pdf_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.pdf')
            msg.attach(pdf_part)
        
        # Slanje
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender, password)
        
        # Primaoci (TO + CC)
        svi_primaoci = [primalac]
        if dodatne_adrese:
            svi_primaoci.extend(dodatne_adrese)
        
        server.send_message(msg)
        server.quit()
        return True, "Email je uspešno poslat!"
    
    except Exception as e:
        return False, f"Greška pri slanju emaila: {str(e)}"

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

def prikazi_klasifikaciju_reciklaze(df, key_suffix=""):
    st.subheader("📊 Klasifikacija reciklaže")
    
    if "Klasifikacija reciklaže" in df.columns:
        klasifikacija = df["Klasifikacija reciklaže"].value_counts().reset_index()
        klasifikacija.columns = ["Klasifikacija", "Broj artikala"]
        
        fig = px.pie(klasifikacija, values="Broj artikala", names="Klasifikacija", 
                     title="Raspored po klasifikaciji reciklaže",
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        fig.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig, use_container_width=True, key=f"klas_rec_{key_suffix}")
        
        st.dataframe(klasifikacija, use_container_width=True)

def prikazi_klasifikaciju_stete(df, key_suffix=""):
    st.subheader("📊 Klasifikacija štete")
    
    if "Klasifikacija štete" in df.columns:
        stete = df["Klasifikacija štete"].value_counts().reset_index()
        stete.columns = ["Klasifikacija štete", "Broj artikala"]
        
        fig = px.bar(stete, x="Klasifikacija štete", y="Broj artikala",
                     title="Najčešće klasifikacije štete",
                     color="Broj artikala", color_continuous_scale="Blues")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True, key=f"steta_{key_suffix}")
        
        st.dataframe(stete, use_container_width=True)

def prikazi_top_brendove(df, key_suffix=""):
    st.subheader("🏆 Top 10 brendova")
    
    if "Brend" in df.columns:
        col1, col2 = st.columns(2)
        
        with col1:
            brendovi = df["Brend"].value_counts().head(10).reset_index()
            brendovi.columns = ["Brend", "Broj artikala"]
            fig = px.bar(brendovi, x="Brend", y="Broj artikala",
                         title="Top 10 brendova po broju artikala",
                         color="Broj artikala", color_continuous_scale="Blues")
            st.plotly_chart(fig, use_container_width=True, key=f"brend_broj_{key_suffix}")
        
        with col2:
            if "Nabavna cena" in df.columns and "Količina" in df.columns:
                df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
                vrednost_brend = df.groupby("Brend")["Ukupna vrednost"].sum().sort_values(ascending=False).head(10).reset_index()
                vrednost_brend.columns = ["Brend", "Ukupna vrednost (RSD)"]
                fig2 = px.bar(vrednost_brend, x="Brend", y="Ukupna vrednost (RSD)",
                              title="Top 10 brendova po ukupnoj vrednosti",
                              color="Ukupna vrednost (RSD)", color_continuous_scale="Greens")
                st.plotly_chart(fig2, use_container_width=True, key=f"brend_vred_{key_suffix}")

def prikazi_robne_grupe(df, key_suffix=""):
    st.subheader("📂 Top 10 robnih grupa")
    
    if "Robna grupa" in df.columns:
        grupe = df["Robna grupa"].value_counts().head(10).reset_index()
        grupe.columns = ["Robna grupa", "Broj artikala"]
        
        fig = px.bar(grupe, x="Robna grupa", y="Broj artikala",
                     title="Top 10 robnih grupa",
                     color="Broj artikala", color_continuous_scale="Reds")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True, key=f"grupa_{key_suffix}")

def prikazi_vremenski_trend(df, key_suffix=""):
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
        st.plotly_chart(fig, use_container_width=True, key=f"trend_{key_suffix}")

def prikazi_prevoznike(df, key_suffix=""):
    st.subheader("🚚 Analiza prevoznika")
    
    if "prevoznik" in df.columns:
        col1, col2 = st.columns(2)
        
        with col1:
            prevoznici = df["prevoznik"].value_counts().reset_index()
            prevoznici.columns = ["Prevoznik", "Broj artikala"]
            fig = px.pie(prevoznici, values="Broj artikala", names="Prevoznik",
                         title="Raspored po prevoznicima")
            st.plotly_chart(fig, use_container_width=True, key=f"prevoz_pie_{key_suffix}")
        
        with col2:
            if "Klasifikacija štete" in df.columns:
                stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
                stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
                fig2 = px.bar(stete_prevoz, x="Prevoznik", y="Broj", color="Klasifikacija štete",
                              title="Klasifikacije štete po prevoznicima",
                              barmode="group")
                st.plotly_chart(fig2, use_container_width=True, key=f"prevoz_bar_{key_suffix}")

def prikazi_napredne_metrike(df):
    st.subheader("📊 Napredne metrike")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 Pareto analiza",
        "⏳ Starost artikala",
        "💰 Top 10 najskupljih",
        "🔄 Detekcija duplikata",
        "📊 Prosečna vrednost po brendu"
    ])
    
    with tab1:
        if "Nabavna cena" in df.columns and "Količina" in df.columns:
            df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
            pareto = df.sort_values("Ukupna vrednost", ascending=False)
            pareto["Kumulativni procenat"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
            
            fig = px.bar(pareto.head(20), x="Naziv artikla", y="Ukupna vrednost",
                         title="Pareto analiza – Top 20 artikala po vrednosti",
                         hover_data=["Broj reklamacije"])
            fig.add_hline(y=pareto["Ukupna vrednost"].sum() * 0.8, line_dash="dash", line_color="red")
            st.plotly_chart(fig, use_container_width=True, key="pareto_chart")
            
            granica = pareto[pareto["Kumulativni procenat"] <= 80].shape[0]
            st.info(f"📊 **{granica}** artikala čini **80%** ukupne vrednosti reciklaže")
            
            pareto_display = pareto.head(20)[["Broj reklamacije", "Naziv artikla", "Ukupna vrednost", "Kumulativni procenat", "Datum obrade", "Klasifikacija štete"]]
            st.dataframe(pareto_display, use_container_width=True)
    
    with tab2:
        if "Datum obrade" in df.columns:
            df_datum = df.copy()
            df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
            danas = datetime.now()
            df_datum["Starost (dani)"] = (danas - df_datum["Datum obrade"]).dt.days
            
            col1, col2, col3 = st.columns(3)
            col1.metric("📅 Prosečna starost", f"{df_datum['Starost (dani)'].mean().round(0):.0f} dana")
            col2.metric("⏳ Najstariji artikal", f"{df_datum['Starost (dani)'].max():.0f} dana")
            col3.metric("🆕 Najmlađi artikal", f"{df_datum['Starost (dani)'].min():.0f} dana")
            
            fig = px.histogram(df_datum, x="Starost (dani)", nbins=20,
                               title="Distribucija starosti artikala")
            st.plotly_chart(fig, use_container_width=True, key="starost_hist")
            
            st.subheader("Top 10 najstarijih artikala")
            st.dataframe(df_datum.nlargest(10, "Starost (dani)")[["Naziv artikla", "Starost (dani)", "Brend", "Broj reklamacije"]], use_container_width=True)
    
    with tab3:
        if "Nabavna cena" in df.columns:
            najskuplji = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend", "Robna grupa", "Datum obrade", "Klasifikacija reciklaže", "Klasifikacija štete"]]
            
            fig = px.bar(najskuplji, x="Naziv artikla", y="Nabavna cena", color="Brend",
                         title="Top 10 najskupljih artikala",
                         hover_data=["Broj reklamacije"])
            st.plotly_chart(fig, use_container_width=True, key="top_skupih")
            
            st.dataframe(najskuplji, use_container_width=True)
    
    with tab4:
        if "Serijski broj" in df.columns:
            duplikati = df[df["Serijski broj"].duplicated(keep=False)]
            if len(duplikati) > 0:
                st.warning(f"⚠️ Pronađeno **{len(duplikati)}** duplih serijskih brojeva")
                
                duplikati_sum = duplikati.groupby("Serijski broj").agg({
                    "Naziv artikla": "first",
                    "Broj reklamacije": lambda x: list(x)
                }).reset_index()
                duplikati_sum.columns = ["Serijski broj", "Naziv artikla", "Brojevi reklamacija"]
                duplikati_sum["Broj pojavljivanja"] = duplikati_sum["Brojevi reklamacija"].apply(len)
                st.dataframe(duplikati_sum, use_container_width=True)
            else:
                st.success("✅ Nema duplih serijskih brojeva")
    
    with tab5:
        if "Brend" in df.columns and "Nabavna cena" in df.columns:
            prosek_brend = df.groupby("Brend")["Nabavna cena"].mean().sort_values(ascending=False).head(10).reset_index()
            prosek_brend.columns = ["Brend", "Prosečna cena (RSD)"]
            
            fig = px.bar(prosek_brend, x="Brend", y="Prosečna cena (RSD)",
                         title="Top 10 brendova po prosečnoj ceni artikla")
            st.plotly_chart(fig, use_container_width=True, key="prosek_brend")
            
            st.dataframe(prosek_brend, use_container_width=True)

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

# Sidebar
with st.sidebar:
    st.title("⚙️ Opcije")
    
    # Tamna tema
    dark_mode = st.toggle("🌙 Tamna tema", value=st.session_state.dark_mode)
    if dark_mode != st.session_state.dark_mode:
        st.session_state.dark_mode = dark_mode
        st.rerun()
    
    st.markdown("---")
    st.info("📊 **Izveštaj o reciklaži**\n\nVerzija 3.0\n\n© Tehnomanija")
    
    if "df" in st.session_state and st.session_state.df is not None:
        st.markdown("---")
        st.caption(f"📅 Poslednji upload: {datetime.now().strftime('%H:%M:%S')}")
        st.caption(f"📊 {len(st.session_state.df)} redova podataka")

# Glavni sadržaj
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
            st.session_state.df = df
        
        st.success(f"✅ Učitano: {len(reciklaza)} redova u 'reciklaža', {len(reklamacije)} redova u 'reklamacije ukupno'")
        st.info(f"📊 Spojeno: {len(df)} redova")
        
        # ======================
        # TABOVI ZA IZVEŠTAJE
        # ======================
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "📊 Pregled",
            "📈 Napredne metrike",
            "🏷️ Analize",
            "📋 Tabela",
            "📤 Izvoz"
        ])
        
        with tab1:
            prikazi_kpi(df)
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            with col1:
                prikazi_klasifikaciju_reciklaze(df, key_suffix="pregled")
            with col2:
                prikazi_klasifikaciju_stete(df, key_suffix="pregled")
            st.markdown("---")
            
            prikazi_top_brendove(df, key_suffix="pregled")
            st.markdown("---")
            
            col1, col2 = st.columns(2)
            with col1:
                prikazi_robne_grupe(df, key_suffix="pregled")
            with col2:
                prikazi_vremenski_trend(df, key_suffix="pregled")
            st.markdown("---")
            
            prikazi_prevoznike(df, key_suffix="pregled")
        
        with tab2:
            prikazi_napredne_metrike(df)
        
        with tab3:
            st.subheader("🏷️ Sve analize")
            prikazi_klasifikaciju_reciklaze(df, key_suffix="analize")
            st.markdown("---")
            prikazi_klasifikaciju_stete(df, key_suffix="analize")
            st.markdown("---")
            prikazi_top_brendove(df, key_suffix="analize")
            st.markdown("---")
            prikazi_robne_grupe(df, key_suffix="analize")
            st.markdown("---")
            prikazi_vremenski_trend(df, key_suffix="analize")
            st.markdown("---")
            prikazi_prevoznike(df, key_suffix="analize")
        
        with tab4:
            prikazi_detaljnu_tabelu(df)
        
        with tab5:
            st.subheader("📤 Izvoz izveštaja")
            
            # Excel
            st.markdown("### 📊 Excel izveštaj")
            st.info("""
            📋 **Excel fajl sadrži 9 sheet-ova:**
            1. zbirno – svi podaci
            2. KPI – ključni pokazatelji
            3. Po klasifikaciji – analiza reciklaže
            4. Po šteti – analiza štete
            5. Top brendovi – top 10 brendova
            6. Top grupe – top 10 robnih grupa
            7. Vremenski trend – mesečni pregled
            8. Po prevozniku – analiza prevoznika
            9. Napredne metrike – Pareto, starost, duplikati
            """)
            
            col1, col2 = st.columns(2)
            with col1:
                excel_data = formatiraj_i_sacuvaj(df)
                st.download_button(
                    label="📥 Preuzmi Excel izveštaj (9 sheet-ova)",
                    data=excel_data,
                    file_name=f"Izvestaj_o_reciklazi_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                # PDF
                with st.spinner("Generišem PDF..."):
                    pdf_data = generisi_pdf(df)
                    st.download_button(
                        label="📄 Preuzmi PDF izveštaj",
                        data=pdf_data,
                        file_name=f"Izvestaj_o_reciklazi_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
            
            st.markdown("---")
            
            # Email
            st.markdown("### 📧 Pošalji izveštaj emailom")
            
            with st.form("email_form"):
                col1, col2 = st.columns(2)
                with col1:
                    email_primalac = st.text_input("Email adresa (obavezno)", placeholder="office@tehnomanija.rs")
                with col2:
                    email_cc = st.text_input("Dodatne adrese (CC, odvojene zarezom)", placeholder="nikola@tehnomanija.rs, milica@tehnomanija.rs")
                
                col1, col2, col3 = st.columns([1, 1, 2])
                with col1:
                    posalji_pdf = st.checkbox("📄 Dodaj PDF prilog", value=True)
                with col2:
                    poslati_excel = st.checkbox("📊 Dodaj Excel prilog", value=True)
                
                submitted = st.form_submit_button("📩 Pošalji izveštaj")
                
                if submitted:
                    if not email_primalac:
                        st.error("Molimo unesite email adresu")
                    else:
                        # Parsiranje CC adresa
                        cc_list = []
                        if email_cc:
                            cc_list = [email.strip() for email in email_cc.split(",") if email.strip()]
                        
                        with st.spinner("Šaljem email..."):
                            # Priprema priloga
                            excel_attach = excel_data if poslati_excel else None
                            pdf_attach = pdf_data if posalji_pdf else None
                            
                            success, message = posalji_email(
                                email_primalac,
                                excel_attach,
                                pdf_attach,
                                cc_list
                            )
                            
                            if success:
                                st.success(f"✅ {message}")
                                if cc_list:
                                    st.info(f"📋 Kopija poslata na: {', '.join(cc_list)}")
                            else:
                                st.error(f"❌ {message}")
else:
    st.info("📂 Molimo uploadujte Excel fajl da biste započeli.")

# ============================
# FOOTER
# ============================

st.markdown("---")
st.caption(f"🔄 Izveštaj o reciklaži v3.0 | {datetime.now().strftime('%d.%m.%Y %H:%M')}")
