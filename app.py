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
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# ============================
# REGISTRACIJA FONTA
# ============================

FONT_NAME = 'Helvetica'
try:
    pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
    FONT_NAME = 'DejaVu'
except:
    try:
        pdfmetrics.registerFont(TTFont('DejaVu', 'DejaVuSans.ttf'))
        pdfmetrics.registerFont(TTFont('DejaVu-Bold', 'DejaVuSans-Bold.ttf'))
        FONT_NAME = 'DejaVu'
    except:
        FONT_NAME = 'Helvetica'

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
# OBRADA PODATAKA
# ============================

def cisti_broj(vrednost):
    if pd.isna(vrednost):
        return ""
    tekst = str(vrednost).strip()
    cifre = re.sub(r'[^0-9]', '', tekst)
    return cifre

def pronadji_kolonu_broja(df):
    for kol in df.columns:
        if "broj" in kol.lower() and "reklamacije" in kol.lower():
            return kol
        if "reklamacija" in kol.lower() and ("broj" in kol.lower() or "id" in kol.lower()):
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
    """Sheet: grupisanje po koloni sa brojem i vrednošću"""
    if kolona not in df.columns:
        ws.cell(row=1, column=1, value=f"Kolona '{kolona}' ne postoji")
        return
    
    # Grupisanje po broju
    grupisan = df[kolona].value_counts().reset_index()
    grupisan.columns = [kolona, "Broj artikala"]
    
    # Grupisanje po vrednosti
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        vrednost_grupisano = df.groupby(kolona)["Ukupna vrednost"].sum().reset_index()
        vrednost_grupisano.columns = [kolona, "Ukupna vrednost (RSD)"]
        grupisan = grupisan.merge(vrednost_grupisano, on=kolona, how="left")
    else:
        grupisan["Ukupna vrednost (RSD)"] = 0
    
    # Header
    ws.cell(row=1, column=1, value=kolona)
    ws.cell(row=1, column=2, value="Broj artikala")
    ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
    primeni_stil_header(ws, 3)
    
    # Podaci
    for row_idx, row in enumerate(grupisan.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        if len(row) > 2:
            ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    primeni_stil_podaci(ws, 3, len(grupisan) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25

def upisi_top_brendove(ws, df):
    """Sheet: Top 10 brendova sa brojem i vrednošću"""
    if "Brend" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Brend' ne postoji")
        return
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        brend_analiza = df.groupby("Brend").agg({
            "Broj reklamacije": "count",
            "Ukupna vrednost": "sum"
        }).reset_index()
        brend_analiza.columns = ["Brend", "Broj artikala", "Ukupna vrednost (RSD)"]
    else:
        brend_analiza = df["Brend"].value_counts().head(10).reset_index()
        brend_analiza.columns = ["Brend", "Broj artikala"]
        brend_analiza["Ukupna vrednost (RSD)"] = 0
    
    brend_analiza = brend_analiza.sort_values("Broj artikala", ascending=False).head(10)
    
    # Header
    ws.cell(row=1, column=1, value="Top 10 brendova")
    ws.cell(row=1, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    ws.cell(row=2, column=1, value="Brend")
    ws.cell(row=2, column=2, value="Broj artikala")
    ws.cell(row=2, column=3, value="Ukupna vrednost (RSD)")
    primeni_stil_header(ws, 3)
    
    for idx, row in enumerate(brend_analiza.itertuples(index=False), start=3):
        ws.cell(row=idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=idx, column=2).value = excel_compatible(row[1])
        ws.cell(row=idx, column=3).value = excel_compatible(row[2])
    primeni_stil_podaci(ws, 3, len(brend_analiza) + 2)
    
    for col in [1, 2, 3]:
        ws.column_dimensions[get_column_letter(col)].width = 25

def upisi_top_grupe(ws, df):
    """Sheet: Top 10 robnih grupa sa brojem i vrednošću"""
    if "Robna grupa" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Robna grupa' ne postoji")
        return
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        grupe = df.groupby("Robna grupa").agg({
            "Broj reklamacije": "count",
            "Ukupna vrednost": "sum"
        }).reset_index()
        grupe.columns = ["Robna grupa", "Broj artikala", "Ukupna vrednost (RSD)"]
    else:
        grupe = df["Robna grupa"].value_counts().head(10).reset_index()
        grupe.columns = ["Robna grupa", "Broj artikala"]
        grupe["Ukupna vrednost (RSD)"] = 0
    
    grupe = grupe.sort_values("Broj artikala", ascending=False).head(10)
    
    ws.cell(row=1, column=1, value="Robna grupa")
    ws.cell(row=1, column=2, value="Broj artikala")
    ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
    primeni_stil_header(ws, 3)
    
    for row_idx, row in enumerate(grupe.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    primeni_stil_podaci(ws, 3, len(grupe) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25

def upisi_vremenski_trend(ws, df):
    """Sheet: Vremenski trend sa brojem i vrednošću"""
    if "Datum obrade" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Datum obrade' ne postoji")
        return
    
    df_datum = df.copy()
    df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
    df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
    
    # Grupisanje po mesecima
    trend = df_datum.groupby("Mesec").agg({
        "Broj reklamacije": "count"
    }).reset_index()
    trend.columns = ["Mesec", "Broj artikala"]
    trend["Mesec"] = trend["Mesec"].astype(str)
    
    # Dodaj vrednost
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df_datum["Ukupna vrednost"] = df_datum["Nabavna cena"] * df_datum["Količina"]
        vrednost_po_mesecu = df_datum.groupby("Mesec")["Ukupna vrednost"].sum().reset_index()
        vrednost_po_mesecu.columns = ["Mesec", "Ukupna vrednost (RSD)"]
        vrednost_po_mesecu["Mesec"] = vrednost_po_mesecu["Mesec"].astype(str)
        trend = trend.merge(vrednost_po_mesecu, on="Mesec", how="left")
    
    # Header
    ws.cell(row=1, column=1, value="Mesec")
    ws.cell(row=1, column=2, value="Broj artikala")
    if "Ukupna vrednost (RSD)" in trend.columns:
        ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
        primeni_stil_header(ws, 3)
    else:
        primeni_stil_header(ws, 2)
    
    # Podaci
    for row_idx, row in enumerate(trend.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        if len(row) > 2:
            ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    
    primeni_stil_podaci(ws, len(trend.columns), len(trend) + 1)
    
    for col_idx in range(1, len(trend.columns) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 25

def upisi_prevoznike(ws, df):
    """Sheet: Po prevozniku sa brojem i vrednošću"""
    if "prevoznik" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'prevoznik' ne postoji")
        return
    
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        prevoznici = df.groupby("prevoznik").agg({
            "Broj reklamacije": "count",
            "Ukupna vrednost": "sum"
        }).reset_index()
        prevoznici.columns = ["Prevoznik", "Broj artikala", "Ukupna vrednost (RSD)"]
    else:
        prevoznici = df["prevoznik"].value_counts().reset_index()
        prevoznici.columns = ["Prevoznik", "Broj artikala"]
        prevoznici["Ukupna vrednost (RSD)"] = 0
    
    # Štete po prevozniku
    if "Klasifikacija štete" in df.columns:
        stete_prevoz = df.groupby(["prevoznik", "Klasifikacija štete"]).size().reset_index()
        stete_prevoz.columns = ["Prevoznik", "Klasifikacija štete", "Broj"]
    
    # Header - glavni
    ws.cell(row=1, column=1, value="Prevoznik")
    ws.cell(row=1, column=2, value="Broj artikala")
    ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
    ws.cell(row=1, column=5, value="Štete po prevozniku")
    ws.cell(row=2, column=5, value="Prevoznik")
    ws.cell(row=2, column=6, value="Klasifikacija štete")
    ws.cell(row=2, column=7, value="Broj")
    
    primeni_stil_header(ws, 3)
    for col in [5, 6, 7]:
        ws.cell(row=1, column=col).font = Font(bold=True, size=11, color=HEADER_BG)
        ws.cell(row=2, column=col).font = Font(bold=True, size=10, color=HEADER_BG)
    
    # Podaci - prevoznici
    for row_idx, row in enumerate(prevoznici.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    primeni_stil_podaci(ws, 3, len(prevoznici) + 1)
    
    # Podaci - štete
    if "Klasifikacija štete" in df.columns:
        for row_idx, row in enumerate(stete_prevoz.itertuples(index=False), start=3):
            ws.cell(row=row_idx, column=5).value = excel_compatible(row[0])
            ws.cell(row=row_idx, column=6).value = excel_compatible(row[1])
            ws.cell(row=row_idx, column=7).value = excel_compatible(row[2])
        primeni_stil_podaci(ws, 7, len(stete_prevoz) + 3)
    
    for col in [1, 2, 3, 5, 6, 7]:
        ws.column_dimensions[get_column_letter(col)].width = 25

def upisi_metrike(ws, df):
    """Sheet: Napredne metrike sa vrednostima"""
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
        
        headers = ["Naziv artikla", "Starost (dani)", "Broj reklamacije", "Vrednost (RSD)"]
        for col_idx, h in enumerate(headers, start=1):
            ws.cell(row=row, column=col_idx, value=h)
        primeni_stil_header(ws, len(headers))
        row += 1
        
        najstariji = df_datum.nlargest(10, "Starost (dani)")[["Naziv artikla", "Starost (dani)", "Broj reklamacije", "Ukupna vrednost"]]
        for _, r in najstariji.iterrows():
            ws.cell(row=row, column=1).value = excel_compatible(r["Naziv artikla"])
            ws.cell(row=row, column=2).value = excel_compatible(r["Starost (dani)"])
            ws.cell(row=row, column=3).value = excel_compatible(r["Broj reklamacije"])
            ws.cell(row=row, column=4).value = excel_compatible(r["Ukupna vrednost"]) if "Ukupna vrednost" in r else None
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
            
            headers = ["Serijski broj", "Naziv artikla", "Broj pojavljivanja", "Brojevi reklamacija", "Vrednost (RSD)"]
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
                ukupna_vrednost = (group["Nabavna cena"] * group["Količina"]).sum() if "Nabavna cena" in group.columns and "Količina" in group.columns else 0
                ws.cell(row=row, column=5).value = excel_compatible(ukupna_vrednost)
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
    
    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = 30

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
# PLOTLY -> SLIKA (ZA PDF)
# ============================

def plotly_to_reportlab_image(fig, width=12*cm, height=8*cm):
    """Konvertuje plotly figuru u ReportLab Image objekat"""
    img_bytes = fig.to_image(format="png", width=800, height=500, scale=1)
    img_buffer = io.BytesIO(img_bytes)
    img_buffer.seek(0)
    return Image(img_buffer, width=width, height=height)

# ============================
# PDF GENERATOR
# ============================

def generisi_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    elements = []
    
    naslov_style = ParagraphStyle('Naslov', parent=styles['Heading1'], fontName=FONT_NAME, fontSize=24, textColor=colors.HexColor('#1F4E79'), alignment=1, spaceAfter=20, spaceBefore=30)
    sekcija_style = ParagraphStyle('Sekcija', parent=styles['Heading3'], fontName=FONT_NAME, fontSize=14, textColor=colors.HexColor('#2E75B6'), spaceAfter=8, spaceBefore=12)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, spaceAfter=4)
    bold_style = ParagraphStyle('Bold', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, textColor=colors.HexColor('#1F4E79'), spaceAfter=4)
    
    ukupno = df["Količina"].sum() if "Količina" in df.columns else 0
    vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    broj_brendova = df["Brend"].nunique() if "Brend" in df.columns else 0
    
    # NASLOVNA
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("Tehnomanija", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph("IZVEŠTAJ O RECIKLAŽI", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
    elements.append(Spacer(1, 2*cm))
    
    kpi_data = [["Ukupno artikala", f"{ukupno:,}"], ["Ukupna vrednost", f"{vrednost:,.2f} RSD"], ["Broj brendova", str(broj_brendova)]]
    kpi_table = Table(kpi_data, colWidths=[7*cm, 5*cm])
    kpi_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 11), ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (0,-1), colors.white), ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#CCCCCC')), ('ALIGN', (1,0), (1,-1), 'RIGHT')]))
    elements.append(kpi_table)
    elements.append(PageBreak())
    
    # 1. KLASIFIKACIJA RECIKLAŽE
    elements.append(Paragraph("1. Klasifikacija reciklaže", sekcija_style))
    if "Klasifikacija reciklaže" in df.columns:
        klas = df["Klasifikacija reciklaže"].value_counts().head(10).reset_index()
        klas.columns = ["Klasifikacija", "Broj"]
        if "Nabavna cena" in df.columns and "Količina" in df.columns:
            df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
            vrednost_klas = df.groupby("Klasifikacija reciklaže")["Ukupna vrednost"].sum().reset_index()
            vrednost_klas.columns = ["Klasifikacija", "Vrednost (RSD)"]
            klas = klas.merge(vrednost_klas, on="Klasifikacija", how="left")
        
        table_data = [["Klasifikacija", "Broj", "Vrednost (RSD)"]] + [[str(r["Klasifikacija"])[:40], str(r["Broj"]), f"{r['Vrednost (RSD)']:,.2f}" if 'Vrednost (RSD)' in r else "0"] for _, r in klas.iterrows()]
        t = Table(table_data, colWidths=[7*cm, 3*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        fig = px.pie(klas.head(8), values="Broj", names="Klasifikacija", title="Klasifikacija reciklaže")
        fig.update_traces(textposition='inside', textinfo='percent+label')
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # 2. KLASIFIKACIJA ŠTETE
    elements.append(Paragraph("2. Klasifikacija štete", sekcija_style))
    if "Klasifikacija štete" in df.columns:
        steta = df["Klasifikacija štete"].value_counts().head(10).reset_index()
        steta.columns = ["Klasifikacija štete", "Broj"]
        if "Nabavna cena" in df.columns and "Količina" in df.columns:
            vrednost_steta = df.groupby("Klasifikacija štete")["Ukupna vrednost"].sum().reset_index()
            vrednost_steta.columns = ["Klasifikacija štete", "Vrednost (RSD)"]
            steta = steta.merge(vrednost_steta, on="Klasifikacija štete", how="left")
        
        table_data = [["Klasifikacija štete", "Broj", "Vrednost (RSD)"]] + [[str(r["Klasifikacija štete"])[:40], str(r["Broj"]), f"{r['Vrednost (RSD)']:,.2f}" if 'Vrednost (RSD)' in r else "0"] for _, r in steta.iterrows()]
        t = Table(table_data, colWidths=[7*cm, 3*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        fig = px.bar(steta.head(8), x="Klasifikacija štete", y="Broj", title="Klasifikacija štete")
        fig.update_layout(xaxis_tickangle=-45)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # 3. TOP 10 BRENDOVA
    elements.append(Paragraph("3. Top 10 brendova", sekcija_style))
    if "Brend" in df.columns:
        if "Nabavna cena" in df.columns and "Količina" in df.columns:
            brend_analiza = df.groupby("Brend").agg({
                "Broj reklamacije": "count",
                "Ukupna vrednost": "sum"
            }).reset_index()
            brend_analiza.columns = ["Brend", "Broj artikala", "Ukupna vrednost (RSD)"]
        else:
            brend_analiza = df["Brend"].value_counts().head(10).reset_index()
            brend_analiza.columns = ["Brend", "Broj artikala"]
            brend_analiza["Ukupna vrednost (RSD)"] = 0
        
        brend_analiza = brend_analiza.sort_values("Broj artikala", ascending=False).head(10)
        
        table_data = [["Brend", "Broj", "Vrednost (RSD)"]] + [[str(r["Brend"]), str(r["Broj artikala"]), f"{r['Ukupna vrednost (RSD)']:,.2f}"] for _, r in brend_analiza.iterrows()]
        t = Table(table_data, colWidths=[7*cm, 3*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        fig = px.bar(brend_analiza, x="Brend", y="Broj artikala", title="Top 10 brendova")
        fig.update_layout(xaxis_tickangle=-45)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # 4. VREMENSKI TREND
    elements.append(Paragraph("4. Vremenski trend", sekcija_style))
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
        
        trend = df_datum.groupby("Mesec").agg({
            "Broj reklamacije": "count"
        }).reset_index()
        trend.columns = ["Mesec", "Broj artikala"]
        trend["Mesec"] = trend["Mesec"].astype(str)
        
        if "Ukupna vrednost" in df_datum.columns:
            vrednost_trend = df_datum.groupby("Mesec")["Ukupna vrednost"].sum().reset_index()
            vrednost_trend.columns = ["Mesec", "Vrednost (RSD)"]
            vrednost_trend["Mesec"] = vrednost_trend["Mesec"].astype(str)
            trend = trend.merge(vrednost_trend, on="Mesec", how="left")
        
        table_data = [["Mesec", "Broj", "Vrednost (RSD)"]] + [[str(r["Mesec"]), str(r["Broj artikala"]), f"{r['Vrednost (RSD)']:,.2f}" if 'Vrednost (RSD)' in r else "0"] for _, r in trend.iterrows()]
        t = Table(table_data, colWidths=[5*cm, 3*cm, 5*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        fig = px.line(trend, x="Mesec", y="Broj artikala", title="Vremenski trend", markers=True)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=8*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # 5. PARETO ANALIZA (ostaje ista)
    elements.append(Paragraph("5. Pareto analiza (80/20)", sekcija_style))
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        pareto = df.sort_values("Ukupna vrednost", ascending=False)
        pareto["Kumulativni %"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
        
        table_data = [["Broj reklamacije", "Naziv", "Vrednost", "Kum.%"]]
        for _, r in pareto.head(15).iterrows():
            table_data.append([str(r["Broj reklamacije"])[:12], str(r["Naziv artikla"])[:25], f"{r['Ukupna vrednost']:,.0f}", f"{r['Kumulativni %']}%"])
        t = Table(table_data, colWidths=[2.5*cm, 5*cm, 3*cm, 2.5*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 7), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.3*cm))
        granica = pareto[pareto["Kumulativni %"] <= 80].shape[0]
        elements.append(Paragraph(f"📊 {granica} artikala čini 80% ukupne vrednosti", bold_style))
        elements.append(Spacer(1, 0.5*cm))
        
        pareto_top = pareto.head(20).reset_index()
        pareto_top.index = range(1, len(pareto_top) + 1)
        fig = px.bar(pareto_top, x=pareto_top.index, y="Ukupna vrednost", title="Pareto analiza - Top 20 artikala")
        fig.add_hline(y=pareto["Ukupna vrednost"].sum() * 0.8, line_dash="dash", line_color="red", annotation_text="80% granica")
        fig.update_layout(xaxis_title="Redni broj artikla", yaxis_title="Vrednost (RSD)")
        img = plotly_to_reportlab_image(fig, width=14*cm, height=8*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # 6. TOP 10 NAJSKUPLJIH
    elements.append(Paragraph("6. Top 10 najskupljih artikala", sekcija_style))
    if "Nabavna cena" in df.columns:
        naj = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend"]]
        table_data = [["Broj reklamacije", "Naziv", "Cena (RSD)", "Brend"]]
        for _, r in naj.iterrows():
            table_data.append([str(r["Broj reklamacije"])[:12], str(r["Naziv artikla"])[:25], f"{r['Nabavna cena']:,.2f}", str(r["Brend"])])
        t = Table(table_data, colWidths=[2.5*cm, 5*cm, 3*cm, 2.5*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 7), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        fig = px.bar(naj, x="Nabavna cena", y="Naziv artikla", orientation='h', title="Top 10 najskupljih", color="Brend")
        fig.update_layout(xaxis_title="Cena (RSD)", yaxis_title="")
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    
    # Završna
    elements.append(PageBreak())
    elements.append(Spacer(1, 5*cm))
    elements.append(Paragraph("Izveštaj generisan automatski", normal_style))
    elements.append(Paragraph(f"Tehnomanija © {datetime.now().year}", normal_style))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ============================
# EMAIL SLANJE
# ============================

def posalji_email(primalac, excel_data=None, pdf_data=None, dodatne_adrese=None):
    try:
        sender = st.secrets["email"]["sender"]
        password = st.secrets["email"]["password"]
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = primalac
        msg["Subject"] = f"Izveštaj o reciklaži – {datetime.now().strftime('%d.%m.%Y')}"
        
        if dodatne_adrese:
            msg["CC"] = ", ".join(dodatne_adrese)
        
        body = f"Poštovani,\n\nU prilogu vam dostavljamo izveštaj o reciklaži sa stanjem od {datetime.now().strftime('%d.%m.%Y')}.\n\nIzveštaj je generisan automatski.\n\nS poštovanjem,\nTehnomanija"
        msg.attach(MIMEText(body, "plain"))
        
        if excel_data is not None and isinstance(excel_data, bytes):
            excel_part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            excel_part.set_payload(excel_data)
            encoders.encode_base64(excel_part)
            excel_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.xlsx')
            msg.attach(excel_part)
        
        if pdf_data is not None and isinstance(pdf_data, bytes):
            pdf_part = MIMEBase('application', 'pdf')
            pdf_part.set_payload(pdf_data)
            encoders.encode_base64(pdf_part)
            pdf_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.pdf')
            msg.attach(pdf_part)
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender, password)
        
        svi_primaoci = [primalac]
        if dodatne_adrese:
            svi_primaoci.extend(dodatne_adrese)
        
        server.send_message(msg)
        server.quit()
        return True, "Email je uspešno poslat!"
    
    except Exception as e:
        return False, f"Greška pri slanju emaila: {str(e)}"

# ============================
# STREAMLIT UI (sa tabovima)
# ============================

st.title("🔄 Izveštaj o reciklaži")
uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Učitavam..."):
        reciklaza, reklamacije = ucitaj_podatke(uploaded_file)
    if reciklaza is not None and reklamacije is not None:
        with st.spinner("Spajam..."):
            lookup = pripremi_lookup(reklamacije)
            df = spoji(reciklaza, lookup)
        st.success(f"✅ Učitano: {len(reciklaza)} redova, spojeno: {len(df)} redova")
        st.markdown("---")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Pregled", "📈 Napredne metrike", "🏷️ Analize", "📋 Tabela", "📤 Izvoz"])
        
        with tab1:
            st.subheader("Ključni pokazatelji")
            col1, col2, col3, col4 = st.columns(4)
            ukupno = df["Količina"].sum() if "Količina" in df.columns else 0
            vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
            col1.metric("📦 Ukupno artikala", f"{ukupno:,}")
            col2.metric("💰 Ukupna vrednost", f"{vrednost:,.2f} RSD")
            col3.metric("🏷️ Broj brendova", df["Brend"].nunique() if "Brend" in df.columns else 0)
            col4.metric("📂 Robnih grupa", df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0)
            st.markdown("---")
            
            st.subheader("Klasifikacija reciklaže")
            if "Klasifikacija reciklaže" in df.columns:
                klas = df["Klasifikacija reciklaže"].value_counts().reset_index()
                klas.columns = ["Klasifikacija", "Broj"]
                fig = px.pie(klas, values="Broj", names="Klasifikacija", title="Klasifikacija reciklaže")
                st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.subheader("Klasifikacija štete")
            if "Klasifikacija štete" in df.columns:
                steta = df["Klasifikacija štete"].value_counts().reset_index()
                steta.columns = ["Klasifikacija štete", "Broj"]
                fig = px.bar(steta, x="Klasifikacija štete", y="Broj", title="Klasifikacija štete")
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("Napredne metrike")
            if "Nabavna cena" in df.columns and "Količina" in df.columns:
                df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
                pareto = df.sort_values("Ukupna vrednost", ascending=False)
                pareto["Kumulativni %"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
                fig = px.bar(pareto.head(20), x="Naziv artikla", y="Ukupna vrednost", title="Pareto analiza - Top 20 artikala")
                st.plotly_chart(fig, use_container_width=True)
                st.info(f"📊 {pareto[pareto['Kumulativni %'] <= 80].shape[0]} artikala čini 80% ukupne vrednosti")
        
        with tab3:
            st.subheader("Sve analize")
            if "Klasifikacija reciklaže" in df.columns:
                klas = df["Klasifikacija reciklaže"].value_counts().reset_index()
                klas.columns = ["Klasifikacija", "Broj"]
                st.dataframe(klas, use_container_width=True)
        
        with tab4:
            st.dataframe(df, use_container_width=True)
        
        with tab5:
            st.subheader("📤 Izvoz izveštaja")
            
            excel_data = None
            pdf_data = None
            
            st.markdown("### 📊 Excel izveštaj")
            st.info("Excel fajl sadrži **9 sheet-ova** sa svim analizama i vrednostima.")
            
            col1, col2 = st.columns(2)
            with col1:
                try:
                    excel_data = formatiraj_i_sacuvaj(df)
                    st.download_button(
                        label="📥 Preuzmi Excel",
                        data=excel_data,
                        file_name=f"izvestaj_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Greška: {e}")
            
            with col2:
                with st.spinner("Generišem PDF..."):
                    try:
                        pdf_data = generisi_pdf(df)
                        st.download_button(
                            label="📄 Preuzmi PDF (sa grafikonima)",
                            data=pdf_data,
                            file_name=f"izvestaj_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"Greška: {e}")
            
            st.markdown("---")
            
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
                        cc_list = []
                        if email_cc:
                            cc_list = [email.strip() for email in email_cc.split(",") if email.strip()]
                        
                        excel_attach = excel_data if (poslati_excel and excel_data is not None) else None
                        pdf_attach = pdf_data if (posalji_pdf and pdf_data is not None) else None
                        
                        if not excel_attach and not pdf_attach:
                            st.warning("Nijedan prilog nije odabran ili nije generisan.")
                        else:
                            with st.spinner("Šaljem email..."):
                                success, message = posalji_email(email_primalac, excel_attach, pdf_attach, cc_list)
                                if success:
                                    st.success(f"✅ {message}")
                                    if cc_list:
                                        st.info(f"📋 Kopija poslata na: {', '.join(cc_list)}")
                                else:
                                    st.error(f"❌ {message}")
