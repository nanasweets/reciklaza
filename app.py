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
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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
    "datum otvaranja": "Reklamacija: Created Date",
}

KOLONE_OUTPUT = [
    "Broj reklamacije", "ID uređaja", "Naziv artikla", "Brend",
    "Robna grupa", "Količina", "Klasifikacija reciklaže", "Datum obrade",
    "Paleta", "Nabavna cena", "ID ulaza", "Skladišna lokacija",
    "Pozicija u rafu", "Klasifikacija štete", "prevoznik",
    "broj pošiljke", "vrednost fakture",
    "reklamacija otvorena od strane", "model reklamacije",
    "klasifikacija", "način razduženja kupca", "serijski broj",
    "datum otvaranja",
    "Starost (dani)",
    "Rok (30 dana)",
    "Prekoračenje (dani)",
    "Status roka",
]

HEADER_BG = "1F4E79"
HEADER_FONT = "FFFFFF"
LOOKUP_BG = "D9E1F2"
LOOKUP_FONT = "1F4E79"

# Parametri
if "kurs_evra" not in st.session_state:
    st.session_state.kurs_evra = 117.0
if "cena_po_paleti_evri" not in st.session_state:
    st.session_state.cena_po_paleti_evri = 110.0
if "pdv_stopa" not in st.session_state:
    st.session_state.pdv_stopa = 20.0

# ============================
# FUNKCIJE ZA ČIŠĆENJE
# ============================

def ocisti_datum(vrednost):
    if pd.isna(vrednost):
        return pd.NaT
    if isinstance(vrednost, (pd.Timestamp, datetime)):
        return vrednost
    tekst = str(vrednost).strip()
    try:
        return pd.to_datetime(tekst, format='%d.%m.%Y', errors='coerce')
    except:
        pass
    try:
        if ' 0:00:00' in tekst:
            tekst = tekst.replace(' 0:00:00', '')
        return pd.to_datetime(tekst, format='%Y-%m-%d', errors='coerce')
    except:
        pass
    try:
        return pd.to_datetime(tekst, errors='coerce')
    except:
        return pd.NaT

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

# ============================
# OBRADA PODATAKA
# ============================

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
    
    if "Datum obrade" in reciklaza.columns:
        reciklaza["Datum obrade"] = reciklaza["Datum obrade"].apply(ocisti_datum)
        reciklaza = reciklaza[reciklaza["Datum obrade"].notna()]
    
    return reciklaza, reklamacije

def pripremi_lookup(reklamacije):
    kolona_broja = pronadji_kolonu_broja(reklamacije)
    rek = reklamacije.copy()
    rek["_kljuc"] = rek[kolona_broja].apply(cisti_broj)
    rek = rek[rek["_kljuc"] != ""]
    return rek.set_index("_kljuc")

def izracunaj_rokove(df):
    df_copy = df.copy()
    
    if "datum otvaranja" in df_copy.columns and "Datum obrade" in df_copy.columns:
        df_copy["datum otvaranja"] = pd.to_datetime(df_copy["datum otvaranja"], errors="coerce")
        df_copy["Datum obrade"] = pd.to_datetime(df_copy["Datum obrade"], errors="coerce")
        
        danas = datetime.now()
        df_copy = df_copy[
            (df_copy["Datum obrade"].notna()) & 
            (df_copy["Datum obrade"] <= danas)
        ]
        
        if len(df_copy) > 0:
            df_copy["Starost (dani)"] = (df_copy["Datum obrade"] - df_copy["datum otvaranja"]).dt.days
            df_copy["Rok (30 dana)"] = df_copy["datum otvaranja"] + pd.Timedelta(days=30)
            df_copy["Prekoračenje (dani)"] = df_copy["Starost (dani)"] - 30
            df_copy["Status roka"] = df_copy["Prekoračenje (dani)"].apply(
                lambda x: "✅ U roku" if (pd.isna(x) or x <= 0) else "❌ Prekoračen"
            )
        else:
            df_copy["Starost (dani)"] = None
            df_copy["Rok (30 dana)"] = None
            df_copy["Prekoračenje (dani)"] = None
            df_copy["Status roka"] = "N/A"
    else:
        df_copy["Starost (dani)"] = None
        df_copy["Rok (30 dana)"] = None
        df_copy["Prekoračenje (dani)"] = None
        df_copy["Status roka"] = "Nedostaju podaci"
    
    return df_copy

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
    df = izracunaj_rokove(df)
    extras = [c for c in df.columns if c not in KOLONE_OUTPUT]
    finalne = [c for c in KOLONE_OUTPUT if c in df.columns] + extras
    return df[finalne]

# ============================
# EXCEL FUNKCIJE
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
    ws.freeze_panes = "A2"

def upisi_kpi(ws, df):
    ukupno_artikala = df["Količina"].sum() if "Količina" in df.columns else 0
    ukupna_vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    broj_brendova = df["Brend"].nunique() if "Brend" in df.columns else 0
    broj_grupa = df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0
    
    u_roku = df[df["Status roka"] == "✅ U roku"].shape[0] if "Status roka" in df.columns else 0
    prekoraceno = df[df["Status roka"] == "❌ Prekoračen"].shape[0] if "Status roka" in df.columns else 0
    
    podaci = [
        ["Ključni pokazatelj", "Vrednost"],
        ["Ukupan broj artikala", ukupno_artikala],
        ["Ukupna vrednost (RSD)", f"{ukupna_vrednost:,.2f}"],
        ["Broj brendova", broj_brendova],
        ["Broj robnih grupa", broj_grupa],
        ["Broj redova", len(df)],
        ["✅ U roku (≤30 dana)", u_roku],
        ["❌ Prekoračen (>30 dana)", prekoraceno],
    ]
    
    for row_idx, row in enumerate(podaci, start=1):
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value)
    
    for col_idx in [1, 2]:
        ws.column_dimensions[get_column_letter(col_idx)].width = 30
    
    for row_idx in range(1, len(podaci) + 1):
        ws.cell(row=row_idx, column=1).font = Font(bold=True, name="Arial", size=11)
        ws.cell(row=row_idx, column=2).font = Font(name="Arial", size=11)
    
    ws.cell(row=1, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    ws.cell(row=1, column=2).font = Font(bold=True, size=12, color=HEADER_BG)

def upisi_grupisano(ws, df, kolona, naslov):
    if kolona not in df.columns:
        ws.cell(row=1, column=1, value=f"Kolona '{kolona}' ne postoji")
        return
    
    df_copy = df.copy()
    if "Datum obrade" in df_copy.columns:
        df_copy["Datum obrade"] = pd.to_datetime(df_copy["Datum obrade"], errors="coerce")
        danas = datetime.now()
        df_copy = df_copy[(df_copy["Datum obrade"].notna()) & (df_copy["Datum obrade"] <= danas)]
    
    grupisan = df_copy[kolona].value_counts().reset_index()
    grupisan.columns = [kolona, "Broj artikala"]
    
    if "Nabavna cena" in df_copy.columns and "Količina" in df_copy.columns:
        df_copy["Ukupna vrednost"] = df_copy["Nabavna cena"] * df_copy["Količina"]
        vrednost_grupisano = df_copy.groupby(kolona)["Ukupna vrednost"].sum().reset_index()
        vrednost_grupisano.columns = [kolona, "Ukupna vrednost (RSD)"]
        grupisan = grupisan.merge(vrednost_grupisano, on=kolona, how="left")
    else:
        grupisan["Ukupna vrednost (RSD)"] = 0
    
    ws.cell(row=1, column=1, value=kolona)
    ws.cell(row=1, column=2, value="Broj artikala")
    ws.cell(row=1, column=3, value="Ukupna vrednost (RSD)")
    primeni_stil_header(ws, 3)
    
    for row_idx, row in enumerate(grupisan.itertuples(index=False), start=2):
        ws.cell(row=row_idx, column=1).value = excel_compatible(row[0])
        ws.cell(row=row_idx, column=2).value = excel_compatible(row[1])
        if len(row) > 2 and row[2] is not None:
            ws.cell(row=row_idx, column=3).value = excel_compatible(row[2])
    primeni_stil_podaci(ws, 3, len(grupisan) + 1)
    
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 25

def upisi_top_brendove(ws, df):
    if "Brend" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Brend' ne postoji")
        return
    
    df_copy = df.copy()
    if "Datum obrade" in df_copy.columns:
        df_copy["Datum obrade"] = pd.to_datetime(df_copy["Datum obrade"], errors="coerce")
        danas = datetime.now()
        df_copy = df_copy[(df_copy["Datum obrade"].notna()) & (df_copy["Datum obrade"] <= danas)]
    
    if "Nabavna cena" in df_copy.columns and "Količina" in df_copy.columns:
        df_copy["Ukupna vrednost"] = df_copy["Nabavna cena"] * df_copy["Količina"]
        brend_analiza = df_copy.groupby("Brend").agg({
            "Broj reklamacije": "count",
            "Ukupna vrednost": "sum"
        }).reset_index()
        brend_analiza.columns = ["Brend", "Broj artikala", "Ukupna vrednost (RSD)"]
    else:
        brend_analiza = df_copy["Brend"].value_counts().head(10).reset_index()
        brend_analiza.columns = ["Brend", "Broj artikala"]
        brend_analiza["Ukupna vrednost (RSD)"] = 0
    
    brend_analiza = brend_analiza.sort_values("Broj artikala", ascending=False).head(10)
    
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

def upisi_vremenski_trend(ws, df):
    if "Datum obrade" not in df.columns:
        ws.cell(row=1, column=1, value="Kolona 'Datum obrade' ne postoji")
        return
    
    df_datum = df.copy()
    df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
    danas = datetime.now()
    df_datum = df_datum[
        (df_datum["Datum obrade"].notna()) & 
        (df_datum["Datum obrade"] <= danas) &
        (df_datum["Datum obrade"] >= datetime(2020, 1, 1))
    ]
    
    if len(df_datum) == 0:
        ws.cell(row=1, column=1, value="Nema validnih datuma za prikaz")
        return
    
    df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
    
    trend = df_datum.groupby("Mesec").agg({
        "Broj reklamacije": "count"
    }).reset_index()
    trend.columns = ["Mesec", "Broj artikala"]
    trend["Mesec"] = trend["Mesec"].astype(str)
    
    if "Nabavna cena" in df_datum.columns and "Količina" in df_datum.columns:
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
        ws.column_dimensions[get_column_letter(col_idx)].width = 25

# ============================
# SPECIFIKACIJA ISPRAVKE FINANSIJSKE VREDNOSTI (ISPRAVLJENA)
# ============================

def upisi_specifikaciju_ispravke(ws, df, kurs_evra=117, cena_po_paleti_evri=110, pdv_stopa=20):
    """Sheet: Specifikacija ispravke finansijske vrednosti"""
    
    ws.cell(row=1, column=1, value="SPECIFIKACIJA ISPRAVKE FINANSIJSKE VREDNOSTI")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color=HEADER_BG)
    
    # Datum
    ws.cell(row=2, column=1, value=f"Datum izveštaja: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    ws.cell(row=2, column=1).font = Font(name="Arial", size=10)
    
    row = 4
    
    if "Skladišna lokacija" not in df.columns or "Nabavna cena" not in df.columns or "Količina" not in df.columns:
        ws.cell(row=row, column=1, value="Nedostaju potrebne kolone")
        return
    
    df_copy = df.copy()
    
    # Ukupni agregati za % obezvređenja
    ukupna_kolicina = df_copy["Količina"].sum()
    ukupna_nabavna_bez_pdv = (df_copy["Nabavna cena"] * df_copy["Količina"]).sum()
    ukupno_paleta = df_copy["Skladišna lokacija"].nunique()
    ukupna_ponuda_bruto = ukupno_paleta * cena_po_paleti_evri * kurs_evra
    ukupna_ponuda_neto = ukupna_ponuda_bruto / (1 + pdv_stopa / 100)
    
    # % obezvređenja
    razlika = ukupna_nabavna_bez_pdv - ukupna_ponuda_neto
    procenat_obezvredenja = (razlika / ukupna_nabavna_bez_pdv) * 100 if ukupna_nabavna_bez_pdv > 0 else 0
    
    # Sumarna tabela (zaglavlje)
    row += 1
    ws.cell(row=row, column=1, value="SUMNARNI PREGLED")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    row += 1
    
    headers_sum = [
        "Magacin",
        "Zbir od Količina",
        "Zbir od Nabavna cena bez PDV-a",
        "MP cena bez PDV-a",
        "Razlika između nabavne i MP cene bez PDV-a",
        "% obezvređenja"
    ]
    
    for col_idx, h in enumerate(headers_sum, start=1):
        ws.cell(row=row, column=col_idx, value=h)
    primeni_stil_header(ws, len(headers_sum))
    row += 1
    
    magacin = "DC 74-Magacin reciklaže"
    ws.cell(row=row, column=1).value = magacin
    ws.cell(row=row, column=2).value = ukupna_kolicina
    ws.cell(row=row, column=3).value = round(ukupna_nabavna_bez_pdv, 2)
    ws.cell(row=row, column=4).value = round(ukupna_ponuda_neto, 2)
    ws.cell(row=row, column=5).value = round(razlika, 2)
    ws.cell(row=row, column=6).value = f"{round(procenat_obezvredenja, 2)}%"
    row += 2
    
    # Detaljna specifikacija (svaki artikal)
    ws.cell(row=row, column=1, value="DETALJNA SPECIFIKACIJA")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color=HEADER_BG)
    row += 1
    
    # Kolone kako je traženo
    headers = [
        "Šifra", "Artikli", "Jed. Mera", "Lager",
        "Ponuda po komadu bez PDV-a",
        "Ponuda po kom sa PDV",
        "Ukupna vrednost ponude sa PDV-om",
        "Nabavna cena bez pdv-a",
        "Nabavna cena sa pdv-om"
    ]
    
    sifra_kolona = "ID uređaja" if "ID uređaja" in df_copy.columns else "Broj reklamacije"
    
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=row, column=col_idx, value=h)
    primeni_stil_header(ws, len(headers))
    row += 1
    
    # Podaci po artiklima
    for _, r in df_copy.iterrows():
        # Izračun za ovaj artikal
        nabavna_bez_pdv = r["Nabavna cena"]
        nabavna_sa_pdv = nabavna_bez_pdv * (1 + pdv_stopa / 100)
        ponuda_po_komadu_neto = nabavna_bez_pdv * (1 - procenat_obezvredenja / 100)
        ponuda_po_komadu_bruto = ponuda_po_komadu_neto * (1 + pdv_stopa / 100)
        ukupna_vrednost_ponude_bruto = ponuda_po_komadu_bruto * r["Količina"]
        
        ws.cell(row=row, column=1).value = excel_compatible(r.get(sifra_kolona, "N/A"))
        ws.cell(row=row, column=2).value = excel_compatible(r["Naziv artikla"])
        ws.cell(row=row, column=3).value = "kom"
        ws.cell(row=row, column=4).value = excel_compatible(r["Količina"])
        ws.cell(row=row, column=5).value = excel_compatible(round(ponuda_po_komadu_neto, 2))
        ws.cell(row=row, column=6).value = excel_compatible(round(ponuda_po_komadu_bruto, 2))
        ws.cell(row=row, column=7).value = excel_compatible(round(ukupna_vrednost_ponude_bruto, 2))
        ws.cell(row=row, column=8).value = excel_compatible(nabavna_bez_pdv)
        ws.cell(row=row, column=9).value = excel_compatible(round(nabavna_sa_pdv, 2))
        row += 1
    
    for col in range(1, 10):
        ws.column_dimensions[get_column_letter(col)].width = 25

# ============================
# PREDLOG PRODAJE (sa PDV obračunom)
# ============================

def upisi_predlog_prodaje(ws, df, kurs_evra=117, cena_po_paleti_evri=110, pdv_stopa=20):
    ws.cell(row=1, column=1, value="PREDLOG ZA PRODAJU TREĆEM LICU")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color=HEADER_BG)
    
    row = 3
    
    if "Skladišna lokacija" not in df.columns:
        ws.cell(row=row, column=1, value="Nema podataka o skladišnim lokacijama")
        return
    
    df_copy = df.copy()
    df_copy["Ukupna nabavna vrednost (RSD)"] = df_copy["Nabavna cena"] * df_copy["Količina"]
    df_copy["Nabavna cena sa PDV"] = df_copy["Nabavna cena"] * (1 + pdv_stopa / 100)
    df_copy["Ukupna nabavna sa PDV"] = df_copy["Nabavna cena sa PDV"] * df_copy["Količina"]
    ukupno_paleta = df_copy["Skladišna lokacija"].nunique()
    
    ws.cell(row=row, column=1, value="PARAMETRI")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Kurs evra (RSD)")
    ws.cell(row=row, column=2, value=kurs_evra)
    row += 1
    ws.cell(row=row, column=1, value="Cena po paleti (€) sa PDV-om")
    ws.cell(row=row, column=2, value=cena_po_paleti_evri)
    row += 1
    ws.cell(row=row, column=1, value="PDV stopa (%)")
    ws.cell(row=row, column=2, value=pdv_stopa)
    row += 2
    
    # Izračun
    ukupna_nabavna_bez_pdv = df_copy["Ukupna nabavna vrednost (RSD)"].sum()
    ukupna_nabavna_sa_pdv = df_copy["Ukupna nabavna sa PDV"].sum()
    ukupna_ponuda_bruto = ukupno_paleta * cena_po_paleti_evri * kurs_evra
    ukupna_ponuda_neto = ukupna_ponuda_bruto / (1 + pdv_stopa / 100)
    gubitak_neto = ukupna_nabavna_bez_pdv - ukupna_ponuda_neto
    gubitak_bruto = ukupna_nabavna_sa_pdv - ukupna_ponuda_bruto
    
    ws.cell(row=row, column=1, value=f"📦 Broj paleta = broj unikatnih lokacija: {ukupno_paleta}")
    row += 1
    ws.cell(row=row, column=1, value=f"💶 Ukupna ponuda (bruto, sa PDV-om): {ukupna_ponuda_bruto:,.2f} RSD")
    row += 1
    ws.cell(row=row, column=1, value=f"💶 Ukupna ponuda (neto, bez PDV-a): {ukupna_ponuda_neto:,.2f} RSD")
    row += 2
    
    # Tabela
    headers = [
        "Broj reklamacije", "Naziv artikla", "Količina",
        "Nabavna cena bez PDV", "Nabavna cena sa PDV",
        "Ukupna nabavna sa PDV", "Skladišna lokacija", "Status roka"
    ]
    
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=row, column=col_idx, value=h)
    primeni_stil_header(ws, len(headers))
    row += 1
    
    for _, r in df_copy.iterrows():
        ws.cell(row=row, column=1).value = excel_compatible(r["Broj reklamacije"])
        ws.cell(row=row, column=2).value = excel_compatible(r["Naziv artikla"])
        ws.cell(row=row, column=3).value = excel_compatible(r["Količina"])
        ws.cell(row=row, column=4).value = excel_compatible(r["Nabavna cena"])
        ws.cell(row=row, column=5).value = excel_compatible(r["Nabavna cena sa PDV"])
        ws.cell(row=row, column=6).value = excel_compatible(r["Ukupna nabavna sa PDV"])
        ws.cell(row=row, column=7).value = excel_compatible(r["Skladišna lokacija"])
        ws.cell(row=row, column=8).value = excel_compatible(r["Status roka"]) if "Status roka" in r else "N/A"
        row += 1
    
    row += 2
    ws.cell(row=row, column=1, value="SAŽETAK")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Ukupno artikala")
    ws.cell(row=row, column=2, value=df_copy["Količina"].sum())
    row += 1
    ws.cell(row=row, column=1, value="Ukupno paleta")
    ws.cell(row=row, column=2, value=ukupno_paleta)
    row += 1
    ws.cell(row=row, column=1, value="Ukupna nabavna vrednost (bez PDV)")
    ws.cell(row=row, column=2, value=f"{ukupna_nabavna_bez_pdv:,.2f}")
    row += 1
    ws.cell(row=row, column=1, value="Ukupna nabavna vrednost (sa PDV)")
    ws.cell(row=row, column=2, value=f"{ukupna_nabavna_sa_pdv:,.2f}")
    row += 1
    ws.cell(row=row, column=1, value="Ukupna ponuda (bruto RSD)")
    ws.cell(row=row, column=2, value=f"{ukupna_ponuda_bruto:,.2f}")
    row += 1
    ws.cell(row=row, column=1, value="Ukupna ponuda (neto RSD)")
    ws.cell(row=row, column=2, value=f"{ukupna_ponuda_neto:,.2f}")
    row += 1
    ws.cell(row=row, column=1, value="GUBITAK (bruto, sa PDV)")
    ws.cell(row=row, column=1).font = Font(bold=True, color="FF0000")
    ws.cell(row=row, column=2, value=f"{gubitak_bruto:,.2f} RSD")
    ws.cell(row=row, column=2).font = Font(bold=True, color="FF0000")
    row += 1
    ws.cell(row=row, column=1, value="GUBITAK (neto, bez PDV)")
    ws.cell(row=row, column=1).font = Font(bold=True, color="FF0000")
    ws.cell(row=row, column=2, value=f"{gubitak_neto:,.2f} RSD")
    ws.cell(row=row, column=2).font = Font(bold=True, color="FF0000")
    
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 25

def upisi_analiza_rokova(ws, df):
    ws.cell(row=1, column=1, value="Analiza rokova reklamacija")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14, color=HEADER_BG)
    
    row = 3
    
    if "Status roka" not in df.columns or "Starost (dani)" not in df.columns:
        ws.cell(row=row, column=1, value="Nedostaju podaci za analizu rokova")
        return
    
    df_copy = df.copy()
    if "Datum obrade" in df_copy.columns:
        df_copy["Datum obrade"] = pd.to_datetime(df_copy["Datum obrade"], errors="coerce")
        danas = datetime.now()
        df_copy = df_copy[(df_copy["Datum obrade"].notna()) & (df_copy["Datum obrade"] <= danas)]
    
    if len(df_copy) == 0:
        ws.cell(row=row, column=1, value="Nema validnih podataka za prikaz")
        return
    
    u_roku = df_copy[df_copy["Status roka"] == "✅ U roku"].shape[0]
    prekoraceno = df_copy[df_copy["Status roka"] == "❌ Prekoračen"].shape[0]
    ukupno = len(df_copy)
    
    ws.cell(row=row, column=1, value="Statistika rokova")
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1
    
    ws.cell(row=row, column=1, value="Ukupno reklamacija")
    ws.cell(row=row, column=2, value=ukupno)
    row += 1
    
    ws.cell(row=row, column=1, value="✅ U roku (≤30 dana)")
    ws.cell(row=row, column=2, value=u_roku)
    ws.cell(row=row, column=3, value=f"{u_roku/ukupno*100:.1f}%" if ukupno > 0 else "0%")
    row += 1
    
    ws.cell(row=row, column=1, value="❌ Prekoračen (>30 dana)")
    ws.cell(row=row, column=2, value=prekoraceno)
    ws.cell(row=row, column=3, value=f"{prekoraceno/ukupno*100:.1f}%" if ukupno > 0 else "0%")
    row += 2
    
    ws.cell(row=row, column=1, value="Prosečno prekoračenje (dani)")
    ws.cell(row=row, column=2, value=excel_compatible(df_copy["Prekoračenje (dani)"].mean().round(0)))
    row += 1
    
    ws.cell(row=row, column=1, value="Maksimalno prekoračenje (dani)")
    ws.cell(row=row, column=2, value=excel_compatible(df_copy["Prekoračenje (dani)"].max()))
    row += 1
    
    ws.cell(row=row, column=1, value="Top 10 najvećih prekoračenja")
    ws.cell(row=row, column=1).font = Font(bold=True, size=11)
    row += 1
    
    headers = ["Broj reklamacije", "Naziv artikla", "Prekoračenje (dani)", "Otvorio"]
    for col_idx, h in enumerate(headers, start=1):
        ws.cell(row=row, column=col_idx, value=h)
    primeni_stil_header(ws, len(headers))
    row += 1
    
    najgori = df_copy.nlargest(10, "Prekoračenje (dani)")[["Broj reklamacije", "Naziv artikla", "Prekoračenje (dani)", "reklamacija otvorena od strane"]]
    for _, r in najgori.iterrows():
        ws.cell(row=row, column=1).value = excel_compatible(r["Broj reklamacije"])
        ws.cell(row=row, column=2).value = excel_compatible(r["Naziv artikla"])
        ws.cell(row=row, column=3).value = excel_compatible(r["Prekoračenje (dani)"])
        ws.cell(row=row, column=4).value = excel_compatible(r["reklamacija otvorena od strane"])
        row += 1
    
    for col in range(1, 5):
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
    
    ws = wb.create_sheet("Vremenski trend")
    upisi_vremenski_trend(ws, df)
    
    ws = wb.create_sheet("Analiza rokova")
    upisi_analiza_rokova(ws, df)
    
    ws = wb.create_sheet("Predlog prodaje")
    upisi_predlog_prodaje(ws, df, st.session_state.kurs_evra, st.session_state.cena_po_paleti_evri, st.session_state.pdv_stopa)
    
    ws = wb.create_sheet("Specifikacija ispravke")
    upisi_specifikaciju_ispravke(ws, df, st.session_state.kurs_evra, st.session_state.cena_po_paleti_evri, st.session_state.pdv_stopa)
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

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
    
    df_copy = df.copy()
    if "Datum obrade" in df_copy.columns:
        df_copy["Datum obrade"] = pd.to_datetime(df_copy["Datum obrade"], errors="coerce")
        danas = datetime.now()
        df_copy = df_copy[(df_copy["Datum obrade"].notna()) & (df_copy["Datum obrade"] <= danas)]
    
    ukupno = df_copy["Količina"].sum() if "Količina" in df_copy.columns else 0
    vrednost = (df_copy["Nabavna cena"] * df_copy["Količina"]).sum() if "Nabavna cena" in df_copy.columns and "Količina" in df_copy.columns else 0
    broj_brendova = df_copy["Brend"].nunique() if "Brend" in df_copy.columns else 0
    
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
    
    # SPECIFIKACIJA ISPRAVKE
    elements.append(Paragraph("Specifikacija ispravke finansijske vrednosti", sekcija_style))
    elements.append(Paragraph(f"Datum: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
    elements.append(Spacer(1, 0.5*cm))
    
    if "Skladišna lokacija" in df_copy.columns and "Nabavna cena" in df_copy.columns:
        ukupno_paleta = df_copy["Skladišna lokacija"].nunique()
        ukupna_kolicina = df_copy["Količina"].sum()
        ukupna_nabavna_bez_pdv = (df_copy["Nabavna cena"] * df_copy["Količina"]).sum()
        ukupna_ponuda_bruto = ukupno_paleta * st.session_state.cena_po_paleti_evri * st.session_state.kurs_evra
        ukupna_ponuda_neto = ukupna_ponuda_bruto / (1 + st.session_state.pdv_stopa / 100)
        razlika = ukupna_nabavna_bez_pdv - ukupna_ponuda_neto
        procenat = (razlika / ukupna_nabavna_bez_pdv) * 100 if ukupna_nabavna_bez_pdv > 0 else 0
        
        sum_data = [
            ["Magacin", "Zbir Količina", "Zbir Nabavna cena bez PDV", "MP cena bez PDV", "Razlika", "% obezvređenja"],
            ["DC 74-Magacin reciklaže", 
             f"{ukupna_kolicina:.0f}", 
             f"{ukupna_nabavna_bez_pdv:,.2f}", 
             f"{ukupna_ponuda_neto:,.2f}",
             f"{razlika:,.2f}",
             f"{procenat:.2f}%"]
        ]
        sum_table = Table(sum_data, colWidths=[4*cm, 2.5*cm, 4*cm, 4*cm, 4*cm, 3*cm])
        sum_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (1,0), (-1,-1), 'RIGHT')]))
        elements.append(sum_table)
        
        # Dodaj kratak opis cene po komadu (prosečna)
        avg_cena_neto = ukupna_ponuda_neto / ukupna_kolicina if ukupna_kolicina > 0 else 0
        elements.append(Spacer(1, 0.5*cm))
        elements.append(Paragraph(f"Prosečna cena po komadu (neto): {avg_cena_neto:,.2f} RSD", bold_style))
    
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
        
        body = f"Poštovani,\n\nU prilogu vam dostavljamo izveštaj o reciklaži sa stanjem od {datetime.now().strftime('%d.%m.%Y')}.\n\nIzveštaj sadrži Specifikaciju ispravke finansijske vrednosti.\n\nIzveštaj je generisan automatski.\n\nS poštovanjem,\nTehnomanija"
        msg.attach(MIMEText(body, "plain"))
        
        if excel_data is not None and isinstance(excel_data, bytes):
            excel_part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            excel_part.set_payload(excel_data)
            encoders.encode_base64(excel_part)
            excel_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_{datetime.now().strftime("%Y%m%d")}.xlsx')
            msg.attach(excel_part)
        
        if pdf_data is not None and isinstance(pdf_data, bytes):
            pdf_part = MIMEBase('application', 'pdf')
            pdf_part.set_payload(pdf_data)
            encoders.encode_base64(pdf_part)
            pdf_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_{datetime.now().strftime("%Y%m%d")}.pdf')
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
        return False, f"Greška: {str(e)}"

# ============================
# STREAMLIT UI
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
        
        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "📊 Pregled",
            "📈 Grafikoni",
            "🏷️ Analize",
            "📋 Tabela",
            "📦 Prodaja 3. licu",
            "⏰ Rokovi",
            "💰 Specifikacija",
            "📤 Izvoz"
        ])
        
        with tab1:
            st.subheader("Ključni pokazatelji")
            col1, col2, col3, col4, col5 = st.columns(5)
            ukupno = df["Količina"].sum() if "Količina" in df.columns else 0
            vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
            col1.metric("📦 Artikala", f"{ukupno:,}")
            col2.metric("💰 Vrednost", f"{vrednost:,.2f} RSD")
            col3.metric("🏷️ Brendova", df["Brend"].nunique() if "Brend" in df.columns else 0)
            col4.metric("📦 Paleta", df["Skladišna lokacija"].nunique() if "Skladišna lokacija" in df.columns else 0)
            
            if "Status roka" in df.columns:
                u_roku = df[df["Status roka"] == "✅ U roku"].shape[0]
                prekoraceno = df[df["Status roka"] == "❌ Prekoračen"].shape[0]
                col5.metric("📊 U roku", f"{u_roku/(u_roku+prekoraceno)*100:.1f}%" if (u_roku+prekoraceno) > 0 else "0%")
        
        with tab2:
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
            
            st.markdown("---")
            st.subheader("Top 10 brendova")
            if "Brend" in df.columns:
                brend = df["Brend"].value_counts().head(10).reset_index()
                brend.columns = ["Brend", "Broj"]
                fig = px.bar(brend, x="Brend", y="Broj", title="Top 10 brendova")
                st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Klasifikacija reciklaže")
                if "Klasifikacija reciklaže" in df.columns:
                    klas = df["Klasifikacija reciklaže"].value_counts().reset_index()
                    klas.columns = ["Klasifikacija", "Broj"]
                    st.dataframe(klas, use_container_width=True)
            
            with col2:
                st.subheader("Klasifikacija štete")
                if "Klasifikacija štete" in df.columns:
                    steta = df["Klasifikacija štete"].value_counts().reset_index()
                    steta.columns = ["Klasifikacija štete", "Broj"]
                    st.dataframe(steta, use_container_width=True)
        
        with tab4:
            st.dataframe(df, use_container_width=True)
        
        with tab5:
            st.subheader("📦 Predlog za prodaju trećem licu")
            st.info("""
            **📌 SVI artikli se prodaju trećem licu (bez obzira na rok)**
            - **Broj paleta** = broj unikatnih **skladišnih lokacija**
            - **Nabavna cena** je **bez PDV-a** (20%)
            - **Ponuda** je **sa PDV-om** (bruto)
            """)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                kurs_evra = st.number_input("💶 Kurs evra (RSD)", min_value=50.0, max_value=200.0, value=st.session_state.kurs_evra, step=0.5)
                st.session_state.kurs_evra = kurs_evra
            with col2:
                cena_po_paleti = st.number_input("💰 Cena po paleti (€)", min_value=10.0, max_value=500.0, value=st.session_state.cena_po_paleti_evri, step=5.0)
                st.session_state.cena_po_paleti_evri = cena_po_paleti
            with col3:
                pdv_stopa = st.number_input("🧾 PDV stopa (%)", min_value=0.0, max_value=50.0, value=st.session_state.pdv_stopa, step=0.5)
                st.session_state.pdv_stopa = pdv_stopa
            
            if "Skladišna lokacija" in df.columns and "Nabavna cena" in df.columns:
                df_copy = df.copy()
                df_copy["Ukupna nabavna"] = df_copy["Nabavna cena"] * df_copy["Količina"]
                df_copy["Nabavna sa PDV"] = df_copy["Nabavna cena"] * (1 + pdv_stopa / 100)
                df_copy["Ukupna nabavna sa PDV"] = df_copy["Nabavna sa PDV"] * df_copy["Količina"]
                
                ukupno_paleta = df_copy["Skladišna lokacija"].nunique()
                ukupno_artikala = df_copy["Količina"].sum()
                ukupna_nabavna = df_copy["Ukupna nabavna sa PDV"].sum()
                ukupna_ponuda_bruto = ukupno_paleta * cena_po_paleti * kurs_evra
                gubitak = ukupna_nabavna - ukupna_ponuda_bruto
                
                st.markdown("### 📊 Sažetak")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("📦 Artikala", f"{ukupno_artikala:,}")
                col2.metric("📦 Paleta", ukupno_paleta)
                col3.metric("💰 Ponuda (bruto)", f"{ukupna_ponuda_bruto:,.2f} RSD")
                col4.metric("📉 Gubitak (bruto)", f"{gubitak:,.2f} RSD")
                
                st.markdown("### 📋 Detalji")
                prikaz = ["Broj reklamacije", "Naziv artikla", "Količina", "Nabavna cena", "Nabavna sa PDV", "Ukupna nabavna sa PDV", "Skladišna lokacija", "Status roka"]
                dostupne = [c for c in prikaz if c in df_copy.columns]
                st.dataframe(df_copy[dostupne], use_container_width=True)
        
        with tab6:
            st.subheader("⏰ Analiza rokova")
            if "Status roka" in df.columns:
                u_roku = df[df["Status roka"] == "✅ U roku"].shape[0]
                prekoraceno = df[df["Status roka"] == "❌ Prekoračen"].shape[0]
                
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ U roku", u_roku)
                col2.metric("❌ Prekoračen", prekoraceno)
                if u_roku + prekoraceno > 0:
                    col3.metric("📊 Procenat", f"{u_roku/(u_roku+prekoraceno)*100:.1f}%")
                
                st.markdown("---")
                st.subheader("Top 10 prekoračenja")
                df_rok = df.copy()
                if "Datum obrade" in df_rok.columns:
                    df_rok["Datum obrade"] = pd.to_datetime(df_rok["Datum obrade"], errors="coerce")
                    danas = datetime.now()
                    df_rok = df_rok[(df_rok["Datum obrade"].notna()) & (df_rok["Datum obrade"] <= danas)]
                
                if len(df_rok) > 0 and "Prekoračenje (dani)" in df_rok.columns:
                    st.dataframe(
                        df_rok.nlargest(10, "Prekoračenje (dani)")[
                            ["Broj reklamacije", "Naziv artikla", "Prekoračenje (dani)", "reklamacija otvorena od strane"]
                        ],
                        use_container_width=True
                    )
            else:
                st.warning("Nedostaju podaci za analizu rokova")
        
                with tab7:
            st.subheader("💰 Specifikacija ispravke finansijske vrednosti")
            st.info("""
            **Dokument prikazuje finansijski efekat prodaje trećem licu.**
            - **% obezvređenja** = (Nabavna vrednost - MP cena) / Nabavna vrednost × 100
            - **Ponuda po komadu bez PDV** = Nabavna cena × (1 - % obezvređenja / 100)
            - **Ponuda sa PDV** = Ponuda bez PDV × (1 + PDV stopa / 100)
            - **Ukupna vrednost ponude** = Ponuda sa PDV × Količina
            """)
            
            st.markdown(f"**Datum izveštaja:** {datetime.now().strftime('%d.%m.%Y %H:%M')}")
            
            if "Skladišna lokacija" in df.columns and "Nabavna cena" in df.columns:
                df_copy = df.copy()
                df_copy["Nabavna cena sa PDV"] = df_copy["Nabavna cena"] * (1 + st.session_state.pdv_stopa / 100)
                df_copy["Ukupna nabavna sa PDV"] = df_copy["Nabavna cena sa PDV"] * df_copy["Količina"]
                
                ukupno_paleta = df_copy["Skladišna lokacija"].nunique()
                ukupna_kolicina = df_copy["Količina"].sum()
                ukupna_nabavna_bez_pdv = (df_copy["Nabavna cena"] * df_copy["Količina"]).sum()
                ukupna_ponuda_bruto = ukupno_paleta * st.session_state.cena_po_paleti_evri * st.session_state.kurs_evra
                ukupna_ponuda_neto = ukupna_ponuda_bruto / (1 + st.session_state.pdv_stopa / 100)
                
                razlika = ukupna_nabavna_bez_pdv - ukupna_ponuda_neto
                procenat = (razlika / ukupna_nabavna_bez_pdv) * 100 if ukupna_nabavna_bez_pdv > 0 else 0
                
                st.markdown("### 📊 Sumarni pregled")
                sum_data = {
                    "Magacin": "DC 74-Magacin reciklaže",
                    "Zbir Količina": f"{ukupna_kolicina:.0f}",
                    "Zbir Nabavna cena bez PDV": f"{ukupna_nabavna_bez_pdv:,.2f}",
                    "MP cena bez PDV": f"{ukupna_ponuda_neto:,.2f}",
                    "Razlika": f"{razlika:,.2f}",
                    "% obezvređenja": f"{procenat:.2f}%"
                }
                st.dataframe(pd.DataFrame([sum_data]), use_container_width=True)
                
                st.markdown("### 📋 Detaljna specifikacija")
                
                # Izračunaj sve kolone
                df_display = df_copy.copy()
                df_display["Ponuda po komadu bez PDV"] = df_display["Nabavna cena"] * (1 - procenat / 100)
                df_display["Ponuda po kom sa PDV"] = df_display["Ponuda po komadu bez PDV"] * (1 + st.session_state.pdv_stopa / 100)
                df_display["Ukupna vrednost ponude sa PDV"] = df_display["Ponuda po kom sa PDV"] * df_display["Količina"]
                df_display["Nabavna cena sa pdv"] = df_display["Nabavna cena"] * (1 + st.session_state.pdv_stopa / 100)
                
                # Odredi šifru – ako postoji ID uređaja, koristi ga, inače Broj reklamacije
                sifra_kol = "ID uređaja" if "ID uređaja" in df_display.columns else "Broj reklamacije"
                
                # Preimenuj kolone za prikaz
                df_display = df_display.rename(columns={
                    sifra_kol: "Šifra",
                    "Naziv artikla": "Artikli",
                    "Količina": "Lager",
                    "Nabavna cena": "Nabavna cena bez pdv-a"
                })
                
                # Definiši redosled kolona za prikaz (sa novim nazivima)
                dostupne = ["Šifra", "Artikli", "Lager", 
                            "Ponuda po komadu bez PDV", "Ponuda po kom sa PDV", 
                            "Ukupna vrednost ponude sa PDV", "Nabavna cena bez pdv-a", "Nabavna cena sa pdv"]
                # Filtriraj samo one koje stvarno postoje
                dostupne = [c for c in dostupne if c in df_display.columns]
                
                st.dataframe(df_display[dostupne], use_container_width=True)
            else:
                st.warning("Nedostaju podaci za specifikaciju.")
        
        with tab8:
            st.subheader("📤 Izvoz")
            
            excel_data = formatiraj_i_sacuvaj(df)
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    label="📥 Excel (10 sheet-ova)",
                    data=excel_data,
                    file_name=f"izvestaj_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                with st.spinner("Generišem PDF..."):
                    try:
                        pdf_data = generisi_pdf(df)
                        st.download_button(
                            label="📄 PDF (sa specifikacijom)",
                            data=pdf_data,
                            file_name=f"izvestaj_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"PDF greška: {e}")
            
            with col3:
                st.markdown("### 📧 Pošalji")
                with st.form("email_form"):
                    email_primalac = st.text_input("Email")
                    col1, col2 = st.columns(2)
                    with col1:
                        posalji_pdf = st.checkbox("PDF", value=True)
                    with col2:
                        poslati_excel = st.checkbox("Excel", value=True)
                    submitted = st.form_submit_button("📩 Pošalji")
                    
                    if submitted:
                        if not email_primalac:
                            st.error("Unesite email")
                        else:
                            excel_attach = excel_data if poslati_excel else None
                            pdf_attach = pdf_data if posalji_pdf else None
                            with st.spinner("Šaljem..."):
                                success, msg = posalji_email(email_primalac, excel_attach, pdf_attach)
                                if success:
                                    st.success(msg)
                                else:
                                    st.error(msg)
