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
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
from io import BytesIO

# ============================
# REGISTRACIJA FONTA (srpska slova)
# ============================

FONT_NAME = 'Helvetica'
try:
    font_path = os.path.join(os.path.dirname(__file__), 'DejaVuSans.ttf')
    font_bold_path = os.path.join(os.path.dirname(__file__), 'DejaVuSans-Bold.ttf')
    if os.path.exists(font_path) and os.path.exists(font_bold_path):
        pdfmetrics.registerFont(TTFont('DejaVu', font_path))
        pdfmetrics.registerFont(TTFont('DejaVu-Bold', font_bold_path))
        FONT_NAME = 'DejaVu'
        print("✅ DejaVu font učitan")
except:
    try:
        pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
        pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
        FONT_NAME = 'DejaVu'
        print("✅ DejaVu font učitan iz sistema")
    except:
        print("⚠️ DejaVu font nije pronađen, koristim Helvetica")

# ============================
# KONFIGURACIJA
# ============================

st.set_page_config(page_title="Izveštaj o reciklaži", layout="wide")

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
# EXCEL FUNKCIJE (skraćene)
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

def formatiraj_i_sacuvaj(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "zbirno"
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
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================
# PDF GENERATOR (sa srpskim slovima i grafikonima)
# ============================

def generisi_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    elements = []
    
    # Stilovi sa srpskim fontom
    naslov_style = ParagraphStyle('Naslov', parent=styles['Heading1'], fontName=FONT_NAME, fontSize=24, textColor=colors.HexColor('#1F4E79'), alignment=1, spaceAfter=20, spaceBefore=30)
    podnaslov_style = ParagraphStyle('Podnaslov', parent=styles['Heading2'], fontName=FONT_NAME, fontSize=16, textColor=colors.HexColor('#1F4E79'), spaceAfter=12, spaceBefore=15)
    sekcija_style = ParagraphStyle('Sekcija', parent=styles['Heading3'], fontName=FONT_NAME, fontSize=14, textColor=colors.HexColor('#2E75B6'), spaceAfter=8, spaceBefore=12)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, spaceAfter=4)
    bold_style = ParagraphStyle('Bold', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, textColor=colors.HexColor('#1F4E79'), spaceAfter=4)
    
    # Podaci
    ukupno = df["Količina"].sum() if "Količina" in df.columns else 0
    vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
    brendovi = df["Brend"].nunique() if "Brend" in df.columns else 0
    
    # NASLOVNA
    elements.append(Spacer(1, 4*cm))
    elements.append(Paragraph("Tehnomanija", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph("IZVEŠTAJ O RECIKLAŽI", naslov_style))
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}", normal_style))
    elements.append(Spacer(1, 2*cm))
    
    kpi_data = [["Ukupno artikala", f"{ukupno:,}"], ["Ukupna vrednost", f"{vrednost:,.2f} RSD"], ["Broj brendova", str(brendovi)], ["Broj redova", str(len(df))]]
    kpi_table = Table(kpi_data, colWidths=[7*cm, 5*cm])
    kpi_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 11), ('BACKGROUND', (0,0), (0,-1), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (0,-1), colors.white), ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#CCCCCC')), ('ALIGN', (1,0), (1,-1), 'RIGHT'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    elements.append(kpi_table)
    elements.append(PageBreak())
    
    # ============================================================
    # 1. KLASIFIKACIJA RECIKLAŽE
    # ============================================================
    elements.append(Paragraph("1. Klasifikacija reciklaže", sekcija_style))
    if "Klasifikacija reciklaže" in df.columns:
        klas_data = df["Klasifikacija reciklaže"].value_counts().head(10).reset_index()
        klas_data.columns = ["Klasifikacija", "Broj"]
        table_data = [["Klasifikacija", "Broj"]] + [[str(row["Klasifikacija"])[:50], str(row["Broj"])] for _, row in klas_data.iterrows()]
        klas_table = Table(table_data, colWidths=[10*cm, 4*cm])
        klas_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(klas_table)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon
        fig, ax = plt.subplots(figsize=(8, 5))
        klas_pie = df["Klasifikacija reciklaže"].value_counts().head(8)
        colors_pie = plt.cm.Blues_r([i/len(klas_pie) for i in range(len(klas_pie))])
        ax.pie(klas_pie.values, labels=klas_pie.index, autopct='%1.0f%%', colors=colors_pie, startangle=90)
        ax.set_title('Raspored po klasifikaciji reciklaže', fontsize=14, fontweight='bold')
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        elements.append(Image(img_buffer, width=15*cm, height=10*cm))
    elements.append(PageBreak())
    
    # ============================================================
    # 2. KLASIFIKACIJA ŠTETE
    # ============================================================
    elements.append(Paragraph("2. Klasifikacija štete", sekcija_style))
    if "Klasifikacija štete" in df.columns:
        steta_data = df["Klasifikacija štete"].value_counts().head(10).reset_index()
        steta_data.columns = ["Klasifikacija štete", "Broj"]
        table_data = [["Klasifikacija štete", "Broj"]] + [[str(row["Klasifikacija štete"])[:50], str(row["Broj"])] for _, row in steta_data.iterrows()]
        steta_table = Table(table_data, colWidths=[10*cm, 4*cm])
        steta_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(steta_table)
        elements.append(Spacer(1, 0.5*cm))
        
        fig, ax = plt.subplots(figsize=(8, 5))
        steta_bar = df["Klasifikacija štete"].value_counts().head(8)
        ax.barh(steta_bar.index, steta_bar.values, color=plt.cm.Reds_r([i/len(steta_bar) for i in range(len(steta_bar))]))
        ax.set_xlabel('Broj artikala')
        ax.set_title('Najčešće klasifikacije štete', fontsize=14, fontweight='bold')
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        elements.append(Image(img_buffer, width=15*cm, height=10*cm))
    elements.append(PageBreak())
    
    # ============================================================
    # 3. TOP 10 BRENDOVA
    # ============================================================
    elements.append(Paragraph("3. Top 10 brendova", sekcija_style))
    if "Brend" in df.columns:
        brend_data = df["Brend"].value_counts().head(10).reset_index()
        brend_data.columns = ["Brend", "Broj"]
        table_data = [["Brend", "Broj"]] + [[str(row["Brend"]), str(row["Broj"])] for _, row in brend_data.iterrows()]
        brend_table = Table(table_data, colWidths=[10*cm, 4*cm])
        brend_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(brend_table)
        elements.append(Spacer(1, 0.5*cm))
        
        fig, ax = plt.subplots(figsize=(8, 5))
        brend_bar = df["Brend"].value_counts().head(10)
        ax.bar(brend_bar.index, brend_bar.values, color=plt.cm.Blues([0.3 + 0.7*i/len(brend_bar) for i in range(len(brend_bar))]))
        ax.set_ylabel('Broj artikala')
        ax.set_title('Top 10 brendova po broju artikala', fontsize=14, fontweight='bold')
        plt.xticks(rotation=45, ha='right')
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        elements.append(Image(img_buffer, width=15*cm, height=8*cm))
    elements.append(PageBreak())
    
    # ============================================================
    # 4. VREMENSKI TREND
    # ============================================================
    elements.append(Paragraph("4. Vremenski trend", sekcija_style))
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
        trend_data = df_datum.groupby("Mesec").size().reset_index()
        trend_data.columns = ["Mesec", "Broj"]
        trend_data["Mesec"] = trend_data["Mesec"].astype(str)
        table_data = [["Mesec", "Broj"]] + [[str(row["Mesec"]), str(row["Broj"])] for _, row in trend_data.iterrows()]
        trend_table = Table(table_data, colWidths=[8*cm, 6*cm])
        trend_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(trend_table)
        elements.append(Spacer(1, 0.5*cm))
        
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(trend_data["Mesec"], trend_data["Broj"], marker='o', linewidth=2, color='#1F4E79')
        ax.fill_between(trend_data["Mesec"], trend_data["Broj"], alpha=0.3, color='#1F4E79')
        ax.set_ylabel('Broj artikala')
        ax.set_title('Broj artikala po mesecima', fontsize=14, fontweight='bold')
        ax.grid(True, linestyle='--', alpha=0.5)
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        elements.append(Image(img_buffer, width=15*cm, height=6*cm))
    elements.append(PageBreak())
    
    # ============================================================
    # 5. PARETO ANALIZA
    # ============================================================
    elements.append(Paragraph("5. Pareto analiza (80/20)", sekcija_style))
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
        pareto = df.sort_values("Ukupna vrednost", ascending=False)
        pareto["Kumulativni %"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
        table_data = [["Broj reklamacije", "Naziv artikla", "Vrednost (RSD)", "Kum.%"]]
        for _, row in pareto.head(15).iterrows():
            table_data.append([str(row["Broj reklamacije"])[:15], str(row["Naziv artikla"])[:30], f"{row['Ukupna vrednost']:,.0f}", f"{row['Kumulativni %']}%"])
        pareto_table = Table(table_data, colWidths=[3*cm, 6*cm, 3*cm, 2.5*cm])
        pareto_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 7), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(pareto_table)
        elements.append(Spacer(1, 0.3*cm))
        granica = pareto[pareto["Kumulativni %"] <= 80].shape[0]
        elements.append(Paragraph(f"📊 {granica} artikala čini 80% ukupne vrednosti", bold_style))
        
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
        elements.append(Image(img_buffer, width=15*cm, height=8*cm))
    elements.append(PageBreak())
    
    # ============================================================
    # 6. TOP 10 NAJSKUPLJIH
    # ============================================================
    elements.append(Paragraph("6. Top 10 najskupljih artikala", sekcija_style))
    if "Nabavna cena" in df.columns:
        najskuplji = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend"]]
        table_data = [["Broj reklamacije", "Naziv artikla", "Cena (RSD)", "Brend"]]
        for _, row in najskuplji.iterrows():
            table_data.append([str(row["Broj reklamacije"])[:15], str(row["Naziv artikla"])[:35], f"{row['Nabavna cena']:,.2f}", str(row["Brend"])])
        naj_table = Table(table_data, colWidths=[3*cm, 6*cm, 3*cm, 3*cm])
        naj_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 8), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC')), ('ALIGN', (2,1), (2,-1), 'RIGHT')]))
        elements.append(naj_table)
        elements.append(Spacer(1, 0.5*cm))
        
        fig, ax = plt.subplots(figsize=(8, 5))
        ax.barh(najskuplji["Naziv artikla"].str[:20], najskuplji["Nabavna cena"], color=plt.cm.Greens([0.3 + 0.7*i/10 for i in range(10)]))
        ax.set_xlabel('Cena (RSD)')
        ax.set_title('Top 10 najskupljih artikala', fontsize=14, fontweight='bold')
        img_buffer = BytesIO()
        plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight')
        img_buffer.seek(0)
        plt.close()
        elements.append(Image(img_buffer, width=15*cm, height=8*cm))
    
    # ZAVRŠNA
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

def posalji_email(primalac, excel_data, pdf_data=None, dodatne_adrese=None):
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
        if excel_data:
            excel_part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            excel_part.set_payload(excel_data)
            encoders.encode_base64(excel_part)
            excel_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.xlsx')
            msg.attach(excel_part)
        if pdf_data:
            pdf_part = MIMEBase('application', 'pdf')
            pdf_part.set_payload(pdf_data)
            encoders.encode_base64(pdf_part)
            pdf_part.add_header('Content-Disposition', f'attachment; filename=izvestaj_reciklaza_{datetime.now().strftime("%Y%m%d")}.pdf')
            msg.attach(pdf_part)
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender, password)
        svi_primaoci = [primalac] + (dodatne_adrese if dodatne_adrese else [])
        server.send_message(msg)
        server.quit()
        return True, "Email je uspešno poslat!"
    except Exception as e:
        return False, f"Greška: {str(e)}"

# ============================
# STREAMLIT UI
# ============================

with st.sidebar:
    st.title("⚙️ Opcije")
    dark_mode = st.toggle("🌙 Tamna tema", value=st.session_state.dark_mode)
    if dark_mode != st.session_state.dark_mode:
        st.session_state.dark_mode = dark_mode
        st.rerun()
    st.markdown("---")
    st.info("📊 **Izveštaj o reciklaži** v3.0\n\n© Tehnomanija")

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
        st.success(f"✅ Učitano: {len(reciklaza)} redova")
        st.info(f"📊 Spojeno: {len(df)} redova")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Pregled", "📈 Napredne metrike", "🏷️ Analize", "📋 Tabela", "📤 Izvoz"])
        
        with tab1:
            st.subheader("📊 Ključni pokazatelji")
            col1, col2, col3, col4 = st.columns(4)
            ukupno = df["Količina"].sum() if "Količina" in df.columns else 0
            vrednost = (df["Nabavna cena"] * df["Količina"]).sum() if "Nabavna cena" in df.columns and "Količina" in df.columns else 0
            col1.metric("📦 Ukupno artikala", f"{ukupno:,}")
            col2.metric("💰 Ukupna vrednost", f"{vrednost:,.2f} RSD")
            col3.metric("🏷️ Broj brendova", df["Brend"].nunique() if "Brend" in df.columns else 0)
            col4.metric("📂 Robnih grupa", df["Robna grupa"].nunique() if "Robna grupa" in df.columns else 0)
            st.markdown("---")
            st.subheader("📊 Klasifikacija reciklaže")
            if "Klasifikacija reciklaže" in df.columns:
                klas = df["Klasifikacija reciklaže"].value_counts().reset_index()
                klas.columns = ["Klasifikacija", "Broj"]
                fig = px.pie(klas, values="Broj", names="Klasifikacija", title="Klasifikacija reciklaže")
                st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            st.subheader("📈 Napredne metrike")
            if "Nabavna cena" in df.columns and "Količina" in df.columns:
                df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
                pareto = df.sort_values("Ukupna vrednost", ascending=False)
                pareto["Kumulativni %"] = (pareto["Ukupna vrednost"].cumsum() / pareto["Ukupna vrednost"].sum() * 100).round(2)
                fig = px.bar(pareto.head(20), x="Naziv artikla", y="Ukupna vrednost", title="Pareto analiza")
                st.plotly_chart(fig, use_container_width=True)
        
        with tab5:
            st.subheader("📤 Izvoz")
            col1, col2 = st.columns(2)
            with col1:
                excel_data = formatiraj_i_sacuvaj(df)
                st.download_button("📥 Excel", excel_data, f"izvestaj_{datetime.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with col2:
                with st.spinner("Generišem PDF..."):
                    pdf_data = generisi_pdf(df)
                    st.download_button("📄 PDF", pdf_data, f"izvestaj_{datetime.now().strftime('%Y%m%d')}.pdf", "application/pdf", use_container_width=True)
            st.markdown("---")
            st.subheader("📧 Pošalji emailom")
            with st.form("email_form"):
                email_primalac = st.text_input("Email")
                submitted = st.form_submit_button("📩 Pošalji")
                if submitted and email_primalac:
                    with st.spinner("Šaljem..."):
                        success, msg = posalji_email(email_primalac, excel_data, pdf_data)
                        if success:
                            st.success(msg)
                        else:
                            st.error(msg)
