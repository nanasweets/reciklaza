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
import base64
import tempfile

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
# PLOTLY -> SLIKA (za PDF)
# ============================

def plotly_to_image(fig, width=800, height=500):
    """Konvertuje plotly figuru u sliku za PDF"""
    # Sačuvaj kao PNG u BytesIO
    img_bytes = fig.to_image(format="png", width=width, height=height, scale=1)
    img_buffer = io.BytesIO(img_bytes)
    img_buffer.seek(0)
    return img_buffer

def plotly_to_reportlab_image(fig, width=12*cm, height=8*cm):
    """Konvertuje plotly figuru u ReportLab Image objekat"""
    img_buffer = plotly_to_image(fig, width=800, height=500)
    return Image(img_buffer, width=width, height=height)

# ============================
# PDF GENERATOR (sa plotly slikama)
# ============================

def generisi_pdf(df):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
    styles = getSampleStyleSheet()
    elements = []
    
    # Stilovi
    naslov_style = ParagraphStyle('Naslov', parent=styles['Heading1'], fontName=FONT_NAME, fontSize=24, textColor=colors.HexColor('#1F4E79'), alignment=1, spaceAfter=20, spaceBefore=30)
    sekcija_style = ParagraphStyle('Sekcija', parent=styles['Heading3'], fontName=FONT_NAME, fontSize=14, textColor=colors.HexColor('#2E75B6'), spaceAfter=8, spaceBefore=12)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, spaceAfter=4)
    bold_style = ParagraphStyle('Bold', parent=styles['Normal'], fontName=FONT_NAME, fontSize=10, textColor=colors.HexColor('#1F4E79'), spaceAfter=4)
    
    # Podaci
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
    
    # ============================================================
    # 1. KLASIFIKACIJA RECIKLAŽE
    # ============================================================
    elements.append(Paragraph("1. Klasifikacija reciklaže", sekcija_style))
    if "Klasifikacija reciklaže" in df.columns:
        klas = df["Klasifikacija reciklaže"].value_counts().head(10).reset_index()
        klas.columns = ["Klasifikacija", "Broj"]
        table_data = [["Klasifikacija", "Broj"]] + [[str(r["Klasifikacija"])[:50], str(r["Broj"])] for _, r in klas.iterrows()]
        t = Table(table_data, colWidths=[10*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon - plotly pie
        fig = px.pie(klas.head(8), values="Broj", names="Klasifikacija", title="Klasifikacija reciklaže")
        fig.update_traces(textposition='inside', textinfo='percent+label')
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # ============================================================
    # 2. KLASIFIKACIJA ŠTETE
    # ============================================================
    elements.append(Paragraph("2. Klasifikacija štete", sekcija_style))
    if "Klasifikacija štete" in df.columns:
        steta = df["Klasifikacija štete"].value_counts().head(10).reset_index()
        steta.columns = ["Klasifikacija štete", "Broj"]
        table_data = [["Klasifikacija štete", "Broj"]] + [[str(r["Klasifikacija štete"])[:50], str(r["Broj"])] for _, r in steta.iterrows()]
        t = Table(table_data, colWidths=[10*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon - plotly bar
        fig = px.bar(steta.head(8), x="Klasifikacija štete", y="Broj", title="Klasifikacija štete")
        fig.update_layout(xaxis_tickangle=-45)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # ============================================================
    # 3. TOP 10 BRENDOVA
    # ============================================================
    elements.append(Paragraph("3. Top 10 brendova", sekcija_style))
    if "Brend" in df.columns:
        brend = df["Brend"].value_counts().head(10).reset_index()
        brend.columns = ["Brend", "Broj"]
        table_data = [["Brend", "Broj"]] + [[str(r["Brend"]), str(r["Broj"])] for _, r in brend.iterrows()]
        t = Table(table_data, colWidths=[10*cm, 4*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon - plotly bar
        fig = px.bar(brend, x="Brend", y="Broj", title="Top 10 brendova")
        fig.update_layout(xaxis_tickangle=-45)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=9*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # ============================================================
    # 4. VREMENSKI TREND
    # ============================================================
    elements.append(Paragraph("4. Vremenski trend", sekcija_style))
    if "Datum obrade" in df.columns:
        df_datum = df.copy()
        df_datum["Datum obrade"] = pd.to_datetime(df_datum["Datum obrade"], errors="coerce")
        df_datum["Mesec"] = df_datum["Datum obrade"].dt.to_period("M")
        trend = df_datum.groupby("Mesec").size().reset_index()
        trend.columns = ["Mesec", "Broj"]
        trend["Mesec"] = trend["Mesec"].astype(str)
        table_data = [["Mesec", "Broj"]] + [[str(r["Mesec"]), str(r["Broj"])] for _, r in trend.iterrows()]
        t = Table(table_data, colWidths=[8*cm, 6*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 9), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon - plotly line
        fig = px.line(trend, x="Mesec", y="Broj", title="Vremenski trend", markers=True)
        img = plotly_to_reportlab_image(fig, width=14*cm, height=8*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # ============================================================
    # 5. PARETO ANALIZA
    # ============================================================
    elements.append(Paragraph("5. Pareto analiza (80/20)", sekcija_style))
    if "Nabavna cena" in df.columns and "Količina" in df.columns:
        df["Ukupna vrednost"] = df["Nabavna cena"] * df["Količina"]
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
        
        # Grafikon - plotly bar sa linijom
        pareto_top = pareto.head(20).reset_index()
        pareto_top.index = range(1, len(pareto_top) + 1)
        fig = px.bar(pareto_top, x=pareto_top.index, y="Ukupna vrednost", title="Pareto analiza - Top 20 artikala")
        fig.add_hline(y=pareto["Ukupna vrednost"].sum() * 0.8, line_dash="dash", line_color="red", annotation_text="80% granica")
        fig.update_layout(xaxis_title="Redni broj artikla", yaxis_title="Vrednost (RSD)")
        img = plotly_to_reportlab_image(fig, width=14*cm, height=8*cm)
        elements.append(img)
    elements.append(PageBreak())
    
    # ============================================================
    # 6. TOP 10 NAJSKUPLJIH
    # ============================================================
    elements.append(Paragraph("6. Top 10 najskupljih artikala", sekcija_style))
    if "Nabavna cena" in df.columns:
        naj = df.nlargest(10, "Nabavna cena")[["Broj reklamacije", "Naziv artikla", "Nabavna cena", "Brend"]]
        table_data = [["Broj reklamacije", "Naziv", "Cena (RSD)", "Brend"]]
        for _, r in naj.iterrows():
            table_data.append([str(r["Broj reklamacije"])[:12], str(r["Naziv artikla"])[:25], f"{r['Nabavna cena']:,.2f}", str(r["Brend"])])
        t = Table(table_data, colWidths=[2.5*cm, 5*cm, 3*cm, 2.5*cm])
        t.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), FONT_NAME), ('FONTSIZE', (0,0), (-1,-1), 7), ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E79')), ('TEXTCOLOR', (0,0), (-1,0), colors.white), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CCCCCC'))]))
        elements.append(t)
        elements.append(Spacer(1, 0.5*cm))
        
        # Grafikon - plotly horizontal bar
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
        
        col1, col2 = st.columns(2)
        with col1:
            excel_data = formatiraj_i_sacuvaj(df)
            st.download_button("📥 Preuzmi Excel", excel_data, f"izvestaj_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True)
        with col2:
            with st.spinner("Generišem PDF..."):
                pdf_data = generisi_pdf(df)
                st.download_button("📄 Preuzmi PDF (sa grafikonima)", pdf_data, f"izvestaj_{datetime.now().strftime('%Y%m%d')}.pdf", use_container_width=True)
