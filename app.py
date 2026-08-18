import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io

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

def ucitaj_podatke(fajl):
    xl = pd.ExcelFile(fajl)
    if SHEET_RECIKLAZA not in xl.sheet_names:
        st.error(f"Sheet '{SHEET_RECIKLAZA}' nije pronađen!")
        return None, None
    if SHEET_REKLAMACIJE not in xl.sheet_names:
        st.error(f"Sheet '{SHEET_REKLAMACIJE}' nije pronađen!")
        return None, None
    reciklaza = pd.read_excel(fajl, sheet_name=SHEET_RECIKLAZA, header=0, dtype=str)
    reklamacije = pd.read_excel(fajl, sheet_name=SHEET_REKLAMACIJE, header=0, dtype=str)
    return reciklaza, reklamacije

def pripremi_lookup(reklamacije):
    """Pravi rečnik sa ključevima kao string (bez razmaka)"""
    rek = reklamacije.copy()
    # Ukloni sve razmake i pretvori u string
    rek["_kljuc"] = rek["Reklamacija: Broj reklamacije"].astype(str).str.replace(" ", "").str.strip()
    return rek.set_index("_kljuc")

def spoji(reciklaza, lookup):
    df = reciklaza.copy()
    # I ovde ukloni razmake
    df["_kljuc"] = df["Broj reklamacije"].astype(str).str.replace(" ", "").str.strip()
    
    for out_kol, src_kol in MAPPING.items():
        if src_kol not in lookup.columns:
            df[out_kol] = None
            continue
        vrednosti = []
        for kljuc in df["_kljuc"]:
            if pd.isna(kljuc) or kljuc == "nan" or kljuc == "":
                vrednosti.append(None)
            elif kljuc in lookup.index:
                vrednosti.append(lookup.at[kljuc, src_kol])
            else:
                vrednosti.append(None)
        df[out_kol] = vrednosti
    
    mask = df["_kljuc"].notna() & (df["_kljuc"] != "nan") & (df["_kljuc"] != "")
    pronadjeno = 0
    nepronadjeno = 0
    for kljuc in df.loc[mask, "_kljuc"]:
        if kljuc in lookup.index:
            pronadjeno += 1
        else:
            nepronadjeno += 1
    
    df = df.drop(columns=["_kljuc"])
    
    extras = [c for c in df.columns if c not in KOLONE_OUTPUT]
    finalne = [c for c in KOLONE_OUTPUT if c in df.columns] + extras
    return df[finalne], pronadjeno, nepronadjeno

def formatiraj_i_sacuvaj(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "zbirno"
    
    lookup_kolone = set(MAPPING.keys())
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
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
    
    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if pd.isna(value) or (isinstance(value, str) and value.lower() in ["nan", "none", ""]):
                cell.value = None
            else:
                cell.value = value
            col_name = df.columns[col_idx - 1]
            cell.font = Font(name="Arial", size=10)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
            if col_name in lookup_kolone:
                cell.fill = PatternFill("solid", fgColor="EEF3FB")
            elif row_idx % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F5F8FD")
    
    for col_idx, kol in enumerate(df.columns, start=1):
        max_len = len(str(kol))
        for row_idx in range(2, min(len(df) + 2, 52)):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val:
                max_len = max(max_len, min(len(str(val)), 40))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3
    
    ws.freeze_panes = "A2"
    
    legenda_row = len(df) + 3
    ws.cell(row=legenda_row, column=1, value="Legenda:").font = Font(bold=True, name="Arial", size=9)
    ws.cell(row=legenda_row + 1, column=1, value="■ Tamno plava = ručno uneti podaci").font = Font(name="Arial", size=9, color=HEADER_BG)
    ws.cell(row=legenda_row + 2, column=1, value="■ Svetlo plava = automatski povučeno").font = Font(name="Arial", size=9, color=LOOKUP_FONT)
    ws.cell(row=legenda_row + 4, column=1, value=f"Generisano: {datetime.now().strftime('%d.%m.%Y %H:%M')}").font = Font(italic=True, name="Arial", size=9, color="888888")
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ============================
# STREAMLIT UI
# ============================

st.set_page_config(page_title="Spajanje reciklaže", layout="wide")
st.title("🔄 Spajanje Excel fajlova za reciklažu")
st.markdown("""
    Uploaduj Excel fajl koji sadrži **dva sheeta**:
    - **`reciklaža`** – tvoji ručno uneti podaci
    - **`reklamacije ukupno`** – baza reklamacija
""")

uploaded_file = st.file_uploader("Izaberi Excel fajl", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Učitavam podatke..."):
        reciklaza, reklamacije = ucitaj_podatke(uploaded_file)
    
    if reciklaza is not None and reklamacije is not None:
        st.success(f"✅ Učitano: {len(reciklaza)} redova u 'reciklaža', {len(reklamacije)} redova u 'reklamacije ukupno'")
        
        with st.spinner("Spajam podatke..."):
            lookup = pripremi_lookup(reklamacije)
            df_final, pronadjeno, nepronadjeno = spoji(reciklaza, lookup)
        
        st.info(f"✅ Pronađeno: **{pronadjeno}** reklamacija, ⚠️ Nije pronađeno: **{nepronadjeno}**")
        
        # Ako ima nepronađenih, prikaži ih
        if nepronadjeno > 0:
            with st.expander("🔍 Prikaži nepronađene brojeve reklamacija"):
                lookup_kolone = list(MAPPING.keys())
                # Uzmi redove gde su sve lookup kolone prazne
                mask = df_final[lookup_kolone].isna().all(axis=1)
                nepronadjeni = df_final.loc[mask, ["Broj reklamacije"] + lookup_kolone]
                st.dataframe(nepronadjeni, use_container_width=True)
                st.warning(f"⚠️ {len(nepronadjeni)} redova nije pronađeno u bazi reklamacija")
        
        st.subheader("📊 Pregled rezultata (prvih 5 redova)")
        st.dataframe(df_final.head(), use_container_width=True)
        
        excel_data = formatiraj_i_sacuvaj(df_final)
        st.download_button(
            label="📥 Preuzmi obrađen Excel fajl",
            data=excel_data,
            file_name=f"Reciklaza_obradjeno_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
