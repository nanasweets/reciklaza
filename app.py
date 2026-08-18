import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import io
import re

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

def cisti_broj(vrednost):
    """Izvlači samo cifre iz vrednosti"""
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
        return None, None, None, None
    if SHEET_REKLAMACIJE not in xl.sheet_names:
        st.error(f"Sheet '{SHEET_REKLAMACIJE}' nije pronađen!")
        return None, None, None, None
    reciklaza = pd.read_excel(fajl, sheet_name=SHEET_RECIKLAZA, header=0)
    reklamacije = pd.read_excel(fajl, sheet_name=SHEET_REKLAMACIJE, header=0)
    kolona_broja = pronadji_kolonu_broja(reklamacije)
    return reciklaza, reklamacije, kolona_broja

def pripremi_lookup(reklamacije, kolona_broja):
    rek = reklamacije.copy()
    rek["_kljuc"] = rek[kolona_broja].apply(cisti_broj)
    rek = rek[rek["_kljuc"] != ""]
    return rek.set_index("_kljuc"), kolona_broja

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
    mask = df["_kljuc"] != ""
    pronadjeno = 0
    nepronadjeno = 0
    nepronadjeni_brojevi = []
    for kljuc in df.loc[mask, "_kljuc"]:
        if kljuc in lookup.index:
            pronadjeno += 1
        else:
            nepronadjeno += 1
            nepronadjeni_brojevi.append(kljuc)
    df = df.drop(columns=["_kljuc"])
    extras = [c for c in df.columns if c not in KOLONE_OUTPUT]
    finalne = [c for c in KOLONE_OUTPUT if c in df.columns] + extras
    return df[finalne], pronadjeno, nepronadjeno, nepronadjeni_brojevi

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
        reciklaza, reklamacije, kolona_broja = ucitaj_podatke(uploaded_file)
    
    if reciklaza is not None and reklamacije is not None:
        st.success(f"✅ Učitano: {len(reciklaza)} redova u 'reciklaža', {len(reklamacije)} redova u 'reklamacije ukupno'")
        
        # 🔥 NOVO: Detaljna provera formata brojeva
        with st.expander("🔍 DETALJNA PROVERA BROJEVA REKLAMACIJA", expanded=True):
            st.markdown("### 📋 Brojevi iz 'reciklaža' (prvih 10):")
            
            # Prikaz originalnih vrednosti i očišćenih
            brojevi_iz_reciklaze = reciklaza["Broj reklamacije"].head(10).tolist()
            brojevi_iz_reciklaze_cisti = [cisti_broj(x) for x in brojevi_iz_reciklaze]
            
            df_poredenje = pd.DataFrame({
                "Originalna vrednost": brojevi_iz_reciklaze,
                "Očišćena vrednost (cifre)": brojevi_iz_reciklaze_cisti,
                "Dužina": [len(str(x)) for x in brojevi_iz_reciklaze_cisti]
            })
            st.dataframe(df_poredenje, use_container_width=True)
            
            st.markdown("### 📋 Brojevi iz 'reklamacije ukupno' (prvih 10):")
            
            # Uzmi prvih 10 iz reklamacije ukupno
            brojevi_iz_reklamacija = reklamacije[kolona_broja].head(10).tolist()
            brojevi_iz_reklamacija_cisti = [cisti_broj(x) for x in brojevi_iz_reklamacija]
            
            df_poredenje2 = pd.DataFrame({
                "Originalna vrednost": brojevi_iz_reklamacija,
                "Očišćena vrednost (cifre)": brojevi_iz_reklamacija_cisti,
                "Dužina": [len(str(x)) for x in brojevi_iz_reklamacija_cisti]
            })
            st.dataframe(df_poredenje2, use_container_width=True)
            
            # Pokušaj da pronađeš prvi broj iz reciklaže u reklamacijama
            if len(brojevi_iz_reciklaze_cisti) > 0 and brojevi_iz_reciklaze_cisti[0]:
                prvi_broj = brojevi_iz_reciklaze_cisti[0]
                postoji = prvi_broj in [cisti_broj(x) for x in reklamacije[kolona_broja].tolist()]
                if postoji:
                    st.success(f"✅ Broj '{prvi_broj}' POSTOJI u 'reklamacije ukupno'")
                else:
                    st.error(f"❌ Broj '{prvi_broj}' NE POSTOJI u 'reklamacije ukupno'")
        
        with st.spinner("Spajam podatke..."):
            lookup, _ = pripremi_lookup(reklamacije, kolona_broja)
            df_final, pronadjeno, nepronadjeno, nepronadjeni_brojevi = spoji(reciklaza, lookup)
        
        st.info(f"✅ Pronađeno: **{pronadjeno}** reklamacija, ⚠️ Nije pronađeno: **{nepronadjeno}**")
        
        if nepronadjeno > 0:
            with st.expander("🔍 Prikaži nepronađene brojeve reklamacije"):
                st.write(f"**Prvih 20 od {len(nepronadjeni_brojevi)} nepronađenih brojeva:**")
                st.write(nepronadjeni_brojevi[:20])
                
                lookup_kolone = list(MAPPING.keys())
                mask = df_final[lookup_kolone].isna().all(axis=1)
                nepronadjeni_redovi = df_final.loc[mask, ["Broj reklamacije"] + lookup_kolone]
                st.dataframe(nepronadjeni_redovi.head(20), use_container_width=True)
                st.warning(f"⚠️ Ukupno {len(nepronadjeni_redovi)} redova nije pronađeno u bazi reklamacija")
        
        st.subheader("📊 Pregled rezultata (prvih 5 redova)")
        st.dataframe(df_final.head(), use_container_width=True)
        
        excel_data = formatiraj_i_sacuvaj(df_final)
        st.download_button(
            label="📥 Preuzmi obrađen Excel fajl",
            data=excel_data,
            file_name=f"Reciklaza_obradjeno_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
