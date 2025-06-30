import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import time
from collections import Counter

# Configuration de la page
st.set_page_config(page_title="Analyse Creusets", layout="wide")
st.title("🔍 Analyse et détection de sets/anomalies")
st.markdown("Dépose ton fichier Excel, puis clique sur **Analyser**. Les seuils sont préconfigurés.")

# Seuils fixes
CLEAN_THRESHOLD   = 60
SET_THRESHOLD     = 80
ANOMALY_THRESHOLD = 70

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    data_cols = df.columns[1:57]
    df[data_cols] = df[data_cols].applymap(
        lambda x: "" if pd.notna(x) and isinstance(x, (int, float)) and (x in [99, 100] or x < CLEAN_THRESHOLD) else x
    )
    # Suppression des lignes vides
    for i in range(len(df)):
        if (df.loc[i, data_cols] == "").sum() >= 40:
            df.loc[i, data_cols] = ""
    # Anomalies de chute brutale
    for i in range(len(df) - 1):
        cur = pd.to_numeric(df.loc[i, data_cols], errors='coerce')
        nxt = pd.to_numeric(df.loc[i+1, data_cols], errors='coerce')
        if ((cur - nxt) >= 15).sum() >= 15:
            df.loc[i, data_cols] = ""
    # Nettoyage vertical
    for col in data_cols:
        for i in range(1, len(df)-1):
            if df.at[i-1, col] == "" and df.at[i+1, col] == "":
                df.at[i, col] = ""
    return df

def detect_sets_and_anomalies(df: pd.DataFrame):
    data_cols = df.columns[1:57]
    set_starts = []
    # Repérer début de set
    for idx in range(len(df)):
        vals = pd.to_numeric(df.loc[idx, data_cols], errors='coerce')
        if (vals > SET_THRESHOLD).sum() >= 40:
            if idx > 0:
                df.loc[idx-1, data_cols] = ""
            set_starts.append(idx+2)

    set_count = 0
    in_set = False
    anomalies_by_set = {}
    anomaly_cells = []
    meta = []
    current = set()
    last = 0

    for idx in range(len(df)):
        vals = pd.to_numeric(df.loc[idx, data_cols], errors='coerce')
        cnt80 = (vals > SET_THRESHOLD).sum()
        cnt70 = (vals > ANOMALY_THRESHOLD).sum()

        if not in_set:
            if cnt80 >= 40:
                set_count += 1
                try:
                    ts = pd.to_datetime(df.iloc[idx,0])
                    date = ts.strftime("%d/%m/%Y")
                except:
                    date = "Inconnu"

                # Sauvegarde anomalies du set précédent
                if last > 0 and current:
                    anomalies_by_set[last] = sorted(current)
                meta.append({
                    "Set": set_count,
                    "Date": date
                })
                last = set_count
                current = set()
                in_set = True
            else:
                # Détection anomalies ponctuelles
                for ci, col in enumerate(data_cols):
                    v = pd.to_numeric(df.at[idx, col], errors='coerce')
                    if v >= SET_THRESHOLD:
                        # Annulation si ligne suivante < SET_THRESHOLD
                        if idx+1 < len(df):
                            nv = pd.to_numeric(df.at[idx+1, col], errors='coerce')
                            if pd.notna(nv) and nv < SET_THRESHOLD:
                                continue
                        colnum = ci + 1
                        if colnum not in current:
                            current.add(colnum)
                            anomaly_cells.append((idx+2, ci+2))
        else:
            if cnt70 < 40:
                in_set = False

    # Dernier set
    if last > 0 and current:
        anomalies_by_set[last] = sorted(current)

    return set_starts, anomaly_cells, meta, anomalies_by_set

def to_excel(df, set_starts, anomaly_cells, meta, byset, top10):
    wb = Workbook()
    ws = wb.active
    ws.title = "Données nettoyées"
    # Écriture des données nettoyées
    for r, row in enumerate(df.itertuples(index=False), start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)

    # Styles
    orange = PatternFill("solid", fgColor="FFA500")
    blue   = PatternFill("solid", fgColor="ADD8E6")

    # Coloration des débuts de set
    for row_idx in set_starts:
        for cell in ws[row_idx]:
            cell.fill = orange
    # Coloration des anomalies
    for row, col in anomaly_cells:
        ws.cell(row=row, column=col).fill = blue

    # Feuille Résumé
    ws2 = wb.create_sheet("Résumé")
    ws2.append(["Set", "Date", "Nb anomalies"])
    total_anomalies = 0
    for m in meta:
        s = m["Set"]
        nb = len(byset.get(s, []))
        total_anomalies += nb
        ws2.append([s, m.get("Date", ""), nb])
    # Ligne total
    ws2.append([])
    ws2.append(["Total", "", total_anomalies])

    # Top 10 Emplacements
    ws2.append([])
    ws2.append(["TOP 10 Emplacements","Occurrences"])
    for val, cnt in top10:
        ws2.append([val, cnt])

    # Ajustement des largeurs
    ws.column_dimensions["A"].width = 20
    for i in range(2, len(df.columns)+1):
        ws.column_dimensions[get_column_letter(i)].width = 5.5
    for col_cells in ws2.columns:
        ws2.column_dimensions[get_column_letter(col_cells[0].column)].width = 15

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# Interface Streamlit
uploaded = st.file_uploader("Téléverse ton fichier .xlsx", type=["xlsx"] )
analyse = st.button("Analyser")

if uploaded and analyse:
    df = pd.read_excel(uploaded, engine="openpyxl")
    df_clean = clean_data(df.copy())
    srows, acells, meta, byset = detect_sets_and_anomalies(df_clean)

    # Calcul Top 10 et Top 5\    all_anoms = [x for vals in byset.values() for x in vals]
    top10 = Counter(all_anoms).most_common(10)
    top5 = top10[:5]

    # Récapitulatif des sets
    recap = pd.DataFrame([
        {"Set": m["Set"], "Date": m["Date"], "Nb anomalies": len(byset.get(m["Set"], []))}
        for m in meta
    ]).set_index("Set")
    st.markdown("## Récapitulatif des sets")
    st.table(recap)

    # Total anomalies affiché sous le tableau
    total_anoms = recap["Nb anomalies"].sum()
    st.markdown(f"**Total anomalies : {total_anoms}**")

    # Top 5 Emplacements
    st.markdown("## Top 5 des emplacements changés le plus souvent")
    df_top5 = pd.DataFrame(top5, columns=["Emplacement","Occurrences"]).set_index("Emplacement")
    st.table(df_top5)

    # Téléchargement Excel
    excel_bytes = to_excel(df_clean, srows, acells, meta, byset, top10)
    st.download_button(
        "📥 Télécharger le rapport Excel",
        data=excel_bytes,
        file_name="analyse_creusets_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
