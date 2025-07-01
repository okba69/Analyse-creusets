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

# Seuils figés
CLEAN_THRESHOLD   = 60
SET_THRESHOLD     = 80
ANOMALY_THRESHOLD = 70

# Fonctions métier

def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    data_cols = df.columns[1:57]
    df[data_cols] = df[data_cols].applymap(
        lambda x: ""
        if pd.notna(x) and isinstance(x, (int, float)) and (x in [99, 100] or x < CLEAN_THRESHOLD)
        else x
    )
    for i in range(len(df)):
        if (df.loc[i, data_cols] == "").sum() >= 40:
            df.loc[i, data_cols] = ""
    for i in range(len(df)-1):
        cur = pd.to_numeric(df.loc[i, data_cols], errors='coerce')
        nxt = pd.to_numeric(df.loc[i+1, data_cols], errors='coerce')
        if ((cur - nxt) >= 15).sum() >= 15:
            df.loc[i, data_cols] = ""
    for col in data_cols:
        for i in range(1, len(df)-1):
            if df.at[i-1, col] == "" and df.at[i+1, col] == "":
                df.at[i, col] = ""
    return df


def detect_sets_and_anomalies(df: pd.DataFrame):
    data_cols = df.columns[1:57]
    set_starts = []
    # repérer début de set
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
    # flags pour chaque emplacement: si on est passé sous 70 après set
    drop_flags = {i+1: False for i in range(len(df.columns)-1)}

    for idx in range(len(df)):
        vals = pd.to_numeric(df.loc[idx, data_cols], errors='coerce')
        cnt80 = (vals > SET_THRESHOLD).sum()
        cnt70 = (vals < ANOMALY_THRESHOLD).sum()

        if cnt80 >= 40 and not in_set:
            set_count += 1
            # init flags après nouveau set
            drop_flags = {i+1: False for i in range(len(data_cols))}
            try:
                ts = pd.to_datetime(df.iloc[idx,0])
                date = ts.strftime("%d/%m/%Y")
            except:
                date = "Inconnu"
            if last > 0 and current:
                anomalies_by_set[last] = sorted(current)
            meta.append({"Set": set_count, "Date": date})
            last = set_count
            current = set()
            in_set = True
        elif in_set:
            # sortir du set
            if cnt70 >= len(data_cols) - ANOMALY_THRESHOLD:  # ajuster si nécessaire
                in_set = False
        # gestion drop flags
        for i, col in enumerate(data_cols, start=1):
            v = pd.to_numeric(df.at[idx, col], errors='coerce')
            if v < ANOMALY_THRESHOLD:
                drop_flags[i] = True
        # détection anomalies selon flags
        if in_set:
            for i, col in enumerate(data_cols, start=1):
                if not drop_flags[i]:
                    continue
                v = pd.to_numeric(df.at[idx, col], errors='coerce')
                if v >= SET_THRESHOLD:
                    nv = None
                    if idx+1 < len(df):
                        nv = pd.to_numeric(df.at[idx+1, col], errors='coerce')
                    if nv is not None and pd.notna(nv) and nv < SET_THRESHOLD:
                        continue
                    if i not in current:
                        current.add(i)
                        anomaly_cells.append((idx+2, i+1))
    # dernier set
    if last > 0 and current:
        anomalies_by_set[last] = sorted(current)

    return set_starts, anomaly_cells, meta, anomalies_by_set


def to_excel(df, set_starts, anomaly_cells, meta, anomalies_by_set):
    wb = Workbook()
    ws = wb.active
    ws.title = "Données nettoyées"
    for r, row in enumerate(df.itertuples(index=False), start=1):
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
    orange = PatternFill("solid", fgColor="FFA500")
    blue   = PatternFill("solid", fgColor="ADD8E6")
    for row_idx in set_starts:
        for cell in ws[row_idx]:
            cell.fill = orange
    for row, col in anomaly_cells:
        ws.cell(row=row, column=col).fill = blue
    ws2 = wb.create_sheet("Résumé")
    ws2.append(["Set", "Date", "Nb anomalies"]);
    total = 0
    for m in meta:
        s = m["Set"]
        nb = len(anomalies_by_set.get(s, []))
        total += nb
        ws2.append([s, m["Date"], nb])
    ws2.append([])
    ws2.append(["Total anomalies", total])
    # ajout anomalies par emplacement
    counter = Counter([x for vals in anomalies_by_set.values() for x in vals])
    ws2.append([])
    ws2.append(["Emplacement","Occurrences"])
    for loc, cnt in sorted(counter.items()):
        ws2.append([loc, cnt])
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
uploaded = st.file_uploader("Téléverse ton fichier .xlsx", type=["xlsx"])
analyse = st.button("Analyser")
if uploaded and analyse:
    df = pd.read_excel(uploaded, engine='openpyxl')
    df_clean = clean_data(df.copy())
    set_starts, anomaly_cells, meta, anomalies_by_set = detect_sets_and_anomalies(df_clean)
    # récapitulatif sets
    recap = pd.DataFrame([{"Set": m["Set"], "Date": m["Date"], "Nb anomalies": len(anomalies_by_set.get(m["Set"], []))} for m in meta]).set_index("Set")
    st.markdown("## Récapitulatif des sets")
    st.table(recap)
    total = recap["Nb anomalies"].sum()
    st.markdown(f"**Total anomalies sur tous les sets : {total}**")
    # anomalies par emplacement
    counter = Counter([x for vals in anomalies_by_set.values() for x in vals])
    df_all = pd.DataFrame(sorted(counter.items()), columns=["Emplacement","Occurrences"]).set_index("Emplacement")
    st.markdown("## Anomalies par emplacement")
    st.table(df_all)
    # téléchargement
    excel_bytes = to_excel(df_clean, set_starts, anomaly_cells, meta, anomalies_by_set)
    st.download_button("📥 Télécharger le rapport Excel", data=excel_bytes, file_name="analyse_creusets_report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
