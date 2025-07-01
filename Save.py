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
    data_cols = list(df.columns[1:57])
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
    # drapeaux pour chute sous seuil
    dropped_flags = {ci: False for ci in range(len(data_cols))}

    for idx in range(len(df)):
        vals = pd.to_numeric(df.loc[idx, data_cols], errors='coerce')
        cnt80 = (vals > SET_THRESHOLD).sum()
        cnt70 = (vals > ANOMALY_THRESHOLD).sum()

        # Mettre à jour dropped_flags pour chaque colonne
        for ci, col in enumerate(data_cols):
            v = pd.to_numeric(df.at[idx, col], errors='coerce')
            if pd.notna(v) and v < ANOMALY_THRESHOLD:
                dropped_flags[ci] = True

        if not in_set:
            if cnt80 >= 40:
                # début de set
                set_count += 1
                try:
                    ts = pd.to_datetime(df.iloc[idx,0])
                    date = ts.strftime("%d/%m/%Y")
                except:
                    date = "Inconnu"

                # sauvegarde set précédent
                if last > 0 and current:
                    anomalies_by_set[last] = sorted(current)
                # reset anomalies pour nouveau set
                meta.append({"Set": set_count, "Date": date})
                last = set_count
                current = set()
                in_set = True
                # reset dropped_flags pour nouveau set
                dropped_flags = {ci: False for ci in range(len(data_cols))}
            else:
                # détection anomalies uniquement après chute sous 70
                for ci, col in enumerate(data_cols):
                    if not dropped_flags[ci]:
                        continue
                    v = pd.to_numeric(df.at[idx, col], errors='coerce')
                    if pd.notna(v) and v >= SET_THRESHOLD:
                        # annulation si valeur suivante < seuil set
                        if idx+1 < len(df):
                            nv = pd.to_numeric(df.at[idx+1, col], errors='coerce')
                            if pd.notna(nv) and nv < SET_THRESHOLD:
                                continue
                        colnum = ci + 1
                        if colnum not in current:
                            current.add(colnum)
                            anomaly_cells.append((idx+2, ci+2))
        else:
            # fin de set
            if cnt70 < ANOMALY_THRESHOLD:
                in_set = False

    # sauvegarde dernier set
    if last > 0 and current:
        anomalies_by_set[last] = sorted(current)

    return set_starts, anomaly_cells, meta, anomalies_by_set


def to_excel(df, set_starts, anomaly_cells, meta, anomalies_by_set, top10):
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
    ws2.append(["Set", "Date", "Nb anomalies"])
    total_anomalies = 0
    for m in meta:
        s = m["Set"]
        nb = len(anomalies_by_set.get(s, []))
        total_anomalies += nb
        ws2.append([s, m["Date"], nb])
    ws2.append([])
    ws2.append(["Total anomalies", total_anomalies])
    ws2.append([])
    ws2.append(["TOP 5 Emplacements","Occurrences"])
    for val, cnt in top10[:5]:
        ws2.append([val, cnt])

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

    all_anoms = [x for vals in anomalies_by_set.values() for x in vals]
    top10 = Counter(all_anoms).most_common(10)

    # Affichage récapitulatif
    recap = pd.DataFrame([
        {"Set": m["Set"], "Date": m["Date"], "Nb anomalies": len(anomalies_by_set.get(m["Set"], []))}
        for m in meta
    ]).set_index("Set")
    st.markdown("## Récapitulatif des sets")
    st.table(recap)
    # Total anomalies
    total = recap["Nb anomalies"].sum()
    st.markdown(f"**Total anomalies sur tous les sets : {total}**")

    # Affichage Top5
    st.markdown("## Top 5 des emplacements changés le plus souvent")
    df_top5 = pd.DataFrame(top10[:5], columns=["Emplacement","Occurrences"]).set_index("Emplacement")
    st.table(df_top5)

    # Bouton de téléchargement
    excel_bytes = to_excel(df_clean, set_starts, anomaly_cells, meta, anomalies_by_set, top10)
    st.download_button(
        "📥 Télécharger le rapport Excel",
        data=excel_bytes,
        file_name="analyse_creusets_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
