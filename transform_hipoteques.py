"""
transform_hipoteques.py
Descarrega les taules 13896 (constituïdes + capital) i 13902 (cancel·lades)
de l'API JSON del INE, filtra Catalunya i Espanya, i genera un Excel
amb dues pestanyes (CAT / ESP) en format Idescat:
- Una fila per període (format 202603)
- Valors bruts amb 2 columnes buides entre cada un
- Últims 12 períodes
"""

import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import json
import sys
from datetime import datetime

# ── Configuració ──────────────────────────────────────────────────────────────

TAULES = {
    "constituides": "13896",
    "cancelades":   "13902",
}

N_PERIODES = 12

# Codis de comunitat autònoma al INE
GEO = {
    "CAT": "09",   # Catalunya
    "ESP": "00",   # Total nacional
}

OUTPUT_FILE = "hipoteques_idescat.xlsx"

# ── Funcions ───────────────────────────────────────────────────────────────────

def fetch_taula(id_taula: str, n: int = 12) -> list:
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.json()

def periode_ine_a_idescat(periode_str: str) -> str:
    """
    Converteix '2026M03' o '2026M3' → '202603'
    """
    try:
        parts = periode_str.split("M")
        any_ = parts[0]
        mes  = parts[1].zfill(2)
        return f"{any_}{mes}"
    except Exception:
        return periode_str

def parse_taula(dades: list, geo_codi: str) -> pd.DataFrame:
    """
    Parseja el JSON del INE i retorna un DataFrame amb:
    index = periode (format Idescat)
    columns = (nom_variable, naturalesa_finca)
    """
    registres = []
    for item in dades:
        # Cada item té: Nombre (nom sèrie), Data, Valor, MetaData (llista de dims)
        meta = {m["Nombre"]: m["Codigo"] for m in item.get("MetaData", [])}

        # Filtre per comunitat autònoma
        ca = meta.get("Comunidades y ciudades autónomas", "")
        if ca != geo_codi:
            continue

        periode = periode_ine_a_idescat(item.get("Periodo", ""))
        nom     = item.get("Nombre", "")
        valor   = item.get("Valor")

        registres.append({
            "periode": periode,
            "nom":     nom,
            "valor":   valor,
        })

    if not registres:
        return pd.DataFrame()

    df = pd.DataFrame(registres)
    df_pivot = df.pivot_table(index="periode", columns="nom", values="valor", aggfunc="first")
    df_pivot = df_pivot.sort_index(ascending=False).head(N_PERIODES)
    return df_pivot

def afegir_columnes_buides(df: pd.DataFrame) -> pd.DataFrame:
    """
    Intercala 2 columnes buides entre cada columna de valors.
    """
    cols_noves = []
    for i, col in enumerate(df.columns):
        cols_noves.append(df[col])
        if i < len(df.columns) - 1:
            cols_noves.append(pd.Series([None] * len(df), index=df.index, name=f"_buit_{i}_1"))
            cols_noves.append(pd.Series([None] * len(df), index=df.index, name=f"_buit_{i}_2"))
    return pd.concat(cols_noves, axis=1)

def escriure_pestanya(ws, df_final: pd.DataFrame, nom_geo: str):
    """
    Escriu el DataFrame a una pestanya d'Excel amb format.
    """
    # Capçalera
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True, size=10)

    ws.append(["Periode"] + list(df_final.columns))
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Dades
    for periode, row in df_final.iterrows():
        ws.append([periode] + list(row))

    # Amplada columnes
    ws.column_dimensions["A"].width = 10
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        letter = col[0].column_letter
        ws.column_dimensions[letter].width = 14

# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    print("Descarregant dades del INE...")

    try:
        dades_const = fetch_taula(TAULES["constituides"], N_PERIODES)
        dades_canc  = fetch_taula(TAULES["cancelades"],   N_PERIODES)
    except Exception as e:
        print(f"ERROR en la descàrrega: {e}")
        sys.exit(1)

    wb = Workbook()
    wb.remove(wb.active)  # Elimina la pestanya per defecte

    for nom_geo, codi_geo in GEO.items():
        print(f"Processant {nom_geo}...")

        df_const = parse_taula(dades_const, codi_geo)
        df_canc  = parse_taula(dades_canc,  codi_geo)

        if df_const.empty and df_canc.empty:
            print(f"  Sense dades per a {nom_geo}, es salta.")
            continue

        # Combinar constituïdes + cancel·lades
        df = pd.concat([df_const, df_canc], axis=1)
        df = df.loc[~df.index.duplicated(keep="first")]
        df = df.sort_index(ascending=False).head(N_PERIODES)

        # Afegir columnes buides
        df_final = afegir_columnes_buides(df)

        # Escriure pestanya
        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya(ws, df_final, nom_geo)
        print(f"  {len(df_final)} períodes escrits.")

    wb.save(OUTPUT_FILE)
    print(f"\nFitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
