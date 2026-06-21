"""
transform_hipoteques.py
Taules INE 13896 (constituides + capital) i 13902 (cancelades)
→ Excel format Idescat: CAT i ESP, 12 periodes, 2 columnes buides entre valors.
"""

import urllib.request
import json
import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import sys

TAULES = {"constituides": "13896", "cancelades": "13902"}
N_PERIODES = 12
GEO_FILTRES = {"CAT": "Cataluña", "ESP": "Total Nacional"}
OUTPUT_FILE = "hipoteques_idescat.xlsx"

def fetch_taula(id_taula, n=12):
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.loads(urllib.request.urlopen(req, timeout=30).read())

def timestamp_a_periode(fecha_ms):
    dt = datetime.datetime.utcfromtimestamp(fecha_ms / 1000)
    return f"{dt.year}{str(dt.month).zfill(2)}"

def neteja_nom(nom, geo_text):
    for s in [f". {geo_text}.", f"{geo_text}. ", "Base nueva. Mensual.", "Base nueva.", " Mensual."]:
        nom = nom.replace(s, " ")
    return nom.strip(" .")

def parse_taula(dades, geo_text):
    registres = []
    for serie in dades:
        nom = serie.get("Nombre", "")
        if geo_text not in nom:
            continue
        variable = neteja_nom(nom, geo_text)
        for d in serie.get("Data", []):
            if d.get("Valor") is None:
                continue
            registres.append({
                "periode":  timestamp_a_periode(d["Fecha"]),
                "variable": variable,
                "valor":    d["Valor"],
            })
    if not registres:
        print(f"  Cap registre per '{geo_text}'")
        return pd.DataFrame()
    df = pd.DataFrame(registres)
    pivot = df.pivot_table(index="periode", columns="variable", values="valor", aggfunc="first")
    return pivot.sort_index(ascending=False).head(N_PERIODES)

def afegir_columnes_buides(df):
    cols = []
    for i, col in enumerate(df.columns):
        cols.append(df[col])
        if i < len(df.columns) - 1:
            cols.append(pd.Series([None]*len(df), index=df.index, name=f"__b{i}a"))
            cols.append(pd.Series([None]*len(df), index=df.index, name=f"__b{i}b"))
    return pd.concat(cols, axis=1)

def escriure_pestanya(ws, df_final):
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(color="FFFFFF", bold=True, size=10)
    caps = ["Periode"] + ["" if str(c).startswith("__b") else str(c) for c in df_final.columns]
    ws.append(caps)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 45
    for periode, row in df_final.iterrows():
        ws.append([periode] + list(row))
    ws.column_dimensions["A"].width = 10
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        ws.column_dimensions[col[0].column_letter].width = 14

def main():
    print("Descarregant dades del INE...")
    try:
        dades_const = fetch_taula(TAULES["constituides"], N_PERIODES)
        dades_canc  = fetch_taula(TAULES["cancelades"],   N_PERIODES)
    except Exception as e:
        print(f"ERROR descàrrega: {e}")
        sys.exit(1)

    print(f"Series constituides: {len(dades_const)} | Cancel·lades: {len(dades_canc)}")
    print("Noms de series disponibles:")
    for s in dades_const:
        print(f"  {s.get('Nombre','')}")

    wb = Workbook()
    wb.remove(wb.active)

    for nom_geo, geo_text in GEO_FILTRES.items():
        print(f"\n{nom_geo} ('{geo_text}')...")
        df_const = parse_taula(dades_const, geo_text)
        df_canc  = parse_taula(dades_canc,  geo_text)
        if df_const.empty and df_canc.empty:
            continue
        df = pd.concat([df_const, df_canc], axis=1)
        df = df.loc[~df.index.duplicated(keep="first")].sort_index(ascending=False).head(N_PERIODES)
        df_final = afegir_columnes_buides(df)
        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya(ws, df_final)
        print(f"  OK: {len(df_final)} periodes, {len(df.columns)} variables.")

    if not wb.sheetnames:
        print("ERROR: cap pestanya creada.")
        sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"\nFitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
