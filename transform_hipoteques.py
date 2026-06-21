"""
transform_hipoteques.py
Taules INE 13896 (constituides + capital) i 13902 (cancelades)
→ Excel format Idescat: Full de càrrega + CAT + ESP
  - Ordre de columnes fix
  - 3 columnes buides entre cada valor
  - 12 períodes
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

# Ordre exacte de columnes tal com el demana Idescat
# Aquests strings han de coincidir amb els noms nets de les sèries del INE
ORDRE_COLUMNES = [
    "Hipoteques constituïdes (nombre)",
    "Finques urbanes (constituïdes)",
    "Habitatges (constituïdes)",
    "Solars (constituïdes)",
    "Altres (constituïdes)",
    "Finques rústiques (constituïdes)",
    "Capital prestat (milions d'euros)",
    "Finques urbanes (capital)",
    "Habitatges (capital)",
    "Solars (capital)",
    "Altres (capital)",
    "Finques rústiques (capital)",
    "Hipoteques cancel·lades (nombre)",
    "Finques urbanes (cancel·lades)",
    "Habitatges (cancel·lades)",
    "Solars (cancel·lades)",
    "Altres (cancel·lades)",
    "Finques rústiques (cancel·lades)",
]

# Capçaleres visibles (sense el sufix entre parèntesis)
CAPS_VISIBLES = [
    "Hipoteques constituïdes (nombre)",
    "Finques urbanes",
    "Habitatges",
    "Solars",
    "Altres",
    "Finques rústiques",
    "Capital prestat (milions d'euros)",
    "Finques urbanes",
    "Habitatges",
    "Solars",
    "Altres",
    "Finques rústiques",
    "Hipoteques cancel·lades (nombre)",
    "Finques urbanes",
    "Habitatges",
    "Solars",
    "Altres",
    "Finques rústiques",
]

def fetch_taula(id_taula, n=12):
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.loads(urllib.request.urlopen(req, timeout=30).read())

def timestamp_a_periode(fecha_ms):
    dt = datetime.datetime.utcfromtimestamp(fecha_ms / 1000)
    return f"{dt.year}{str(dt.month).zfill(2)}"

def parse_taula(dades, geo_text):
    """
    Retorna dict: { periode → { nom_serie_net → valor } }
    Imprimeix els noms de sèries trobats per facilitar el mapeig.
    """
    registres = []
    for serie in dades:
        nom = serie.get("Nombre", "")
        if geo_text not in nom:
            continue
        print(f"  Serie trobada: {nom}")
        for d in serie.get("Data", []):
            if d.get("Valor") is None:
                continue
            registres.append({
                "periode":  timestamp_a_periode(d["Fecha"]),
                "nom_raw":  nom,
                "valor":    d["Valor"],
            })
    return registres

def construir_df(registres_const, registres_canc, geo_text):
    """
    Construeix el DataFrame amb les columnes en l'ordre correcte.
    Necessitem mapar els noms raw del INE als nostres noms interns.
    """
    # Tots els registres junts amb etiqueta
    tots = []
    for r in registres_const:
        tots.append(r)
    for r in registres_canc:
        tots.append(r)

    if not tots:
        return pd.DataFrame()

    df_raw = pd.DataFrame(tots)

    # Mapeig automàtic: agrupem per periode i nom_raw
    pivot = df_raw.pivot_table(index="periode", columns="nom_raw", values="valor", aggfunc="first")
    pivot = pivot.sort_index(ascending=False).head(N_PERIODES)

    print(f"\n  Columnes disponibles per {geo_text}:")
    for c in pivot.columns:
        print(f"    {c}")

    return pivot

def afegir_columnes_buides(df, n_buides=3):
    """Intercala n_buides columnes buides entre cada columna de valors."""
    cols = []
    for i, col in enumerate(df.columns):
        cols.append(df[col])
        if i < len(df.columns) - 1:
            for j in range(n_buides):
                cols.append(pd.Series([None]*len(df), index=df.index, name=f"__b{i}_{j}"))
    return pd.concat(cols, axis=1)

def escriure_full_carrega(ws):
    """Primera pestanya buida per a ús futur."""
    ws.title = "Full de càrrega"
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(color="FFFFFF", bold=True, size=11)
    ws["A1"] = "Full de càrrega"
    ws["A1"].fill = fill
    ws["A1"].font = font
    ws.column_dimensions["A"].width = 30

def escriure_pestanya_dades(ws, pivot, nom_geo):
    """Escriu les dades amb les columnes en l'ordre raw disponible + 3 buides."""
    fill_cap = PatternFill("solid", fgColor="1F3864")
    font_cap = Font(color="FFFFFF", bold=True, size=9)
    fill_alt = PatternFill("solid", fgColor="E8F0F7")

    df_final = afegir_columnes_buides(pivot, n_buides=3)

    # Capçaleres: nom raw sense la part geo i sufixos
    def neteja(nom):
        for s in [f". {nom_geo}.", f"{nom_geo}. ", "Base nueva. Mensual.", "Base nueva.", " Mensual.", ". "]:
            nom = nom.replace(s, " ")
        return nom.strip(" .")

    caps = ["Periode"] + [
        "" if str(c).startswith("__b") else neteja(str(c))
        for c in df_final.columns
    ]
    ws.append(caps)

    for cell in ws[1]:
        cell.fill = fill_cap
        cell.font = font_cap
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 50

    for i, (periode, row) in enumerate(df_final.iterrows()):
        ws.append([periode] + list(row))
        if i % 2 == 1:
            for cell in ws[i+2]:
                cell.fill = fill_alt

    ws.column_dimensions["A"].width = 10
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        ws.column_dimensions[col[0].column_letter].width = 13

def main():
    print("Descarregant dades del INE...")
    try:
        dades_const = fetch_taula(TAULES["constituides"], N_PERIODES)
        dades_canc  = fetch_taula(TAULES["cancelades"],   N_PERIODES)
    except Exception as e:
        print(f"ERROR descàrrega: {e}")
        sys.exit(1)

    print(f"Series constituides: {len(dades_const)} | Cancel·lades: {len(dades_canc)}")

    wb = Workbook()
    ws_carrega = wb.active
    escriure_full_carrega(ws_carrega)

    for nom_geo, geo_text in GEO_FILTRES.items():
        print(f"\nProcessant {nom_geo} ('{geo_text}')...")
        reg_const = parse_taula(dades_const, geo_text)
        reg_canc  = parse_taula(dades_canc,  geo_text)

        if not reg_const and not reg_canc:
            print(f"  Sense dades per a {nom_geo}")
            continue

        pivot = construir_df(reg_const, reg_canc, geo_text)
        if pivot.empty:
            continue

        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya_dades(ws, pivot, geo_text)
        print(f"  OK: {len(pivot)} periodes, {len(pivot.columns)} variables.")

    if len(wb.sheetnames) <= 1:
        print("ERROR: no s'han pogut crear les pestanyes de dades.")
        sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"\nFitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
