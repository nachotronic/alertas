"""
transform_hipoteques.py
Taules INE 13896 (constituides + capital) i 13902 (cancelades)
→ Excel format Idescat: Full de càrrega + CAT + ESP
"""

import urllib.request
import json
import datetime
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import sys

TAULES = {"constituides": "13896", "cancelades": "13902"}
N_PERIODES = 12
GEO_FILTRES = {"CAT": "Cataluña", "ESP": "Total Nacional"}
OUTPUT_FILE = "hipoteques_idescat.xlsx"

# (paraules clau per cercar al nom raw, nom visible, taula font)
COLUMNES = [
    (["Total fincas", "Número"],        "Hipoteques constituïdes (nombre)",  "const"),
    (["fincas urbanas", "Número"],      "Finques urbanes",                   "const"),
    (["Viviendas", "Número"],           "Habitatges",                        "const"),
    (["Solares", "Número"],             "Solars",                            "const"),
    (["Otros", "Número"],               "Altres",                            "const"),
    (["rústicas", "Número"],            "Finques rústiques",                 "const"),
    (["Total fincas", "Importe"],       "Capital prestat (milions d'euros)", "const"),
    (["fincas urbanas", "Importe"],     "Finques urbanes",                   "const"),
    (["Viviendas", "Importe"],          "Habitatges",                        "const"),
    (["Solares", "Importe"],            "Solars",                            "const"),
    (["Otros", "Importe"],              "Altres",                            "const"),
    (["rústicas", "Importe"],           "Finques rústiques",                 "const"),
    (["Total fincas", "cancelada"],     "Hipoteques cancel·lades (nombre)",  "canc"),
    (["fincas urbanas", "cancelada"],   "Finques urbanes",                   "canc"),
    (["Viviendas", "cancelada"],        "Habitatges",                        "canc"),
    (["Solares", "cancelada"],          "Solars",                            "canc"),
    (["urbanas: otros", "cancelada"],   "Altres",                            "canc"),
    (["rústicas", "cancelada"],         "Finques rústiques",                 "canc"),
]

def fetch_taula(id_taula, n=12):
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.loads(urllib.request.urlopen(req, timeout=30).read())

def timestamp_a_periode(fecha_ms):
    dt = datetime.datetime.utcfromtimestamp(fecha_ms / 1000)
    return f"{dt.year}{str(dt.month).zfill(2)}"

def parse_taula(dades, geo_text):
    """Retorna dict { nom_raw_normalitzat → { periode → valor } }"""
    result = {}
    for serie in dades:
        nom_full = serie.get("Nombre", "")
        if geo_text not in nom_full:
            continue
        nom_raw = re.sub(r'\s+', ' ', nom_full).strip()
        for d in serie.get("Data", []):
            if d.get("Valor") is None:
                continue
            periode = timestamp_a_periode(d["Fecha"])
            if nom_raw not in result:
                result[nom_raw] = {}
            result[nom_raw][periode] = d["Valor"]
    return result

def cerca_clau(raw_dict, paraules):
    for clau in raw_dict:
        clau_lower = clau.lower()
        if all(p.lower() in clau_lower for p in paraules):
            return clau
    return None

def construir_files(raw_const, raw_canc):
    """Construeix llista de files en ordre correcte, sense pandas."""
    # Obtenir períodes
    periodes = set()
    for d in list(raw_const.values()) + list(raw_canc.values()):
        periodes.update(d.keys())
    periodes = sorted(periodes, reverse=True)[:N_PERIODES]

    # Pre-calcular claus per cada columna
    claus = []
    for paraules, nom_visible, font in COLUMNES:
        raw = raw_const if font == "const" else raw_canc
        clau = cerca_clau(raw, paraules)
        claus.append((clau, font, raw_const if font == "const" else raw_canc))
        print(f"  '{nom_visible}' ({font}) → '{clau}'")

    # Construir files
    files = []
    for periode in periodes:
        fila = [periode]
        for clau, font, raw in claus:
            val = raw.get(clau, {}).get(periode) if clau else None
            fila.append(val)
            # 3 columnes buides
            fila.extend([None, None, None])
        # Treure les 3 últimes buides (no cal després de l'última columna)
        fila = fila[:-3]
        files.append(fila)

    return periodes, files

def escriure_full_carrega(ws):
    ws.title = "Full de càrrega"
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(color="FFFFFF", bold=True, size=11)
    ws["A1"] = "Full de càrrega"
    ws["A1"].fill = fill
    ws["A1"].font = font
    ws.column_dimensions["A"].width = 30

def escriure_pestanya_dades(ws, files):
    fill_cap = PatternFill("solid", fgColor="1F3864")
    font_cap = Font(color="FFFFFF", bold=True, size=9)

    # Capçaleres
    caps = ["Periode"]
    for paraules, nom_visible, font in COLUMNES:
        caps.append(nom_visible)
        caps.extend(["", "", ""])
    caps = caps[:-3]  # treure les 3 últimes buides
    ws.append(caps)

    for cell in ws[1]:
        cell.fill = fill_cap
        cell.font = font_cap
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 50

    for fila in files:
        ws.append(fila)

    ws.column_dimensions["A"].width = 10
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        ws.column_dimensions[col[0].column_letter].width = 12

def main():
    print("Descarregant dades del INE...")
    try:
        dades_const = fetch_taula(TAULES["constituides"], N_PERIODES)
        dades_canc  = fetch_taula(TAULES["cancelades"],   N_PERIODES)
    except Exception as e:
        print(f"ERROR descàrrega: {e}")
        sys.exit(1)

    wb = Workbook()
    escriure_full_carrega(wb.active)

    for nom_geo, geo_text in GEO_FILTRES.items():
        print(f"\nProcessant {nom_geo}...")
        raw_const = parse_taula(dades_const, geo_text)
        raw_canc  = parse_taula(dades_canc,  geo_text)

        if not raw_const and not raw_canc:
            print(f"  Sense dades.")
            continue

        periodes, files = construir_files(raw_const, raw_canc)
        if not files:
            continue

        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya_dades(ws, files)
        print(f"  OK: {len(files)} periodes.")

    if len(wb.sheetnames) <= 1:
        print("ERROR: cap pestanya de dades creada.")
        sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"Fitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
