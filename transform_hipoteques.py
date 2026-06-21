"""
transform_hipoteques.py
Taules INE 13896 (constituides + capital) i 13902 (cancelades)
→ Excel format Idescat: Full de càrrega + CAT + ESP
Mapeig per nom exacte del INE.
"""

import urllib.request
import json
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import sys

TAULES = {"constituides": "13896", "cancelades": "13902"}
N_PERIODES = 12
OUTPUT_FILE = "hipoteques_idescat.xlsx"

# Ordre exacte desitjat: (prefix_tipus_finca, mesura, nom_visible)
# prefix_tipus_finca = text ABANS de ". {GEO}." al camp Nombre
ORDRE_CONST = [
    ("Total fincas",         "Número de hipotecas",  "Hipoteques constituïdes (nombre)"),
    ("Total fincas urbanas", "Número de hipotecas",  "Finques urbanes"),
    ("Viviendas",            "Número de hipotecas",  "Habitatges"),
    ("Solares",              "Número de hipotecas",  "Solars"),
    ("Otros",                "Número de hipotecas",  "Altres"),
    ("Total fincas rústicas","Número de hipotecas",  "Finques rústiques"),
    ("Total fincas",         "Importe de hipotecas", "Capital prestat (milions d'euros)"),
    ("Total fincas urbanas", "Importe de hipotecas", "Finques urbanes"),
    ("Viviendas",            "Importe de hipotecas", "Habitatges"),
    ("Solares",              "Importe de hipotecas", "Solars"),
    ("Otros",                "Importe de hipotecas", "Altres"),
    ("Total fincas rústicas","Importe de hipotecas", "Finques rústiques"),
]

ORDRE_CANC = [
    ("Total fincas",         "cancelada", "Hipoteques cancel·lades (nombre)"),
    ("Total fincas urbanas", "cancelada", "Finques urbanes"),
    ("Viviendas",            "cancelada", "Habitatges"),
    ("Solares",              "cancelada", "Solars"),
    ("Fincas urbanas: otros","cancelada", "Altres"),
    ("Total fincas rústicas","cancelada", "Finques rústiques"),
]

def fetch_taula(id_taula, n=12):
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.loads(urllib.request.urlopen(req, timeout=30).read())

def timestamp_a_periode(fecha_ms):
    dt = datetime.datetime.utcfromtimestamp(fecha_ms / 1000)
    return f"{dt.year}{str(dt.month).zfill(2)}"

def indexar_per_geo(dades, geo_text):
    """
    Retorna dict: { (prefix_tipus, mesura) → { periode → valor } }
    On prefix_tipus = text abans de '. {geo_text}.'
    i mesura = 'Número de hipotecas' / 'Importe de hipotecas' / 'cancelada'
    """
    index = {}
    for serie in dades:
        nom = serie.get("Nombre", "")
        # Format: "Tipus. GEO. Mesura. Base nueva. Mensual."
        # o per cancel·lades: "Tipus. GEO. Hipoteca cancelada"
        marcador = f". {geo_text}. "
        if marcador not in nom:
            continue
        parts = nom.split(marcador)
        if len(parts) < 2:
            continue
        tipus = parts[0].strip()
        resta = parts[1].strip().rstrip(". ")
        # Normalitzar mesura
        if "Número de hipotecas" in resta:
            mesura = "Número de hipotecas"
        elif "Importe de hipotecas" in resta:
            mesura = "Importe de hipotecas"
        elif "cancelada" in resta.lower():
            mesura = "cancelada"
        else:
            mesura = resta

        clau = (tipus, mesura)
        if clau not in index:
            index[clau] = {}
        for d in serie.get("Data", []):
            if d.get("Valor") is None:
                continue
            periode = timestamp_a_periode(d["Fecha"])
            index[clau][periode] = d["Valor"]
    return index

def construir_files(index_const, index_canc, geo_text):
    # Obtenir tots els períodes
    periodes = set()
    for d in list(index_const.values()) + list(index_canc.values()):
        periodes.update(d.keys())
    periodes = sorted(periodes, reverse=True)[:N_PERIODES]

    # Definir columnes en ordre
    columnes = []
    for tipus, mesura, nom_visible in ORDRE_CONST:
        clau = (tipus, mesura)
        data = index_const.get(clau, {})
        columnes.append((nom_visible, data))
        if not data:
            print(f"  AVÍS: no trobat ({tipus} / {mesura})")

    for tipus, mesura, nom_visible in ORDRE_CANC:
        clau = (tipus, mesura)
        data = index_canc.get(clau, {})
        columnes.append((nom_visible, data))
        if not data:
            print(f"  AVÍS: no trobat ({tipus} / {mesura})")

    # Construir files amb 3 columnes buides entre valors
    caps = ["Periode"]
    for nom_visible, _ in columnes:
        caps.append(nom_visible)
        caps.extend(["", "", ""])
    caps = caps[:-3]

    files = []
    for periode in periodes:
        fila = [periode]
        for i, (nom_visible, data) in enumerate(columnes):
            fila.append(data.get(periode))
            if i < len(columnes) - 1:
                fila.extend([None, None, None])
        files.append(fila)

    return caps, files

def escriure_full_carrega(ws):
    ws.title = "Full de càrrega"
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(color="FFFFFF", bold=True, size=11)
    ws["A1"] = "Full de càrrega"
    ws["A1"].fill = fill
    ws["A1"].font = font
    ws.column_dimensions["A"].width = 30

def escriure_pestanya(ws, caps, files):
    fill_cap = PatternFill("solid", fgColor="1F3864")
    font_cap = Font(color="FFFFFF", bold=True, size=9)

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

    for nom_geo, geo_text in [("CAT", "Cataluña"), ("ESP", "Total Nacional")]:
        print(f"\nProcessant {nom_geo}...")
        index_const = indexar_per_geo(dades_const, geo_text)
        index_canc  = indexar_per_geo(dades_canc,  geo_text)

        print(f"  Claus constituides: {list(index_const.keys())}")
        print(f"  Claus cancel·lades: {list(index_canc.keys())}")

        caps, files = construir_files(index_const, index_canc, geo_text)
        if not files:
            print(f"  Sense dades.")
            continue

        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya(ws, caps, files)
        print(f"  OK: {len(files)} periodes, {len(caps)} columnes (incl. buides).")

    if len(wb.sheetnames) <= 1:
        print("ERROR: cap pestanya de dades creada.")
        sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"\nFitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
