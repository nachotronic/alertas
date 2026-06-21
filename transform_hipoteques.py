"""
transform_hipoteques.py — versió neta
INE 13896 (constituïdes + capital) + 13902 (cancel·lades)
→ Excel Idescat: Full de càrrega + CAT + ESP
  - 12 períodes, ordre creixent
  - 3 columnes buides entre cada valor
"""

import urllib.request, json, datetime, sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

OUTPUT_FILE  = "hipoteques_idescat.xlsx"
N_PERIODES   = 12
N_FETCH      = 36

# Ordre final de columnes: (tipus_finca, mesura, capçalera_visible)
COLUMNES = [
    ("Total fincas",          "num",    "Hipoteques constituïdes (nombre)"),
    ("Total fincas urbanas",  "num",    "Finques urbanes"),
    ("Viviendas",             "num",    "Habitatges"),
    ("Solares",               "num",    "Solars"),
    ("Otros",                 "num",    "Altres"),
    ("Total fincas rústicas", "num",    "Finques rústiques"),
    ("Total fincas",          "imp",    "Capital prestat (milions d'euros)"),
    ("Total fincas urbanas",  "imp",    "Finques urbanes"),
    ("Viviendas",             "imp",    "Habitatges"),
    ("Solares",               "imp",    "Solars"),
    ("Otros",                 "imp",    "Altres"),
    ("Total fincas rústicas", "imp",    "Finques rústiques"),
    ("Total fincas",          "canc",   "Hipoteques cancel·lades (nombre)"),
    ("Total fincas urbanas",  "canc",   "Finques urbanes"),
    ("Viviendas",             "canc",   "Habitatges"),
    ("Solares",               "canc",   "Solars"),
    ("Fincas urbanas: otros", "canc",   "Altres"),
    ("Total fincas rústicas", "canc",   "Finques rústiques"),
]

def fetch(taula, n):
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    return json.loads(urllib.request.urlopen(req, timeout=30).read())

def periode(d):
    # Usar Anyo i FK_Periodo del INE (el timestamp pot donar mes incorrecte)
    return f"{d['Anyo']}{d['FK_Periodo']:02d}"

def mesura(nom):
    n = nom.lower()
    if "número de hipotecas" in n: return "num"
    if "importe de hipotecas" in n: return "imp"
    if "cancelada"           in n: return "canc"
    return None

def indexar(dades, geo):
    """{ (tipus, mesura) → { periode → valor } }"""
    idx = {}
    for s in dades:
        nom = s.get("Nombre", "")
        if geo not in nom: continue
        tipus = nom.split(".")[0].strip()
        m = mesura(nom)
        if not m: continue
        clau = (tipus, m)
        if clau not in idx: idx[clau] = {}
        for d in s.get("Data", []):
            if d.get("Valor") is not None:
                idx[clau][periode(d)] = d["Valor"]
    return idx

def construir(idx_const, idx_canc):
    # Períodes: unió de totes les claus, els 12 més recents en ordre creixent
    periodes = set()
    for d in list(idx_const.values()) + list(idx_canc.values()):
        periodes |= d.keys()
    periodes = sorted(sorted(periodes, reverse=True)[:N_PERIODES])

    # Caps
    caps = ["Periode"]
    for _, _, cap in COLUMNES:
        caps += [cap, "", "", ""]
    caps = caps[:-3]

    # Files
    files = []
    for p in periodes:
        fila = [p]
        for i, (tipus, m, _) in enumerate(COLUMNES):
            idx = idx_const if m in ("num", "imp") else idx_canc
            val = idx.get((tipus, m), {}).get(p)
            fila.append(val)
            if i < len(COLUMNES) - 1:
                fila += [None, None, None]
        files.append(fila)

    return caps, files

def escriure_full_carrega(ws):
    ws.title = "Full de càrrega"
    ws["A1"] = "Full de càrrega"
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=11)
    ws.column_dimensions["A"].width = 30

def escriure_dades(ws, caps, files):
    fill = PatternFill("solid", fgColor="1F3864")
    font = Font(color="FFFFFF", bold=True, size=9)
    ws.append(caps)
    for cell in ws[1]:
        cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 50
    for fila in files:
        ws.append(fila)
    ws.column_dimensions["A"].width = 10
    for col in ws.iter_cols(min_col=2, max_col=ws.max_column):
        ws.column_dimensions[col[0].column_letter].width = 12

def main():
    print("Descarregant...")
    try:
        d13896 = fetch("13896", N_FETCH)
        d13902 = fetch("13902", N_FETCH)
    except Exception as e:
        print(f"ERROR: {e}"); sys.exit(1)

    wb = Workbook()
    escriure_full_carrega(wb.active)

    for nom, geo in [("CAT", "Cataluña"), ("ESP", "Total Nacional")]:
        print(f"Processant {nom}...")
        idx_c = indexar(d13896, geo)
        idx_x = indexar(d13902, geo)
        print(f"  Claus const: {list(idx_c.keys())}")
        print(f"  Claus canc:  {list(idx_x.keys())}")
        caps, files = construir(idx_c, idx_x)
        if not files:
            print("  Sense dades."); continue
        ws = wb.create_sheet(title=nom)
        escriure_dades(ws, caps, files)
        print(f"  OK: {len(files)} períodes.")

    if len(wb.sheetnames) <= 1:
        print("ERROR: cap pestanya creada."); sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"Fitxer: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
