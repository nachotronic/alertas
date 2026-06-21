"""
transform_hipoteques.py
Descarrega les taules 13896 (constituïdes + capital) i 13902 (cancel·lades)
de l'API JSON del INE i genera un Excel amb dues pestanyes (CAT / ESP).
"""

import urllib.request
import json
import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
import sys

# ── Configuració ──────────────────────────────────────────────────────────────

TAULES = {
    "constituides": "13896",
    "cancelades":   "13902",
}

N_PERIODES = 12

GEO_FILTRES = {
    "CAT": "Cataluña",
    "ESP": "Total Nacional",
}

OUTPUT_FILE = "hipoteques_idescat.xlsx"

# ── Funcions ───────────────────────────────────────────────────────────────────

def fetch_taula(id_taula: str, n: int = 12) -> list:
    url = f"https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/{id_taula}?nult={n}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    data = urllib.request.urlopen(req, timeout=30).read()
    return json.loads(data)

def timestamp_a_periode(anyo: int, fecha_ms: int) -> str:
    """
    Converteix timestamp en ms → format YYYYMM.
    Usa UTC per evitar problemes de timezone al servidor de GitHub Actions.
    """
    dt = datetime.datetime.utcfromtimestamp(fecha_ms / 1000)
    return f"{dt.year}{str(dt.month).zfill(2)}"

def nom_net(nom: str, geo_text: str) -> str:
    """Elimina la part geogràfica i els sufixos estàndard del INE."""
    resultat = nom
    resultat = resultat.replace(f". {geo_text}.", ".")
    resultat = resultat.replace(f"{geo_text}. ", "")
    resultat = resultat.replace("Base nueva. Mensual.", "")
    resultat = resultat.replace("Base nueva.", "")
    resultat = resultat.strip(". ")
    return resultat

def parse_taula(dades: list, geo_text: str) -> pd.DataFrame:
    """
    Filtra per geo_text dins el camp Nombre,
    retorna DataFrame: index=periode (YYYYMM), columns=descripció variable.
    """
    registres = []
    for serie in dades:
        nom = serie.get("Nombre", "")
        if geo_text not in nom:
            continue

        variable = nom_net(nom, geo_text)

        for d in serie.get("Data", []):
            if d.get("Valor") is None:
                continue
            periode = timestamp_a_periode(d["Anyo"], d["Fecha"])
            registres.append({
                "periode":  periode,
                "variable": variable,
                "valor":    d["Valor"],
            })

    if not registres:
        print(f"  ⚠ Cap registre trobat per '{geo_text}'")
        return pd.DataFrame()

    df = pd.DataFrame(registres)
    df_pivot = df.pivot_table(
        index="periode",
        columns="variable",
        values="valor",
        aggfunc="first"
    )
    df_pivot = df_pivot.sort_index(ascending=False).head(N_PERIODES)
    return df_pivot

def afegir_columnes_buides(df: pd.DataFrame) -> pd.DataFrame:
    """Intercala 2 columnes buides entre cada columna de valors."""
    cols = []
    for i, col in enumerate(df.columns):
        cols.append(df[col])
        if i < len(df.columns) - 1:
            cols.append(pd.Series([None] * len(df), index=df.index, name=f"__b{i}a"))
            cols.append(pd.Series([None] * len(df), index=df.index, name=f"__b{i}b"))
    return pd.concat(cols, axis=1)

def escriure_pestanya(ws, df_final: pd.DataFrame):
    """Escriu el DataFrame a una pestanya d'Excel amb format."""
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(color="FFFFFF", bold=True, size=10)

    # Capçaleres (columnes buides sense nom)
    caps = ["Periode"] + [
        "" if str(c).startswith("__b") else str(c)
        for c in df_final.columns
    ]
    ws.append(caps)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[1].height = 45

    # Dades
    for periode, row in df_final.iterrows():
        ws.append([periode] + list(row))

    # Amplades
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

    print(f"Registres constituïdes: {len(dades_const)}")
    print(f"Registres cancel·lades: {len(dades_canc)}")

    # Diagnòstic: mostrar tots els Nombre disponibles per detectar el text geo correcte
    print("\nNoms de sèries disponibles (constituïdes):")
    for s in dades_const:
        print(f"  {s.get('Nombre','')}")

    wb = Workbook()
    wb.remove(wb.active)

    for nom_geo, geo_text in GEO_FILTRES.items():
        print(f"\nProcessant {nom_geo} ('{geo_text}')...")

        df_const = parse_taula(dades_const, geo_text)
        df_canc  = parse_taula(dades_canc,  geo_text)

        if df_const.empty and df_canc.empty:
            print(f"  Sense dades per a {nom_geo}, es salta.")
            continue

        df = pd.concat([df_const, df_canc], axis=1)
        df = df.loc[~df.index.duplicated(keep="first")]
        df = df.sort_index(ascending=False).head(N_PERIODES)
        df_final = afegir_columnes_buides(df)

        ws = wb.create_sheet(title=nom_geo)
        escriure_pestanya(ws, df_final)
        print(f"  ✅ {len(df_final)} períodes, {len(df.columns)} variables.")

    if not wb.sheetnames:
        print("\nERROR: cap pestanya creada.")
        print("Revisa els textos de GEO_FILTRES amb els noms de sèries mostrats a dalt.")
        sys.exit(1)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Fitxer generat: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
