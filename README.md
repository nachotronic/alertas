# alertes

Monitor automàtic de fonts estadístiques externes per a Idescat.

Cada workflow comprova el calendari oficial de la font, detecta si hi ha publicació nova, descarrega les dades, les transforma al format Idescat i envia el fitxer per email.

## Workflows actius

| Workflow | Font | Freqüència | Taules |
|---|---|---|---|
| `hipoteques.yml` | INE | Cada dia 9:10h · publica ~dia 22 | t=13896, t=13902 |

## Format de sortida

- Excel `.xlsx` amb pestanyes **CAT** i **ESP**
- Una fila per període en format `YYYYMM` (ex: `202603`)
- Valors bruts amb **2 columnes buides** entre cada un (per a variacions calculades pel sistema Idescat)
- Últims **12 períodes**

## Secrets necessaris

| Secret | Descripció |
|---|---|
| `GMAIL_USER` | Compte Gmail per a l'enviament |
| `GMAIL_APP_PASSWORD` | App Password de 16 caràcters |
| `NOTIFY_EMAILS` | Destinataris separats per comes |

## Afegir noves fonts

Cada nova font és un workflow independent (`.github/workflows/nom_font.yml`) + script Python de transformació (`transform_nom_font.py`).

## Fonts pendents

- [ ] Fomento — Licitació oficial en construcció (XLS estàtics, detecció per hash)
- [ ] Aena — Passatgers (font manual, upload al portal)
- [ ] Portal de fonts (GitHub Pages amb estat de cada font)
