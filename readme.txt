
Snel: schrijf README.md via PowerShell

Voer dit uit in C:\Users\patri\Documenten\Whiskies:

@'
# Whiskies — Loting, Scannen & Klassement

![CI](https://github.com/patrickgeys740-svg/Whiskies/actions/workflows/ci.yml/badge.svg)

Toolkit om whist-avonden vlot te organiseren: loting & score-invoer in Excel, CSV-export naar Python, en HTML/PDF-klassementen.

## Inhoud
- [Overzicht](#overzicht)
- [Projectstructuur](#projectstructuur)
- [Vereisten](#vereisten)
- [Installatie](#installatie)
- [Snel starten (Quick start)](#snel-starten-quick-start)
- [Dagelijkse flow](#dagelijkse-flow)
- [Klassement genereren](#klassement-genereren)
- [Troubleshooting](#troubleshooting)
- [Voor ontwikkelaars](#voor-ontwikkelaars)
- [Roadmap](#roadmap)

---

## Overzicht

- **Excel (`Whistloting.xlsm`)**  
  Loting per ronde, invoer van scores (10/6/3/1), consolidatie in tabblad `samen`.  
  Cel **`Loting!H8`** bevat de datum (vaak `=VANDAAG()`), waarvan we de **maandnaam** gebruiken.

- **Klassement (Python, map `klassement/`)**  
  - `scores_ophalen.py`: leest `samen` uit Excel en schrijft `data/<jaar>/<Maand>.csv`.  
  - `main.py`: bouwt HTML per maand, `Dames.html` en `Controle.html` (+ optioneel PDF via wkhtmltopdf).  
  - `run_all.bat`: one-shot script dat CSV’s exporteert én HTML/PDF genereert.

- **Scannen (map `Inschrijvingen/`, optioneel)**  
  Scripts om lidkaarten te scannen en Excel automatisch bij te werken.

Alles werkt locatie-onafhankelijk (Documenten, OneDrive, Dropbox…) zolang je vanuit deze projectmap werkt.

---

## Projectstructuur



Whiskies/
├─ Inschrijvingen/ # scanprogramma (optioneel)
├─ klassement/
│ ├─ main.py
│ ├─ run_all.bat
│ ├─ scores_ophalen.py
│ ├─ html_generator/
│ ├─ verwerking/
│ ├─ pdf_exporter/
│ ├─ hulp/extra.css
│ ├─ data/ # gegenereerde CSV’s (niet in Git)
│ └─ output/ # gegenereerde HTML/PDF (niet in Git)
└─ Whistloting.xlsm # Excel workbook (mag ook in klassement/ staan)


> De scripts zoeken automatisch een `whis*.xlsm` in **`klassement/`** of **één map hoger** (de repo-root).  
> Desnoods kun je een expliciet pad meegeven met `--xlsm`.

---

## Vereisten

- **Windows**
- **Microsoft Excel** (voor `.xlsm`)
- **Python 3.11+** → `python --version`
- Python-pakketten: `pandas`, `openpyxl`, `pdfkit`  
  *(optioneel voor testen/webapp: `pytest`, `flask`, `xlwings`)*
- **wkhtmltopdf** voor PDF’s  
  Test: `where wkhtmltopdf` en `"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" --version`

Installeer pakketten:
```powershell
pip install pandas openpyxl pdfkit
# optioneel
pip install pytest flask xlwings

Snel starten (Quick start)
# Ga naar de klassement-map
cd .\klassement

# Alles in één: exporteer CSV + bouw HTML (en PDF's indien gevraagd)
.\run_all.bat
# of zonder vragen:
.\run_all.bat 2025 Ja


Resultaat:

HTML: klassement\output\<jaar>\html\

PDF: klassement\output\<jaar>\pdf\

Dagelijkse flow

Excel voorbereiden

Open Whistloting.xlsm.

Voer loting/scores in (scores 10–6–3–1).

Volledige herberekening: Ctrl+Alt+Shift+F9.

Bewaar (zodat Loting!H8 een actuele opgeslagen waarde heeft).

Sluit Excel (vermijdt lock-problemen).

CSV exporteren (per maand)

In klassement/:

python .\scores_ophalen.py --jaar 2025


Dit leest de maand uit Loting!H8 en schrijft data\2025\<Maand>.csv.

Het competitiejaar is het startjaar (bijv. 2025 voor Sep 2025 t/m Jun 2026).

Klassement genereren
Alles-in-één
cd .\klassement
.\run_all.bat
# of .\run_all.bat 2025 Ja

Handmatig

CSV’s maken met scores_ophalen.py.

HTML/PDF bouwen:

python .\main.py


Wat wordt gemaakt?

Per maand: September.html, Oktober.html, … (met navigatie bovenaan; nav wordt verborgen in PDF)

Dames.html (altijd)

Controle.html (vanaf maand 2, toont vorige vs. huidige maand)

Snel openen:

start "" .\output\2025\html\
start "" .\output\2025\pdf\

Troubleshooting

PermissionError / ~$Whistloting.xlsm
Dat is een Excel-lockfile. Sluit Excel en run opnieuw.

Lege/foute maand uit H8
Excel niet herberekend/opgeslagen. Doe Ctrl+Alt+Shift+F9, Opslaan, sluit, probeer opnieuw.

wkhtmltopdf niet gevonden
Installeer, zet in PATH, of geef het pad door in pdf_exporter/pdf.py.
PDF’s zijn A4 staand, navigatie verborgen via @media print.

Geen Controle.html
Die verschijnt pas vanaf maand 2. Dames.html komt altijd.

Voor ontwikkelaars
Tests (indien aanwezig)
cd .\klassement
pytest -q

GitHub Actions (CI)

Eenvoudige workflow in .github/workflows/ci.yml (Python 3.11, deps, pytest).
Badge staat bovenaan deze README.

Git-tips

klassement/data/ en klassement/output/ staan niet in Git (worden bij run aangemaakt).

Grote bestanden (zoals .xlsm) staan via Git LFS.

Roadmap

✅ Excel voor loting en puntentelling

✅ CSV-export + HTML (maanden, Dames, Controle)

✅ PDF’s (A4 staand; nav verborgen)

✅ Basis CI (tests)

⏳ Optioneel: build-artefacts/Pages publicatie

⏳ Optioneel: één UI om alles te starten (Flask/desktop)

'@ | Set-Content -Encoding UTF8 README.md


Daarna (optioneel) meteen committen en pushen:
```powershell
git add README.md
git commit -m "Add README handleiding"
git push


klassement/
├── main.py
├── data/
│   ├── 2025/
│   │   ├── September.csv
│   │   └── ...
│   └── 2026/
│       ├── September.csv
│       └── ...
├── output/
│   └── <jaar>/
│       ├── html/
│       └── pdf/
├── verwerking/
│   ├── __init__.py
│   └── data_loader.py         # score(), verwerk(), verwerk_klassement(), tel_punten, maanden
├── html_generator/
│   ├── __init__.py
│   └── generator.py           # maak_html(), maak_controle_bestand() + helpers
├── pdf_exporter/
│   ├── __init__.py
│   └── pdf.py                 # maak_pdf(), wkhtmltopdf_available()
└── hulp/
    └── extra.css





./run_all.sh
