# main.py
# Geschreven door Patrick - herwerkt op 13 november 2025
# Opgeschoond + batch-proof (CLI args) door Patrick’s IT Technicus

import os
import sys
import logging
import argparse
from glob import glob
from typing import Optional, List

from verwerking.data_loader import verwerk_klassement  # noqa: F401 (wordt elders gebruikt)
from html_generator.generator import maak_html, maak_controle_bestand
from pdf_exporter.pdf import maak_pdf, wkhtmltopdf_available
from verwerking.helpers import maanden  # verwacht: ['September','Oktober',...,'Augustus'] of jouw seizoenvolgorde

# === BASISPAD van het script ===
BASE = os.path.dirname(os.path.abspath(__file__))

# === Dynamische pad-functies ===
def data_dir(jaar: str) -> str:
    return os.path.join(BASE, "data", jaar)

def out_html_dir(jaar: str) -> str:
    return os.path.join(BASE, "output", jaar, "html")

def out_pdf_dir(jaar: str) -> str:
    return os.path.join(BASE, "output", jaar, "pdf")

# === Logginginstellingen ===
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def normalize_pdf_choice(value: Optional[str]) -> bool:
    """
    Converteer input naar True (PDF) / False (geen PDF).
    Default: True.
    """
    if value is None:
        return True
    v = value.strip().lower()
    if v in ("", "ja", "j", "y", "yes", "true", "1"):
        return True
    if v in ("nee", "n", "no", "false", "0"):
        return False
    # Onbekend -> default True, maar wel melden
    logging.warning(f"Onbekende PDF-keuze '{value}', ik neem 'Ja'.")
    return True


def detecteer_beschikbare_maanden(datadir: str) -> List[str]:
    """
    Geeft de maandnamen terug waarvoor er effectief een CSV bestaat in datadir,
    in de volgorde van 'maanden' (jouw helperlijst).
    """
    if not os.path.exists(datadir):
        logging.error(f"Map met data voor het jaar bestaat niet: {datadir}")
        sys.exit(1)

    bestaande = {os.path.splitext(os.path.basename(p))[0] for p in glob(os.path.join(datadir, "*.csv"))}
    gevonden = [m for m in maanden if m in bestaande]

    # Eventuele "vreemde" CSV namen ook loggen (handig bij typfouten)
    vreemd = sorted([m for m in bestaande if m not in maanden])
    if vreemd:
        logging.warning(f"Onbekende maandbestanden (niet in maanden-lijst): {', '.join(vreemd)}")

    return gevonden


def zorg_voor_outputmappen(htmldir: str, pdfdir: str):
    os.makedirs(htmldir, exist_ok=True)
    os.makedirs(pdfdir, exist_ok=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Genereer HTML/PDF klassement per seizoen/jaar-map.")
    parser.add_argument("--jaar", help="Seizoen/startjaar-map (bv. 2025). Als leeg: interactief vragen.")
    parser.add_argument("--pdf", help="PDF genereren? Ja/Nee (ook j/n/yes/no). Als leeg: interactief vragen.")
    return parser.parse_args()


def main():
    args = parse_args()

    # === Jaar bepalen: CLI > interactief > default
    jaar = (args.jaar or "").strip()
    if not jaar:
        jaar = input("Jaar (bv. 2025): ").strip() or "2025"

    # === PDF keuze bepalen: CLI > interactief > default
    if args.pdf is not None:
        pdf_gewenst = normalize_pdf_choice(args.pdf)
    else:
        pdf_keuze = input("Wil je ook PDF's genereren? (Ja/Nee): ").strip()
        pdf_gewenst = normalize_pdf_choice(pdf_keuze)

    # === Paden per jaar ===
    datadir = data_dir(jaar)
    htmldir = out_html_dir(jaar)
    pdfdir = out_pdf_dir(jaar)

    # === Check PDF-mogelijkheid ===
    if pdf_gewenst and not wkhtmltopdf_available():
        logging.warning("wkhtmltopdf is niet gevonden. PDF-generatie wordt overgeslagen.")
        pdf_gewenst = False

    # === CSV's detecteren (effectief aanwezige maanden) ===
    maand_namen = detecteer_beschikbare_maanden(datadir)
    logging.info(f"{len(maand_namen)} maand(en) gevonden in {datadir}")

    # === Outputmappen aanmaken ===
    zorg_voor_outputmappen(htmldir, pdfdir)

    # === Verwerk per maand die er echt is ===
    for maand in maand_namen:
        maand_nr = maanden.index(maand) + 1  # jouw maak_html verwacht maand_nr (1-based)
        logging.info(f"HTML genereren voor maand {maand}")
        maak_html(jaar, maand_nr)

        if pdf_gewenst:
            html_file = os.path.join(htmldir, f"{maand}.html")
            maak_pdf(html_file)

    # === Controle- en damesoverzicht ===
    maak_controle_bestand(jaar, len(maand_namen))

    if pdf_gewenst:
        dames_html = os.path.join(htmldir, "Dames.html")
        maak_pdf(dames_html)

    logging.info("✅ Verwerking afgerond.")


if __name__ == "__main__":
    main()
