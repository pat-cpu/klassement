# pdf_exporter/pdf.py
import os, shutil, tempfile, time
from typing import Optional

try:
    import pdfkit
except Exception:
    pdfkit = None


def wkhtmltopdf_available() -> bool:
    return shutil.which("wkhtmltopdf") is not None


def wkhtmltopdf_available_path() -> str:
    return shutil.which("wkhtmltopdf") or ""


def _derive_pdf_path_from_html(html_file: str) -> str:
    html_dir = os.path.dirname(html_file)
    base = os.path.splitext(os.path.basename(html_file))[0] + ".pdf"
    if os.sep + "html" in html_dir:
        pdf_dir = html_dir.replace(os.sep + "html", os.sep + "pdf")
    else:
        pdf_dir = html_dir
    os.makedirs(pdf_dir, exist_ok=True)
    return os.path.join(pdf_dir, base)


def maak_pdf(
    html_file: str,
    wkhtmltopdf_path: Optional[str] = None,
    orientation: str = "Portrait",
) -> Optional[str]:
    """
    Genereer een PDF uit een HTML-bestand.
    - orientation: "Portrait" of "Landscape"
    - Verbergt .no-print (navigatie) expliciet, ook als print-CSS genegeerd zou worden.
    """
    if pdfkit is None:
        print("[PDF] pdfkit ontbreekt. Installeer met: pip install pdfkit")
        return None
    if not os.path.exists(html_file):
        print(f"[PDF] HTML-bestand niet gevonden: {html_file}")
        return None
    if not wkhtmltopdf_available() and not wkhtmltopdf_path:
        print("[PDF] wkhtmltopdf niet gevonden. Zet in PATH of geef pad mee.")
        return None

    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path) if wkhtmltopdf_path else None
    pdf_file = _derive_pdf_path_from_html(html_file)

    # Tijdelijke CSS die .no-print altijd verbergt (als extra zekerheid)
    tmp_css = os.path.join(
        tempfile.gettempdir(),
        f"hide_nav_{int(time.time()*1000)}.css"
    )
    css_content = (
        ".no-print, nav.no-print { display: none !important; }\n"
        "@media print { .no-print, nav.no-print { display: none !important; } }\n"
    )
    with open(tmp_css, "w", encoding="utf-8") as f:
        f.write(css_content)

    options = {
        "enable-local-file-access": "",   # nodig om lokale CSS/links te laden
        "encoding": "UTF-8",
        "page-size": "A4",
        "orientation": orientation,       # "Portrait" of "Landscape"
        "print-media-type": "",           # respecteer @media print
        "margin-top": "8mm",
        "margin-right": "8mm",
        "margin-bottom": "8mm",
        "margin-left": "8mm",
        "user-style-sheet": tmp_css,      # forceer verbergen van .no-print
        # "zoom": "0.95",                 # optioneel: iets kleiner schalen
    }

    try:
        print(f"[PDF] {html_file} -> {pdf_file}")
        pdfkit.from_file(html_file, pdf_file, options=options, configuration=config)
        print(f"[PDF] Klaar: {pdf_file}")
        return pdf_file
    except Exception as e:
        print("[PDF] Fout bij genereren:")
        print(e)
        return None
    finally:
        try:
            os.remove(tmp_css)
        except Exception:
            pass
