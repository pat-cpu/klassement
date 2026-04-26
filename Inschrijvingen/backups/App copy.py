# --------------------- USB -------------------
from __future__ import annotations

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from pathlib import Path
import xlwings as xw
import logging
from threading import Lock
import time
import os
import sys
import atexit

# COM (Windows)
if sys.platform.startswith("win"):
    import pythoncom  # type: ignore


# -------------------------------------------------
# Flask
# -------------------------------------------------
app = Flask(__name__)
app.secret_key = "geheim123"  # in productie: via env var
logging.basicConfig(level=logging.INFO)

# -------------------------------------------------
# Excel config
# -------------------------------------------------
filename = Path(__file__).parent.parent / "Whistloting.xlsm"
tabblad = "AanwezigReserve"

# 1 gebruiker: globals is ok
scanned_data: list[dict] = []
total_scans = 0

# Excel/COM wordt gelukkiger met rust
excel_lock = Lock()

# Banner status cache
excel_ok = False
excel_last_check = 0.0
excel_last_error = ""
EXCEL_CHECK_TTL_SEC = 5  # cache (maar we gaan nauwelijks nog checken)

# Auto-open: maar 1 keer proberen (anders opent hij Excel telkens opnieuw)
excel_open_attempted = False

# ✅ Warm-up: maar 1 keer checken/openen bij eerste scan
excel_warmed_up = False

# ✅ Persistent Excel instance (belangrijk!)
xl_app: xw.App | None = None
xl_book: xw.Book | None = None


# -------------------------------------------------
# Helpers
# -------------------------------------------------
def normalize_barcode(x) -> str:
    return str(x).strip().split(".")[0]


def try_open_excel_file_once() -> None:
    """Probeert Excelbestand te openen (Windows). Doet dit max 1x per app-run."""
    global excel_open_attempted
    if excel_open_attempted:
        return
    excel_open_attempted = True

    if sys.platform.startswith("win"):
        try:
            os.startfile(str(filename))  # opent in Excel (zichtbaar)
            logging.info("Excelbestand automatisch geopend.")
        except Exception as e:
            logging.warning(f"Kon Excel niet automatisch openen: {e}")
    else:
        logging.info("Auto-open is enkel voorzien voor Windows (os.startfile).")


def excel_disconnect() -> None:
    """Hard reset van Excel/Workbook (ook bij COM-zombies)."""
    global xl_app, xl_book

    try:
        if xl_book is not None:
            xl_book.close()
    except Exception:
        pass
    xl_book = None

    try:
        if xl_app is not None:
            xl_app.quit()
    except Exception:
        pass
    xl_app = None


def excel_connect() -> None:
    """Zorg dat we 1 Excel instance + 1 open workbook hebben."""
    global xl_app, xl_book

    if not filename.exists():
        raise FileNotFoundError(f"Bestand niet gevonden: {filename}")

    # COM init (veilig, ook al is threaded=False)
    if sys.platform.startswith("win"):
        pythoncom.CoInitialize()  # type: ignore

    if xl_app is not None and xl_book is not None:
        return

    xl_app = xw.App(visible=False, add_book=False)
    xl_app.display_alerts = False
    xl_app.screen_updating = False

    xl_book = xl_app.books.open(
        str(filename),
        update_links=False,
        read_only=False
    )


def excel_ensure(retries: int = 2) -> None:
    """Connect met retry + reset bij COM-fouten."""
    last: Exception | None = None
    for attempt in range(retries + 1):
        try:
            excel_connect()
            return
        except Exception as e:
            last = e
            excel_disconnect()
            time.sleep(0.5 * (attempt + 1))
    # als het na retries nog faalt: exception doorgeven
    assert last is not None
    raise last


atexit.register(excel_disconnect)

def check_excel_ready() -> tuple[bool, str]:
    """Startcheck: kan Excel open + sheet ok + F bevat enkel A/leeg/getal.
       Waarschuwt als er al gescand is (getallen in F).
    """
    if not filename.exists():
        return False, f"Bestand niet gevonden: {filename}"

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            try:
                sh = xl_book.sheets[tabblad]
            except Exception:
                return False, f"Tabblad '{tabblad}' niet gevonden in {filename.name}"

            barcodes = sh.range("A3:A100").options(ndim=1).value or []
            fvals    = sh.range("F3:F100").options(ndim=1).value or []

            # minstens 1 barcode aanwezig
            if not any(bc is not None and str(bc).strip() != "" for bc in barcodes):
                return False, f"Geen barcodes gevonden in kolom A (A3:A100) op tabblad '{tabblad}'."

            bad_rows = []
            already_scanned_rows = 0

            for idx, bc in enumerate(barcodes):
                if bc is None or str(bc).strip() == "":
                    continue

                v = fvals[idx] if idx < len(fvals) else None
                if v is None or str(v).strip() == "":
                    continue

                s = str(v).strip()
                if s.upper() == "A":
                    continue

                # getal = ok (al gescand / ronden)
                try:
                    n = int(float(s))
                    if n < 1:
                        bad_rows.append(3 + idx)
                    else:
                        already_scanned_rows += 1
                except Exception:
                    bad_rows.append(3 + idx)

            if bad_rows:
                toon = ", ".join(map(str, bad_rows[:10]))
                extra = " …" if len(bad_rows) > 10 else ""
                return False, (
                    f"Niet klaar: tabblad '{tabblad}' kolom F bevat ongeldige waarden. "
                    f"Toegelaten: 'A' of een positief getal. Probleem in rijen: {toon}{extra}."
                )

            if already_scanned_rows > 0:
                return True, (
                    f"Excel ok. Opgelet: er staan al {already_scanned_rows} ingevulde aantallen in kolom F "
                    f"(er is waarschijnlijk al eens gescand)."
                )

        return True, "Excel is klaar om te scannen."

    except Exception as e:
        logging.warning(f"Excel niet klaar: {e}")
        excel_disconnect()
        return False, (
            "Excel is niet klaar. Open 'Whistloting.xlsm' één keer in Excel, "
            "klik indien gevraagd op 'Inhoud inschakelen', sluit het bestand en probeer opnieuw."
        )










def get_excel_status(force: bool = False) -> tuple[bool, str]:
    """Cached Excel status."""
    global excel_ok, excel_last_check, excel_last_error

    now = time.time()
    if (not force) and (now - excel_last_check) < EXCEL_CHECK_TTL_SEC:
        return excel_ok, (excel_last_error if not excel_ok else "Excel is klaar.")

    ok, msg = check_excel_ready()
    excel_ok = ok
    excel_last_check = now
    excel_last_error = "" if ok else msg
    return ok, msg


# -------------------------------------------------
# Excel actions
# -------------------------------------------------
def update_quantity(barcode: str, aantal: int) -> tuple[bool, str, str | None]:
    """Zet aantal voor barcode. Return: (ok, msg, naam)."""
    global total_scans

    barcode = normalize_barcode(barcode)

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            sh = xl_book.sheets[tabblad]

            barcodes = sh.range("A3:A100").options(ndim=1).value or []
            aantallen = sh.range("F3:F100").options(ndim=1).value or []

            for idx, cell_value in enumerate(barcodes):
                if cell_value is None:
                    continue

                if normalize_barcode(cell_value) == barcode:
                    row = 3 + idx
                    huidig = aantallen[idx] if idx < len(aantallen) else None

                    if isinstance(huidig, (int, float)) and huidig > 0:
                        return False, f"Barcode {barcode} is al ingevoerd met aantal {int(huidig)}.", None

                    naam = sh.range(f"D{row}").value
                    sh.range(f"F{row}").value = int(aantal)

                    xl_book.save()

                    scanned_data.append({"barcode": barcode, "naam": naam, "aantal": int(aantal)})
                    total_scans += 1
                    return True, f"Barcode {barcode} succesvol bijgewerkt.", str(naam) if naam is not None else None

            return False, f"Barcode {barcode} werd niet gevonden in de Excel-lijst.", None

    except Exception as e:
        logging.exception("Fout bij update_quantity")
        # Excel kan in zombie-state zitten -> resetten
        excel_disconnect()
        return False, f"Fout bij het bijwerken van de barcode: {e}", None


def remove_quantity(barcode: str, aantal: int) -> bool:
    """Zet F terug op 'A' als barcode+aantal matchen."""
    barcode = normalize_barcode(barcode)

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            sh = xl_book.sheets[tabblad]

            barcodes = sh.range("A3:A100").options(ndim=1).value or []
            aantallen = sh.range("F3:F100").options(ndim=1).value or []

            for idx, value in enumerate(barcodes):
                if value is None:
                    continue

                if normalize_barcode(value) == barcode:
                    huidig = aantallen[idx] if idx < len(aantallen) else None
                    try:
                        if int(huidig) == int(aantal):
                            rij = 3 + idx
                            sh.range(f"F{rij}").value = "A"
                            xl_book.save()
                            return True
                    except (TypeError, ValueError):
                        pass

            return False

    except Exception:
        logging.exception("Fout bij remove_quantity")
        excel_disconnect()
        return False


# -------------------------------------------------
# Routes
# -------------------------------------------------
@app.route("/excel_status")
def excel_status():
    """Snel status voor banner: na warm-up geen extra Excel open/close."""
    global excel_warmed_up

    force = request.args.get("force") == "1"

    # Na warm-up: altijd OK antwoorden (geen extra Excel-check)
    if excel_warmed_up and not force:
        return jsonify({
            "ok": True,
            "msg": "Excel is klaar (warm-up gedaan).",
            "file": str(filename),
            "sheet": tabblad
        })

    # Voor warm-up: niet automatisch Excel openen, enkel info tonen
    if not excel_warmed_up and not force:
        return jsonify({
            "ok": False,
            "msg": "Nog niet getest — scan om te starten (of klik 'Test Excel').",
            "file": str(filename),
            "sheet": tabblad
        })

    # Force-test (op knop)
    ok, msg = get_excel_status(force=True)
    if ok:
        excel_warmed_up = True

    return jsonify({
        "ok": ok,
        "msg": msg,
        "file": str(filename),
        "sheet": tabblad
    })


@app.route("/", methods=["GET", "POST"])
def index():
    global excel_warmed_up

    if request.method == "POST":
        barcode = request.form.get("barcode", "").strip()
        aantal_str = request.form.get("aantal", "").strip()

        if not barcode:
            flash("Barcode is vereist.", "error")
            return redirect(url_for("index"))

        if not aantal_str:
            flash("Aantal is vereist.", "error")
            return redirect(url_for("index"))

        # ✅ Alleen bij eerste scan: warm-up check
        if not excel_warmed_up:
            ok_now, msg_now = get_excel_status(force=True)
            if not ok_now:
                try_open_excel_file_once()
                flash(msg_now, "warning")
                flash("Excel werd geopend. Sluit het bestand en scan opnieuw.", "warning")
                return redirect(url_for("index"))
            excel_warmed_up = True

        # Vanaf hier: geen extra checks, gewoon schrijven
        try:
            aantal = int(aantal_str)
        except ValueError:
            flash("Voer een geldig getal in voor het aantal.", "error")
            return redirect(url_for("index"))

        ok_upd, msg_upd, _ = update_quantity(barcode, aantal)
        flash(msg_upd, "success" if ok_upd else "error")

        # Als Excel toch ineens faalt: warm-up resetten zodat je opnieuw kan starten
        if (not ok_upd) and msg_upd.startswith("Fout bij het bijwerken"):
            excel_warmed_up = False

        return redirect(url_for("index"))

    # -----------------------------
    # GET: bereken tellers voor UI
    # -----------------------------
    total_quantity = 0
    for item in scanned_data:
        try:
            total_quantity += int(item.get("aantal", 0) or 0)
        except (TypeError, ValueError):
            pass

    count_2 = sum(1 for item in scanned_data if int(item.get("aantal", 0) or 0) == 2)
    count_3 = sum(1 for item in scanned_data if int(item.get("aantal", 0) or 0) == 3)
    count_4 = sum(1 for item in scanned_data if int(item.get("aantal", 0) or 0) == 4)

    return render_template(
        "index.html",
        scanned_data=scanned_data,
        total_scans=total_scans,
        total_quantity=total_quantity,
        count_2=count_2,
        count_3=count_3,
        count_4=count_4,
    )


@app.route("/remove_entry", methods=["POST"])
def remove_entry():
    global scanned_data, total_scans, excel_warmed_up

    barcode = request.form.get("barcode", "").strip()
    aantal = request.form.get("aantal", "").strip()

    try:
        aantal_int = int(aantal)
    except ValueError:
        flash("Ongeldig aantal ingevoerd.", "error")
        return redirect(url_for("index"))

    # ✅ Geen check per remove. Als Excel faalt: waarschuwing en warm-up reset
    for i, item in enumerate(scanned_data):
        try:
            if item["barcode"] == barcode and int(item["aantal"]) == aantal_int:
                if remove_quantity(barcode, aantal_int):
                    scanned_data.pop(i)
                    total_scans -= 1
                    flash(f"Barcode {barcode} succesvol verwijderd.", "success")
                else:
                    flash("Verwijderen in Excel mislukte.", "error")
                    excel_warmed_up = False
                break
        except (ValueError, TypeError, KeyError):
            continue
    else:
        flash(f"Barcode {barcode} niet gevonden of aantal komt niet overeen.", "error")

    return redirect(url_for("index"))


# -------------------------------------------------
# Start
# -------------------------------------------------
if __name__ == "__main__":
    # Tijdens scannen: debug uit. (reloader staat al uit)
    app.run(debug=False, threaded=False, use_reloader=False)
