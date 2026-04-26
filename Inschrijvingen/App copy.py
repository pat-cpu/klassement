# App.py — THE WHISKIES inschrijvingen scanner (Flask + xlwings)
# Kolommen Excel:
#   A = Barcode
#   D = Naam (volledige naam)
#   F = "A" (afwezig) of getal (aantal ronden)

from __future__ import annotations

from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from pathlib import Path
import xlwings as xw
import logging
from threading import Lock
import time
import os
import sys
import json
import atexit
from datetime import datetime
import signal
import shutil

# COM (Windows)
if sys.platform.startswith("win"):
    import pythoncom  # type: ignore


# -------------------------------------------------
# Flask
# -------------------------------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "geheim123")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s"
)

# -------------------------------------------------
# Excel config
# -------------------------------------------------
filename = Path(__file__).parent / "Whistloting.xlsm"
tabblad = "AanwezigReserve"

# Scan state
scanned_data: list[dict] = []
total_scans = 0

excel_lock = Lock()

# Status cache
excel_ok = False
excel_last_check = 0.0
excel_last_error = ""
EXCEL_CHECK_TTL_SEC = 5

# Excel auto-open max 1x
excel_open_attempted = False

# Warm-up flag
excel_warmed_up = False

# Persistent Excel
xl_app: xw.App | None = None
xl_book: xw.Book | None = None

# Save throttling
SAVE_EVERY_N_SCANS = 5
last_save_time = 0.0
MIN_SAVE_INTERVAL_SEC = 1.5

# Scan log
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
SCAN_LOG_PATH = OUTPUT_DIR / "scans_log.jsonl"

# Backup config
BACKUP_DIR = Path(__file__).parent / "backups"
BACKUP_DIR.mkdir(parents=True, exist_ok=True)
BACKUP_EVERY_N_SCANS = 10
KEEP_LAST_BACKUPS = 20


# -------------------------------------------------
# Helpers
# -------------------------------------------------
def normalize_barcode(x) -> str:
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if "." in s:
        left, right = s.split(".", 1)
        if right.isdigit():
            s = left
    return s


def parse_positive_int(v) -> int | None:
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s.upper() == "A":
        return None
    try:
        n = int(float(s))
        return n if n > 0 else None
    except Exception:
        return None


def append_scan_log(entry: dict) -> None:
    try:
        entry = dict(entry)
        entry.setdefault("ts", datetime.now().isoformat(timespec="seconds"))
        line = json.dumps(entry, ensure_ascii=False)
        with SCAN_LOG_PATH.open("a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        logging.warning(f"Kon scanlog niet schrijven: {e}")


def cleanup_old_backups(keep_last: int = KEEP_LAST_BACKUPS) -> None:
    try:
        files = sorted(
            BACKUP_DIR.glob(f"{filename.stem}_*{filename.suffix}"),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        for old_file in files[keep_last:]:
            try:
                old_file.unlink()
            except Exception:
                pass
    except Exception as e:
        logging.warning(f"Opruimen backups mislukt: {e}")


def make_excel_backup(reason: str = "manual") -> Path | None:
    """Maakt een timestamp-backup van het Excelbestand."""
    try:
        if not filename.exists():
            logging.warning("Geen backup gemaakt: Excelbestand niet gevonden.")
            return None

        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        backup_name = f"{filename.stem}_{reason}_{ts}{filename.suffix}"
        backup_path = BACKUP_DIR / backup_name

        try:
            if xl_book is not None:
                xl_book.save()
        except Exception:
            pass

        shutil.copy2(filename, backup_path)
        cleanup_old_backups()
        logging.info(f"Backup gemaakt: {backup_path}")
        return backup_path

    except Exception as e:
        logging.warning(f"Backup maken mislukt: {e}")
        return None


def try_open_excel_file_once() -> None:
    global excel_open_attempted
    if excel_open_attempted:
        return
    excel_open_attempted = True

    if sys.platform.startswith("win"):
        try:
            os.startfile(str(filename))
            logging.info("Excelbestand automatisch geopend.")
        except Exception as e:
            logging.warning(f"Kon Excel niet automatisch openen: {e}")


def excel_disconnect() -> None:
    global xl_app, xl_book

    try:
        if xl_book is not None:
            try:
                xl_book.save()
            except Exception:
                pass
            try:
                xl_book.close()
            except Exception:
                pass
    finally:
        xl_book = None

    try:
        if xl_app is not None:
            try:
                xl_app.quit()
            except Exception:
                pass
    finally:
        xl_app = None


def excel_connect() -> None:
    global xl_app, xl_book

    if not filename.exists():
        raise FileNotFoundError(f"Bestand niet gevonden: {filename}")

    if xl_app is not None and xl_book is not None:
        return

    xl_app = xw.App(visible=False, add_book=False)
    xl_app.display_alerts = False
    xl_app.screen_updating = False

    try:
        xl_app.calculation = "manual"
    except Exception:
        pass

    xl_book = xl_app.books.open(
        str(filename),
        update_links=False,
        read_only=False
    )


def excel_ensure(retries: int = 2) -> None:
    last: Exception | None = None
    for attempt in range(retries + 1):
        try:
            excel_connect()
            return
        except Exception as e:
            last = e
            excel_disconnect()
            time.sleep(0.4 * (attempt + 1))
    assert last is not None
    raise last


def sheet_and_last_row():
    assert xl_book is not None
    sh = xl_book.sheets[tabblad]
    last_row = sh.range("A" + str(sh.cells.last_cell.row)).end("up").row
    if last_row < 3:
        last_row = 3
    return sh, last_row


def maybe_save(force: bool = False) -> None:
    global last_save_time
    if xl_book is None:
        return

    now = time.time()
    if force or (now - last_save_time) >= MIN_SAVE_INTERVAL_SEC:
        try:
            xl_book.save()
        finally:
            last_save_time = now


atexit.register(excel_disconnect)


# -------------------------------------------------
# Status / checks
# -------------------------------------------------
def check_excel_ready_start() -> tuple[bool, str]:
    if not filename.exists():
        return False, f"Bestand niet gevonden: {filename}"

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            try:
                sh, last_row = sheet_and_last_row()
            except Exception:
                return False, f"Tabblad '{tabblad}' niet gevonden in {filename.name}"

            barcodes = sh.range(f"A3:A{last_row}").options(ndim=1).value or []
            fvals = sh.range(f"F3:F{last_row}").options(ndim=1).value or []

            if not any(bc is not None and str(bc).strip() != "" for bc in barcodes):
                return False, f"Geen barcodes gevonden in kolom A (vanaf A3) op tabblad '{tabblad}'."

            bad_rows = []
            already_scanned = 0

            for idx, bc in enumerate(barcodes):
                if bc is None or str(bc).strip() == "":
                    continue

                v = fvals[idx] if idx < len(fvals) else None
                if v is None or str(v).strip() == "":
                    continue

                s = str(v).strip()
                if s.upper() == "A":
                    continue

                n = parse_positive_int(s)
                if n is None:
                    bad_rows.append(3 + idx)
                else:
                    already_scanned += 1

            if bad_rows:
                toon = ", ".join(map(str, bad_rows[:10]))
                extra = " …" if len(bad_rows) > 10 else ""
                return False, (
                    f"Niet klaar: tabblad '{tabblad}' kolom F bevat ongeldige waarden. "
                    f"Toegelaten: 'A' of een positief getal. Probleem in rijen: {toon}{extra}."
                )

            if already_scanned > 0:
                return True, (
                    f"Excel ok. Opgelet: er staan al {already_scanned} ingevulde aantallen in kolom F."
                )

        return True, "Excel is klaar om te scannen."

    except Exception as e:
        logging.warning(f"Excel startcheck faalde: {e}")
        excel_disconnect()
        return False, (
            "Excel is niet klaar. Open 'Whistloting.xlsm' één keer in Excel, "
            "klik indien gevraagd op 'Inhoud inschakelen', sluit het bestand en probeer opnieuw."
        )


def ping_excel_light() -> tuple[bool, str]:
    try:
        with excel_lock:
            excel_ensure(retries=1)
            assert xl_book is not None
            sh = xl_book.sheets[tabblad]
            _ = sh.range("A1").value
        return True, "Excel leeft."
    except Exception:
        excel_disconnect()
        return False, "Excel reageert niet meer."


def get_excel_status(force: bool = False) -> tuple[bool, str]:
    global excel_ok, excel_last_check, excel_last_error

    now = time.time()
    if (not force) and (now - excel_last_check) < EXCEL_CHECK_TTL_SEC:
        return excel_ok, (excel_last_error if not excel_ok else "Excel is klaar.")

    ok, msg = check_excel_ready_start()
    excel_ok = ok
    excel_last_check = now
    excel_last_error = "" if ok else msg
    return ok, msg


# -------------------------------------------------
# Excel acties
# -------------------------------------------------
def update_quantity(barcode: str, aantal: int) -> tuple[bool, str, str | None]:
    global total_scans

    barcode = normalize_barcode(barcode)

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            sh, last_row = sheet_and_last_row()
            barcodes = sh.range(f"A3:A{last_row}").options(ndim=1).value or []
            fvals = sh.range(f"F3:F{last_row}").options(ndim=1).value or []

            for idx, cell_value in enumerate(barcodes):
                if cell_value is None or str(cell_value).strip() == "":
                    continue

                if normalize_barcode(cell_value) == barcode:
                    row = 3 + idx
                    huidig_raw = fvals[idx] if idx < len(fvals) else None
                    huidig = parse_positive_int(huidig_raw)
                    naam = sh.range(f"D{row}").value

                    if huidig is not None and huidig > 0:
                        return False, f"{naam} ({barcode}) is al ingevoerd met aantal {huidig}.", None

                    sh.range(f"F{row}").value = int(aantal)

                    scanned_data.append({"barcode": barcode, "naam": naam, "aantal": int(aantal)})
                    total_scans += 1

                    append_scan_log({
                        "actie": "add",
                        "barcode": barcode,
                        "naam": str(naam) if naam is not None else "",
                        "aantal": int(aantal),
                    })

                    if total_scans % SAVE_EVERY_N_SCANS == 0:
                        maybe_save()

                    if total_scans % BACKUP_EVERY_N_SCANS == 0:
                        make_excel_backup(reason="autosave")

                    return True, f"Barcode {barcode} succesvol bijgewerkt.", str(naam) if naam is not None else None

            return False, f"Barcode {barcode} werd niet gevonden in kolom A.", None

    except Exception as e:
        logging.exception("Fout bij update_quantity")
        excel_disconnect()
        return False, f"Fout bij het bijwerken van de barcode: {e}", None


def remove_quantity(barcode: str, aantal: int) -> bool:
    barcode = normalize_barcode(barcode)

    try:
        with excel_lock:
            excel_ensure(retries=2)
            assert xl_book is not None

            sh, last_row = sheet_and_last_row()
            barcodes = sh.range(f"A3:A{last_row}").options(ndim=1).value or []
            fvals = sh.range(f"F3:F{last_row}").options(ndim=1).value or []

            for idx, value in enumerate(barcodes):
                if value is None or str(value).strip() == "":
                    continue

                if normalize_barcode(value) == barcode:
                    huidig_raw = fvals[idx] if idx < len(fvals) else None
                    huidig = parse_positive_int(huidig_raw)

                    if huidig is not None and int(huidig) == int(aantal):
                        rij = 3 + idx
                        sh.range(f"F{rij}").value = "A"

                        append_scan_log({
                            "actie": "remove",
                            "barcode": barcode,
                            "aantal": int(aantal),
                        })

                        maybe_save(force=True)
                        make_excel_backup(reason="remove")
                        return True

            return False

    except Exception:
        logging.exception("Fout bij remove_quantity")
        excel_disconnect()
        return False


# -------------------------------------------------
# Shutdown
# -------------------------------------------------
def shutdown_all() -> None:
    try:
        with excel_lock:
            maybe_save(force=True)
            make_excel_backup(reason="shutdown")
            excel_disconnect()
    finally:
        os._exit(0)


def _handle_sig(signum, frame):
    shutdown_all()


try:
    signal.signal(signal.SIGINT, _handle_sig)
    signal.signal(signal.SIGTERM, _handle_sig)
except Exception:
    pass


# -------------------------------------------------
# Routes
# -------------------------------------------------
@app.route("/excel_status")
def excel_status():
    global excel_warmed_up

    force = request.args.get("force") == "1"

    if excel_warmed_up and not force:
        ok_ping, _ = ping_excel_light()
        if not ok_ping:
            excel_warmed_up = False
            return jsonify({
                "ok": False,
                "msg": "Excel viel weg. Sluit Excel volledig en start opnieuw.",
                "file": str(filename),
                "sheet": tabblad
            })

        return jsonify({
            "ok": True,
            "msg": "Excel is klaar (warm-up ok).",
            "file": str(filename),
            "sheet": tabblad
        })

    if not excel_warmed_up and not force:
        return jsonify({
            "ok": False,
            "msg": "Nog niet getest — scan om te starten.",
            "file": str(filename),
            "sheet": tabblad
        })

    ok, msg = get_excel_status(force=True)
    if ok:
        excel_warmed_up = True

    return jsonify({
        "ok": ok,
        "msg": msg,
        "file": str(filename),
        "sheet": tabblad
    })


@app.route("/shutdown", methods=["POST"])
def shutdown():
    shutdown_all()
    return "OK"


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
            flash("Aantal ronden is vereist.", "error")
            return redirect(url_for("index"))

        if not excel_warmed_up:
            ok_now, msg_now = get_excel_status(force=True)
            if not ok_now:
                try_open_excel_file_once()
                flash(msg_now, "warning")
                flash("Excel werd geopend. Sluit het bestand en scan opnieuw.", "warning")
                return redirect(url_for("index"))
            excel_warmed_up = True

        try:
            aantal = int(aantal_str)
            if aantal < 1 or aantal > 10:
                raise ValueError
        except ValueError:
            flash("Voer een geldig getal in voor het aantal (1–10).", "error")
            return redirect(url_for("index"))

        ok_upd, msg_upd, _ = update_quantity(barcode, aantal)
        flash(msg_upd, "success" if ok_upd else "error")

        if (not ok_upd) and msg_upd.startswith("Fout bij het bijwerken"):
            excel_warmed_up = False

        return redirect(url_for("index"))

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
    scanned_data_sorted=sorted(scanned_data, key=lambda x: (x.get("naam") or "").lower()),
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
    if sys.platform.startswith("win"):
        try:
            pythoncom.CoInitialize()  # type: ignore
            atexit.register(lambda: pythoncom.CoUninitialize())  # type: ignore
        except Exception:
            pass

    atexit.register(lambda: maybe_save(force=True))

    try:
        make_excel_backup(reason="startup")
    except Exception:
        pass

    app.run(debug=False, threaded=False, use_reloader=False)