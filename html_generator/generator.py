import os
from html import escape
from typing import Dict, List

from verwerking.data_loader import verwerk_klassement, tel_punten, maanden

# Afkortingen (eerste 3 letters)
maanden_kort = [m[:3] for m in maanden]


def get_jaar(maand: str, jaar: str) -> str:
    """Competitie Sep–Dec in 'jaar', Jan–Jun in 'jaar+1'."""
    p = maanden.index(maand)  # 0..9
    return jaar if p < 4 else str(int(jaar) + 1)


def html_string(l: List[int | str], sep: str) -> str:
    """Zet sep tussen items, l kan getallen bevatten."""
    return sep.join(str(v) for v in l) if l else ""


def html_string_scores(scores: List[int]) -> str:
    """Genereer <td> cellen; laatste score oranje."""
    if not scores:
        return ""
    res = "".join(f"<td>{v}" for v in scores[:-1])
    res += f'<td style="background-color: orange">{scores[-1]}'
    return res


def _project_root() -> str:
    # generator.py zit in html_generator/, dus 1 map omhoog = projectroot
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _html_dir(jaar: str) -> str:
    return os.path.join(_project_root(), "output", jaar, "html")


def _html_path(jaar: str, filename: str) -> str:
    return os.path.join(_html_dir(jaar), filename)


def _display_name(wie_key: str) -> str:
    """
    Toonbare naam: zet interne tokens terug en escape HTML.
    - '_eacute' -> 'é'
    - Escapet &, <, >, "
    """
    return escape(wie_key.replace("_eacute", "é"))


def _nav_html(jaar: str, laatste_maand: int) -> str:
    """Navigatiebalk met links naar beschikbare maanden, Dames, Controle."""
    links = []
    for i in range(laatste_maand):
        maand = maanden[i]
        label = maanden_kort[i]
        links.append(f'<a href="{maand}.html">{escape(label)}</a>')
    links.append('<a href="Dames.html">Dames</a>')
    if laatste_maand > 1:
        links.append('<a href="Controle.html">Controle</a>')
    sep = '<span style="color:#999;margin:0 6px">|</span>'
    return (
        '<nav class="no-print" style="margin:8px 0 14px 0;font-family:Arial,Helvetica,sans-serif;">'
        + sep.join(links)
        + "</nav>"
    )


def maak_punten(uitslagen: List[List[int]], laatste_maand: int) -> List[List[int]]:
    """
    Van ruwe uitslagen (per maand max 4 punten) naar:
    - per maand exact 3 punten (nul-opvullen)
    - reservepunten (4e) verzamelen en inzetten waar 0 staat
    Resultaat: [[a,b,c], ..., [reserve...]] (laatste item = lijst met niet-ingezette reserves)
    """
    punten_per_maand: List[List[int]] = []
    for punten in uitslagen:
        score = punten[:3]
        while len(score) < 3:
            score.append(0)
        punten_per_maand.append(score)

    while len(punten_per_maand) != laatste_maand:
        punten_per_maand.append([0, 0, 0])

    reserve: List[int] = []
    for punten in uitslagen:
        if len(punten) == 4 and punten[3] > 0:
            reserve.append(punten[3])

    tel = 0
    for i in range(len(punten_per_maand)):
        for j in range(3):
            if punten_per_maand[i][j] == 0 and tel < len(reserve):
                punten_per_maand[i][j] = reserve[tel]
                tel += 1

    niet_ingezet = reserve[tel:]
    punten_per_maand.append(niet_ingezet)
    return punten_per_maand


def _maak_html_lijnen(
    maandelijks: Dict[str, List[List[int]]],
    punten_per_maand: Dict[str, List[List[int]]],
    laatste_maand: int,
) -> Dict[str, str]:
    """Bouwt per speler de HTML-cellen (géén <tr>, alleen aaneengeplakte <td>…-cellen)."""
    html_lijnen: Dict[str, str] = {}

    for wie in punten_per_maand:
        uitslagen = punten_per_maand[wie]  # per maand 3 punten + [reserve] achteraan
        scores = [sum(s) for s in uitslagen[0:laatste_maand]]

        reserve = uitslagen[laatste_maand] if len(uitslagen) > laatste_maand else []

        details: List[int] = []
        if len(maandelijks.get(wie, [])) >= laatste_maand:
            details = [p for p in maandelijks[wie][laatste_maand - 1] if p > 0]

        # Veiligheidsfilter tegen spookrijen:
        # geen enkel punt, geen details, geen reserve => niet tonen
        if sum(scores) == 0 and not details and not reserve:
            continue

        aantal_boompjes = tel_punten(uitslagen[0:laatste_maand])
        wie_p = _display_name(wie)

        cells: List[str] = []
        cells.append(f'<td class="name">{wie_p}</td>')
        cells.append(f"<td>{aantal_boompjes}</td>")
        cells.append(f'<td class="details"><nobr>{html_string(details, "/")}</nobr></td>')
        cells.append(f'<td class="reserve">{html_string(reserve, "/")}</td>')

        if scores:
            for v in scores[:-1]:
                cells.append(f"<td>{v}</td>")
            cells.append(f'<td class="last-score">{scores[-1]}</td>')

        cells.append(f"<td>{sum(scores)}</td>")

        if len(wie.strip()) > 2:
            html_lijnen[wie] = "".join(cells)

    return html_lijnen


def _head_begin(title: str) -> List[str]:
    """Head-lijnen met CSS en inline print-CSS die nav verbergt in PDF,
    én consistente celhoogte + vaste nummerkolom.
    """
    return [
        "<!DOCTYPE html>",
        '<html lang="nl">',
        "<head>",
        '  <meta charset="utf-8">',
        f"  <title>{title}</title>",
        '  <link rel="stylesheet" type="text/css" href="extra.css" />',
        "  <style>",
        "    @media print { .no-print, nav.no-print { display: none !important; } }",
        "    table { border-collapse: collapse; }",
        "    th, td { padding: 2px 6px; vertical-align: middle; line-height: 1.25; }",
        "    tr { page-break-inside: avoid; }",
        "    td.rown { width: 28px; text-align: right; white-space: nowrap; }",
        "    td.name { white-space: nowrap; }",
        "    td.last-score { background-color: orange; }",
        "    td.details, td.reserve { text-align: center; font-style: italic; }",
        "  </style>",
        "</head>",
        "<body>",
    ]


def _maak_klassement_bestand(
    html_lijnen: Dict[str, str], jaar: str, laatste_maand: int
) -> None:
    maand = maanden[laatste_maand - 1]
    pad = _html_path(jaar, f"{maand}.html")
    os.makedirs(os.path.dirname(pad), exist_ok=True)

    with open(pad, "w", encoding="utf-8") as html:
        for lijn in _head_begin(f"Klassement {maand} {get_jaar(maand, jaar)}"):
            print(lijn, file=html)

        print(_nav_html(jaar, laatste_maand), file=html)

        print(f"<h1>{maand} {get_jaar(maand, jaar)}</h1>", file=html)
        print('<table border="1">', file=html)

        print("<tr><th>#<th>naam<th>GS<th>detail<th>extra", file=html)
        print(f'<td>{html_string(maanden_kort[:laatste_maand], "<td>")}', file=html)
        print("<th>TOT", file=html)

        nr = 1
        for wie in html_lijnen:
            print(f'<tr><td class="rown">{nr})</td>{html_lijnen[wie]}', file=html)
            nr += 1

        print("</table>", file=html)
        print("</body>", file=html)
        print("</html>", file=html)

    print(f"{pad} gewijzigd")


def _maak_overzicht_dames(jaar: str, dames: Dict[str, List[int]]) -> None:
    pad = _html_path(jaar, "Dames.html")
    os.makedirs(os.path.dirname(pad), exist_ok=True)

    dames_gesorteerd = dict(
        sorted(dames.items(), key=lambda x: sum(x[1]), reverse=True)
    )

    beschikbare_mnd = max((len(v) for v in dames.values()), default=0)

    with open(pad, "w", encoding="utf-8") as html:
        for lijn in _head_begin("Dames totaal"):
            print(lijn, file=html)

        print(_nav_html(jaar, beschikbare_mnd), file=html)

        print("<h1>Dames totaal</h1>", file=html)
        print("<ol>", file=html)
        for wie, aantallen in dames_gesorteerd.items():
            wie_p = _display_name(wie)
            print(f"<li>{wie_p}: {sum(aantallen)} </li>", file=html)
        print("</ol>", file=html)

        print("</body>", file=html)
        print("</html>", file=html)

    print(f"{pad} gemaakt")


def _maak_controle_bestand(
    jaar: str,
    laatste_maand: int,
    html_lijnen_vorig: Dict[str, str],
    html_lijnen_nu: Dict[str, str],
) -> None:
    maand = maanden[laatste_maand - 1]
    vorige_maand = maanden[laatste_maand - 2]
    pad = _html_path(jaar, "Controle.html")
    os.makedirs(os.path.dirname(pad), exist_ok=True)

    with open(pad, "w", encoding="utf-8") as html:
        for lijn in _head_begin(f"Controle voor {maand}"):
            print(lijn, file=html)

        print(_nav_html(jaar, laatste_maand), file=html)

        print('<table border="0" style="width:100%; table-layout:fixed;">', file=html)
        print("<tr>", file=html)

        print('<td style="vertical-align:top; width:50%;">', file=html)
        print(
            f'<h2 style="margin:0 0 8px 0;">{vorige_maand} {get_jaar(vorige_maand, jaar)}</h2>',
            file=html,
        )
        print('<table border="1" style="width:100%;">', file=html)
        print("<tr><th>naam<th>GS<th>detail<th>extra", file=html)
        print(f'<td>{html_string(maanden_kort[:laatste_maand - 1], "<td>")}', file=html)
        print("<th>TOT", file=html)
        for wie in html_lijnen_vorig:
            print(f"<tr>{html_lijnen_vorig[wie]}", file=html)
        print("</table>", file=html)
        print("</td>", file=html)

        print('<td style="vertical-align:top; width:50%;">', file=html)
        print(
            f'<h2 style="margin:0 0 8px 0;">{maand} {get_jaar(maand, jaar)}</h2>',
            file=html,
        )
        print('<table border="1" style="width:100%;">', file=html)
        print("<tr><th>naam<th>GS<th>detail<th>extra", file=html)
        print(f'<td>{html_string(maanden_kort[:laatste_maand], "<td>")}', file=html)
        print("<th>TOT", file=html)
        for wie in html_lijnen_nu:
            print(f"<tr>{html_lijnen_nu[wie]}", file=html)
        print("</table>", file=html)
        print("</td>", file=html)

        print("</tr>", file=html)
        print("</table>", file=html)

        print("</body>", file=html)
        print("</html>", file=html)

    print(f"{pad} gewijzigd voor {maand}")


# ===================== Publieke functies =====================


def maak_html(jaar: str, laatste_maand: int) -> None:
    """Schrijf het klassement voor 'laatste_maand'."""
    try:
        maandelijks, _ = verwerk_klassement(jaar, laatste_maand)

        punten_per_maand = {
            wie: maak_punten(maandelijks[wie], laatste_maand) for wie in maandelijks
        }

        html_lijnen = _maak_html_lijnen(maandelijks, punten_per_maand, laatste_maand)

        # Belangrijk:
        # enkel sorteren op spelers die effectief in html_lijnen zitten
        score = {
            wie: sum(sum(s) for s in punten_per_maand[wie][0:laatste_maand])
            for wie in html_lijnen
        }

        scores_gesorteerd = dict(
            sorted(score.items(), key=lambda x: x[1], reverse=True)
        )

        for wie in list(scores_gesorteerd.keys()):
            scores_gesorteerd[wie] = html_lijnen[wie]

        _maak_klassement_bestand(scores_gesorteerd, jaar, laatste_maand)

    except FileNotFoundError:
        print("Er ontbreken maanden (CSV-bestanden niet gevonden).")


def maak_controle_bestand(jaar: str, laatste_maand: int) -> None:
    """Maak Dames.html en (vanaf maand 2) Controle.html."""
    try:
        maandelijks_nu, dames = verwerk_klassement(jaar, laatste_maand)
    except FileNotFoundError as e:
        print(f"[Controle] CSV ontbreekt voor maand {laatste_maand}: {e}")
        return

    _maak_overzicht_dames(jaar, dames)

    if laatste_maand < 2:
        print("Geen controlebestand gemaakt (te weinig maanden).")
        return

    ppm_nu = {w: maak_punten(maandelijks_nu[w], laatste_maand) for w in maandelijks_nu}
    lijnen_nu = _maak_html_lijnen(maandelijks_nu, ppm_nu, laatste_maand)

    maandelijks_vor, _ = verwerk_klassement(jaar, laatste_maand - 1)
    ppm_vor = {
        w: maak_punten(maandelijks_vor[w], laatste_maand - 1) for w in maandelijks_vor
    }
    lijnen_vor = _maak_html_lijnen(maandelijks_vor, ppm_vor, laatste_maand - 1)

    _maak_controle_bestand(jaar, laatste_maand, lijnen_vor, lijnen_nu)