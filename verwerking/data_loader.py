from typing import Tuple
import os
import csv

# ✅ Maandenlijst (voor consistentie)
maanden = [
    "September",
    "Oktober",
    "November",
    "December",
    "Januari",
    "Februari",
    "Maart",
    "April",
    "Mei",
    "Juni",
]


def score(lijn: str):
    """Verwerk 1 regel uit een CSV-bestand tot (naam, puntenlijst, damesaantal)."""
    if ";" in lijn:
        kolommen = lijn.split(";")
    else:
        kolommen = lijn.split(",")

    if len(kolommen) < 20:
        return "geen lid"

    # Naam opschonen:
    # - spaties voor/achter weg
    # - dubbele spaties in naam weg
    wie = " ".join((kolommen[0] or "").strip().split())
    if not wie:
        return "geen lid"

    punten_str = kolommen[2:6]
    punten = []

    for p in punten_str:
        try:
            p_clean = (p or "").strip()
            if p_clean.isdigit():
                punten.append(int(p_clean))
            else:
                punten.append(0)
        except Exception:
            print(wie, punten_str)
            print("fout:", lijn)
            input("stop")

    # Verplaats 4e score naar voren als één van de eerste drie 0 is
    if len(punten) >= 4 and punten[3] != 0 and 0 in punten[0:3]:
        p = punten.index(0)
        punten[p] = punten[3]
        punten[3] = 0

    try:
        dames = sum(
            int((p or "").strip())
            for p in kolommen[16:20]
            if (p or "").strip()
        )
        if len(wie) > 0 and wie[0].upper() != "Z":
            return wie, punten, dames
        else:
            return "geen lid"
    except Exception:
        print("dames", lijn)
        return "geen lid"


def verwerk(bestandsnaam: str) -> Tuple[dict, dict]:
    """Leest één CSV-bestand in en retourneert:
    - scores_maand: naam -> scores
    - dames_maand: naam -> aantal dames
    """
    scores_maand = {}
    dames_maand = {}

    with open(bestandsnaam, encoding="utf-8") as invoer:
        invoer.readline()  # eerste lijn overslaan
        invoer.readline()  # tweede lijn overslaan
        for lijn in invoer:
            tup = score(lijn)
            try:
                if isinstance(tup, tuple) and len(tup) == 3:
                    wie, punten, dames = tup
                    if len(wie) > 2:
                        # vervang é door _eacute
                        p = wie.find("é")
                        if p >= 0:
                            wie = wie.replace(wie[p:p + 1], "_eacute")

                        scores_maand[wie] = punten
                        dames_maand[wie] = dames
            except Exception:
                print("verwerk:", lijn.strip())
                print(tup)

    return scores_maand, dames_maand


def tel_punten(uitslagen: list) -> int:
    """Telt het aantal > 0 scores in een lijst van lijsten."""
    tel = 0
    for punten in uitslagen:
        for p in punten:
            if p > 0:
                tel += 1
    return tel


def verwerk_klassement(jaar: str, laatste_maand: int) -> Tuple[dict, dict]:
    """Verwerkt alle CSV-bestanden tot en met 'laatste_maand' voor opgegeven jaar.
    Annie telt NIET mee in het klassement.
    """
    maandelijks = {}
    dames = {}

    for i in range(laatste_maand):
        maand = maanden[i]
        pad = os.path.join("data", jaar, f"{maand}.csv")
        scores_maand, dames_maand = verwerk(pad)

        for wie, punten in scores_maand.items():
            # Annie mag wel in controles meetellen, maar niet in klassement
            wie_norm = (wie or "").strip().lower().replace(" ", "")
            if "annie" in wie_norm:
                continue

            if wie not in maandelijks:
                maandelijks[wie] = []
                dames[wie] = []
                for _ in range(i):
                    maandelijks[wie].append([0, 0, 0, 0])
                    dames[wie].append(0)

            gespeeld = tel_punten(maandelijks[wie])
            if gespeeld < 30:
                aantal = min(4, 30 - gespeeld)
                maandelijks[wie].append(punten[0:aantal])

            dames[wie].append(dames_maand.get(wie, 0))

    return maandelijks, dames


def controleer_csv(jaar: str, maand_nr: int) -> bool:
    """Controleer CSV: tel 10/6/3/1 en check dat kolom G (effectief gespeeld) overeenkomt
    met het aantal ingevulde scores in C–F.
    """
    maand = maanden[maand_nr]
    pad = os.path.join("data", jaar, f"{maand}.csv")

    rondes = {10: 0, 6: 0, 3: 0, 1: 0}
    totaal_scores = 0

    fouten = []  # (lineno, naam, gespeeld_g, ingevuld, scores)

    def to_int(s: str):
        s = (s or "").strip()
        if s == "":
            return None
        try:
            return int(float(s.replace(",", ".")))
        except Exception:
            return None

    with open(pad, encoding="utf-8", newline="") as f:
        sample = f.read(2048)
        f.seek(0)
        delim = ";" if sample.count(";") >= sample.count(",") else ","

        reader = csv.reader(f, delimiter=delim)

        # 2 headerlijnen overslaan zoals in jouw verwerk()
        next(reader, None)
        next(reader, None)

        for lineno, row in enumerate(reader, start=3):
            if len(row) < 7:
                continue

            naam = " ".join(((row[0] or "").strip()).split())

            raw_scores = row[2:6]
            scores = []
            for x in raw_scores:
                v = to_int(x)
                scores.append(v if v in (1, 3, 6, 10) else 0)

            ingevuld = sum(1 for v in scores if v != 0)

            gespeeld_g = to_int(row[6])
            if gespeeld_g is None:
                continue

            for v in scores:
                if v != 0:
                    rondes[v] += 1
                    totaal_scores += 1

            if gespeeld_g != ingevuld:
                fouten.append((lineno, naam, gespeeld_g, ingevuld, scores))

    print(f"\nControle van {maand}")
    for k in (10, 6, 3, 1):
        print(f"{k:2d} : {rondes[k]}")
    print(f"Totaal geldige scores: {totaal_scores}")

    if fouten:
        print("❌ Fouten (kolom G ≠ ingevulde rondes in C–F) (max 20):")
        for lineno, naam, g, ingevuld, scores in fouten[:20]:
            print(f"  rij {lineno} | {naam} | G={g} | ingevuld={ingevuld} | rondes={scores}")
    else:
        print("✅ Geen fouten gevonden (G klopt met ingevulde rondes).")

    return True


def main():
    controleer_csv("2025", 0)
    controleer_csv("2025", 1)
    controleer_csv("2025", 2)
    controleer_csv("2025", 3)  # December
    # controleer_csv("2025", 4)  # Januari
    # controleer_csv("2025", 5)
    # controleer_csv("2025", 6)
    # controleer_csv("2025", 7)
    # controleer_csv("2025", 8)
    # controleer_csv("2025", 9)


if __name__ == "__main__":
    main()