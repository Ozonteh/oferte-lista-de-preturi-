"""Importă o listă de prețuri în formatul vechi (Google Sheets exportat .xlsx)
într-un folder de furnizor: produse.csv + imagini/.

Formatul vechi are un rând de cap de tabel cu coloanele (nume acceptate):
    Denumire(a) | Poze | Link la sait | Descriere | Preț / Pret Vinzare |
    Preț Promo | Oferta
O foaie poate avea mai multe secțiuni (ex. câte una pe marcă), fiecare cu
capul ei de tabel; fiecare cap de tabel pornește o secțiune nouă.

Coloanele „Achiziție” și „SKU” se ignoră: prețurile de achiziție nu intră
niciodată în lista pentru clienți.

Linkuri: dacă în celulă e scrisă o adresă, se ia adresa scrisă (hyperlinkul
din spate rămâne uneori vechi sau duce la o ciornă); altfel se ia hyperlinkul.
Diferențele se afișează, ca să fie trecute în VERIFICARI.md.

Poze: contează doar pozele puse în coloana „Poze”. O poză se aplică tuturor
produselor de la rândul ei în jos, până la următoarea poză. Celelalte poze
(sigle) se salvează ca logo.png / logo-<n>.png.

Citirea se oprește după 20 de rânduri goale la rând, ca să nu prindă datele
interne puse mai jos în foaie (prețuri de achiziție, containere).

Utilizare:
    python scripts/importa_lista_veche.py lista.xlsx furnizori/<furnizor>

După import, completați coloana `categorie` și corectați denumirile în
produse.csv, apoi rulați scripts/genereaza_lista.py.
"""

import csv
import io
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
from PIL import Image

CAP_TABEL = {
    "denumire": "denumire",
    "denumirea": "denumire",
    "poze": "poze",
    "poza": "poze",
    "link la sait": "link",
    "link la site": "link",
    "descriere": "descriere",
    "preț": "pret",
    "pret": "pret",
    "pret vinzare": "pret",
    "preț vânzare": "pret",
    "pret vanzare": "pret",
    "preț vinzare": "pret",
    "preț promo": "promo",
    "pret promo": "promo",
    "oferta": "oferta",
    "ofertă": "oferta",
}
IGNORATE = {"achizitie", "achiziție", "sku", "cod"}
RANDURI_GOALE_STOP = 20
LATURA_MAX_IMAGINE = 400
COLOANE_CSV = [
    "categorie", "denumire", "link_produs", "pret_eur", "pret_promo_eur",
    "nota", "link_oferta", "imagine",
]


def normalizeaza(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def slug(text):
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "produs"


def pe_alb(imagine):
    """RGB pe fundal alb. Fără asta, fundalul transparent devine negru."""
    if imagine.mode in ("RGBA", "LA", "P"):
        imagine = imagine.convert("RGBA")
        alb = Image.new("RGBA", imagine.size, "white")
        return Image.alpha_composite(alb, imagine).convert("RGB")
    return imagine.convert("RGB")


def cap_de_tabel(row):
    """Coloanele recunoscute din rând, dacă rândul e un cap de tabel."""
    coloane = {}
    for cell in row:
        cheie = normalizeaza(cell.value).lower()
        if cheie in CAP_TABEL:
            coloane[CAP_TABEL[cheie]] = cell.column
    return coloane if "denumire" in coloane and "pret" in coloane else None


def numar_sau_text(valoare):
    if isinstance(valoare, (int, float)):
        return f"{valoare:g}"
    text = normalizeaza(valoare)
    return "" if text in ("-", "link") else text


def citeste_link(celula, avertismente, denumire):
    text = normalizeaza(celula.value)
    hyperlink = celula.hyperlink.target if celula.hyperlink else ""
    if text.startswith("http"):
        if hyperlink and hyperlink.rstrip("/") != text.rstrip("/"):
            avertismente.append(f"{denumire}: scris {text} | hyperlink {hyperlink} -> am luat textul")
        return text
    return hyperlink


def main(cale_xlsx, folder_furnizor):
    folder = Path(folder_furnizor)
    (folder / "imagini").mkdir(parents=True, exist_ok=True)
    ws = openpyxl.load_workbook(cale_xlsx).active

    col, sectiune, goale = None, 0, 0
    produse, avertismente, col_poze, primul_cap = [], [], set(), None
    for row in ws.iter_rows():
        rand = row[0].row
        cap = cap_de_tabel(row)
        if cap:
            col, sectiune, goale = cap, sectiune + 1, 0
            primul_cap = primul_cap or rand
            if "poze" in cap:
                col_poze.add(cap["poze"])
            ignorate = [normalizeaza(c.value) for c in row
                        if normalizeaza(c.value).lower() in IGNORATE]
            if ignorate:
                print(f"  secțiunea {sectiune} (rândul {rand}): ignor coloanele {ignorate}")
            continue
        if not col:
            continue
        denumire = normalizeaza(ws.cell(rand, col["denumire"]).value)
        if not denumire:
            goale += 1
            if goale >= RANDURI_GOALE_STOP:
                break
            continue
        goale = 0

        def celula(cheie):
            return ws.cell(rand, col[cheie]) if cheie in col else None

        pret = numar_sau_text(celula("pret").value)
        promo = numar_sau_text(celula("promo").value) if "promo" in col else ""
        nota = normalizeaza(celula("descriere").value) if "descriere" in col else ""
        if promo and not re.fullmatch(r"[\d.]+", promo):
            nota, promo = "; ".join(filter(None, [nota, promo])), ""
        oferta = celula("oferta")
        produse.append({
            "rand": rand,
            "sectiune": sectiune,
            "categorie": "",
            "denumire": denumire,
            "link_produs": citeste_link(celula("link"), avertismente, denumire) if "link" in col else "",
            "pret_eur": pret if re.fullmatch(r"[\d.]+", pret) else "",
            "pret_promo_eur": promo,
            "nota": nota,
            "link_oferta": oferta.hyperlink.target if oferta is not None and oferta.hyperlink else "",
            "imagine": "",
        })

    # Pozele din coloana „Poze” sunt ale produselor; restul sunt sigle.
    poze, sigle = [], []
    for img in ws._images:
        rand, coloana = img.anchor._from.row + 1, img.anchor._from.col + 1
        imagine = pe_alb(Image.open(io.BytesIO(img._data())))
        if coloana in col_poze and primul_cap and rand >= primul_cap:
            poze.append((rand, imagine))
        else:
            sigle.append((rand, imagine))
    poze.sort(key=lambda p: p[0])
    sigle.sort(key=lambda p: p[0])
    if len(sigle) == 1:
        sigle[0][1].save(folder / "logo.png")
    for i, (_, imagine) in enumerate(sigle if len(sigle) > 1 else [], start=1):
        imagine.save(folder / f"logo-{i}.png")

    salvate = {}
    for produs in produse:
        sursa = [p for p in poze if p[0] <= produs["rand"]]
        if not sursa:
            continue
        rand_poza, imagine = sursa[-1]
        if rand_poza not in salvate:
            nume = f"{slug(produs['denumire'])}.jpg"
            imagine.thumbnail((LATURA_MAX_IMAGINE, LATURA_MAX_IMAGINE))
            imagine.save(folder / "imagini" / nume, quality=90)
            salvate[rand_poza] = nume
        produs["imagine"] = salvate[rand_poza]

    with open(folder / "produse.csv", "w", newline="", encoding="utf-8-sig") as f:
        scriitor = csv.DictWriter(f, fieldnames=COLOANE_CSV, extrasaction="ignore")
        scriitor.writeheader()
        scriitor.writerows(produse)

    print(f"{len(produse)} produse în {sectiune} secțiuni, {len(salvate)} poze, "
          f"{len(sigle)} sigle -> {folder}")
    for mesaj in avertismente:
        print("  atenție, link:", mesaj)
    fara_pret = [p["denumire"] for p in produse if not p["pret_eur"]]
    if fara_pret:
        print("  fără preț (apar „La cerere”):", "; ".join(fara_pret))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2])
