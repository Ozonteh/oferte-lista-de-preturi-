"""Generează lista de prețuri Ozonteh pentru un furnizor, în format Excel.

Fișierul rezultat are trei foi:
  - „Lista de prețuri”   produsele pe categorii: poză, denumire, preț EUR,
                         preț promo, preț MDL, notă, link la site, link la ofertă;
  - „Ofertă comercială”  ofertă pentru client, cu produsele alese din listă;
  - „Catalog”            foaie ascunsă din care citește oferta (nu se editează).

Utilizare:
    python scripts/genereaza_lista.py furnizori/shaktiman
    python scripts/genereaza_lista.py --toti        # toți furnizorii

Rezultat: output/Lista de preturi - <Furnizor>.xlsx
"""

import csv
import datetime as dt
import io
import json
import re
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from PIL import Image, ImageChops

RADACINA = Path(__file__).resolve().parent.parent
COLOANE_CSV = [
    "categorie", "denumire", "link_produs", "pret_eur", "pret_promo_eur",
    "nota", "link_oferta", "imagine",
]
FOAIE_LISTA = "Lista de prețuri"
FOAIE_OFERTA = "Ofertă comercială"
FOAIE_CATALOG = "Catalog"
FOAIE_INSTRUCTIUNI = "Instrucțiuni"
# Tabelul de produse din pagina de start. Scriptul scripts/apps_script/oferta_completa.gs
# citește produsele și linkurile prin numele OfertaProduse / OfertaLinkuri.
LINII_OFERTA = 6
PRIMA_LINIE_OFERTA = 17
# Google Sheets nu are indentare: textul de pe benzi primește spații în față.
SPATIU = "  "
EMU_PE_PIXEL = 9525

FONT = "Arial"
# Pagina de start a ofertei urmează designul Ozonteh: titlu League Spartan, text Open Sans.
FONT_TITLU = "League Spartan"
FONT_TEXT = "Open Sans"
GRI_TEXT = "666666"
GRI = "595959"
GRI_DESCHIS = "D9D9D9"
LINK = "0563C1"
ROSU_PROMO = "C00000"
GALBEN_INPUT = "FFF2CC"
FORMAT_EUR = '#,##0 "€"'
FORMAT_MDL = '#,##0 "lei"'
FORMAT_DATA = "DD.MM.YYYY"


# ---------------------------------------------------------------- utilitare

def font(**kw):
    kw.setdefault("name", FONT)
    kw.setdefault("size", 10)
    return Font(**kw)


def umplere(culoare):
    return PatternFill("solid", fgColor=culoare)


def chenar_jos(culoare=GRI_DESCHIS, stil="thin"):
    return Border(bottom=Side(style=stil, color=culoare))


def chenar_tot(culoare=GRI_DESCHIS):
    s = Side(style="thin", color=culoare)
    return Border(left=s, right=s, top=s, bottom=s)


def scrie(ws, coord, valoare=None, **stil):
    celula = ws[coord]
    if valoare is not None:
        celula.value = valoare
    for cheie, val in stil.items():
        setattr(celula, cheie, val)
    return celula


def link(ws, coord, url, text):
    celula = ws[coord]
    celula.value = text
    celula.hyperlink = url
    celula.font = font(color=LINK, underline="single")
    return celula


def latime_px(ws, litera):
    return int(ws.column_dimensions[litera].width * 7 + 5)


def taie_margini_albe(im, prag=18, margine=0.04):
    """Taie marginile albe din jurul utilajului, ca poza să umple celula."""
    rgb = im.convert("RGB")
    diferenta = ImageChops.difference(rgb, Image.new("RGB", rgb.size, "white")).convert("L")
    cadru = diferenta.point(lambda p: 255 if p > prag else 0).getbbox()
    if not cadru:
        return im
    pad = round(max(cadru[2] - cadru[0], cadru[3] - cadru[1]) * margine)
    cadru = (max(0, cadru[0] - pad), max(0, cadru[1] - pad),
             min(im.width, cadru[2] + pad), min(im.height, cadru[3] + pad))
    return im.crop(cadru)


def alb_transparent(im, prag=235):
    """Fundalul alb al unei sigle devine transparent, ca sigla să stea pe banda colorată."""
    im = im.convert("RGBA")
    alfa = im.convert("L").point(lambda p: 0 if p >= prag else 255)
    im.putalpha(ImageChops.multiply(im.getchannel("A"), alfa))
    return im


def pune_imagine(ws, sursa, col, rand, max_l, max_h, dx=None, dy=None, taie=False,
                 zona_dreapta=None, transparent=False):
    """Pune o imagine încadrată în max_l x max_h px, la celula (col, rand).

    `col` și `rand` încep de la 0. Fără `dx`, imaginea se centrează pe lățimea
    coloanei; cu `zona_dreapta` (px) se aliniază la dreapta zonei. Imaginea se
    salvează la dublul dimensiunii, ca să fie clară.
    """
    im = Image.open(sursa)
    im = im.convert("RGBA" if im.mode in ("RGBA", "LA", "P") else "RGB")
    if taie and im.mode == "RGB":
        im = taie_margini_albe(im)
    if transparent:
        im = alb_transparent(im)
    raport = min(max_l / im.width, max_h / im.height)
    l, h = round(im.width * raport), round(im.height * raport)
    copie = im.copy()
    copie.thumbnail((l * 2, h * 2))
    buf = io.BytesIO()
    if copie.mode == "RGBA":
        copie.save(buf, format="PNG", optimize=True)
    else:
        copie.save(buf, format="JPEG", quality=88)
    buf.seek(0)
    img = XLImage(buf)
    if zona_dreapta is not None:
        dx = max(0, zona_dreapta - l - 24)
    if dx is None:
        litera = ws.cell(1, col + 1).column_letter
        dx = max(0, (latime_px(ws, litera) - l) // 2)
    if dy is None:
        dy = max(0, (round(ws.row_dimensions[rand + 1].height * 4 / 3) - h) // 2)
    img.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col, colOff=dx * EMU_PE_PIXEL, row=rand, rowOff=dy * EMU_PE_PIXEL),
        ext=XDRPositiveSize2D(l * EMU_PE_PIXEL, h * EMU_PE_PIXEL),
    )
    ws.add_image(img)


def numar(text):
    text = (text or "").strip().replace(" ", "").replace(",", ".")
    if not text:
        return None
    if not re.fullmatch(r"\d+(\.\d+)?", text):
        raise ValueError(text)
    valoare = float(text)
    return int(valoare) if valoare.is_integer() else valoare


# ---------------------------------------------------------------- date

def citeste_furnizor(folder):
    folder = Path(folder)
    config = json.loads((folder / "furnizor.json").read_text(encoding="utf-8"))
    with open(folder / "produse.csv", encoding="utf-8-sig", newline="") as f:
        cititor = csv.DictReader(f)
        lipsa = set(COLOANE_CSV) - set(cititor.fieldnames or [])
        if lipsa:
            raise SystemExit(f"{folder}/produse.csv: lipsesc coloanele {sorted(lipsa)}")
        randuri = list(cititor)

    erori, avertismente, produse, vazute = [], [], [], set()
    for nr, r in enumerate(randuri, start=2):
        r = {k: (r.get(k) or "").strip() for k in COLOANE_CSV}
        if not r["denumire"]:
            continue
        unde = f"produse.csv rândul {nr} ({r['denumire']})"
        if r["denumire"] in vazute:
            erori.append(f"{unde}: denumire dublată — oferta n-ar ști ce preț să ia")
        vazute.add(r["denumire"])
        if not r["categorie"]:
            erori.append(f"{unde}: lipsește categoria")
        try:
            r["pret_eur"] = numar(r["pret_eur"])
            r["pret_promo_eur"] = numar(r["pret_promo_eur"])
        except ValueError as e:
            erori.append(f"{unde}: preț nenumeric „{e}” (lăsați gol pentru „La cerere”)")
            continue
        if r["pret_eur"] is None:
            avertismente.append(f"{unde}: fără preț, apare „La cerere”")
        if r["pret_promo_eur"] and r["pret_eur"] and r["pret_promo_eur"] >= r["pret_eur"]:
            avertismente.append(f"{unde}: prețul promo nu e mai mic decât prețul de listă")
        for cheie in ("link_produs", "link_oferta"):
            if r[cheie] and not r[cheie].startswith("http"):
                erori.append(f"{unde}: {cheie} nu e un link valid")
        if not r["link_produs"]:
            avertismente.append(f"{unde}: lipsește linkul la produs")
        if r["imagine"] and not (folder / "imagini" / r["imagine"]).exists():
            avertismente.append(f"{unde}: nu găsesc imagini/{r['imagine']}")
            r["imagine"] = ""
        produse.append(r)

    categorii = {p["categorie"] for p in produse}
    for categorie, logo in config.get("logouri_categorii", {}).items():
        if categorie not in categorii:
            avertismente.append(f"furnizor.json: sigla pentru categoria „{categorie}”, care nu există")
        elif not (folder / logo).exists():
            avertismente.append(f"furnizor.json: nu găsesc {logo}")
    for mesaj in avertismente:
        print("  atenție:", mesaj)
    if erori:
        raise SystemExit("Erori în date:\n  " + "\n  ".join(erori))
    if not produse:
        raise SystemExit(f"{folder}/produse.csv nu are produse.")
    return config, produse


def pe_categorii(produse):
    categorii = {}
    for p in produse:
        categorii.setdefault(p["categorie"], []).append(p)
    return categorii


# ---------------------------------------------------------------- foaia 1

def foaia_lista(wb, brand, config, produse, folder):
    ws = wb.active
    ws.title = FOAIE_LISTA
    ws.sheet_view.showGridLines = False
    culoare = brand["culoare_principala"]

    latimi = {"A": 5, "B": 19, "C": 50, "D": 13, "E": 13, "F": 15, "G": 22, "H": 15, "I": 16}
    for litera, latime in latimi.items():
        ws.column_dimensions[litera].width = latime

    # Antet compact: trei rânduri subțiri + capul de tabel, înghețate împreună.
    for rand, inaltime in ((1, 28), (2, 16), (3, 20)):
        ws.row_dimensions[rand].height = inaltime
    pune_imagine(ws, RADACINA / "brand" / "logo-ozonteh.png", 0, 0, 140, 46, dx=8, dy=16)
    scrie(ws, "C1", "LISTA DE PREȚURI", font=font(size=18, bold=True, color=culoare),
          alignment=Alignment(vertical="bottom"))
    scrie(ws, "C2", config["titlu"], font=font(size=11, color=GRI),
          alignment=Alignment(vertical="center"))
    if config.get("logo") and (folder / config["logo"]).exists():
        pune_imagine(ws, folder / config["logo"], 7, 0, 200, 46, dx=12, dy=14, taie=True)

    scrie(ws, "C3", "Prețuri în EUR, fără TVA", font=font(bold=True),
          alignment=Alignment(vertical="center"))
    scrie(ws, "E3", "Curs EUR/MDL:", font=font(bold=True),
          alignment=Alignment(horizontal="right", vertical="center"))
    # Curs live în Google Sheets; Excel nu are GOOGLEFINANCE (#NAME?), deci IFERROR
    # cade pe cursul de rezervă din brand/ozonteh.json.
    scrie(ws, "F3", f'=IFERROR(GOOGLEFINANCE("CURRENCY:EURMDL"),{brand["curs_eur_mdl"]})',
          font=font(bold=True, color=culoare),
          border=chenar_tot(), number_format="0.00",
          alignment=Alignment(horizontal="center", vertical="center"))
    actualizat = dt.date.fromisoformat(config["data_actualizare"]).strftime("%d.%m.%Y")
    scrie(ws, "G3", f"Actualizat: {actualizat}", font=font(color=GRI),
          alignment=Alignment(horizontal="right", vertical="center"))
    wb.defined_names["Curs"] = DefinedName("Curs", attr_text=f"'{FOAIE_LISTA}'!$F$3")

    # Cap de tabel
    cap = ["Nr.", "Foto", "Denumire produs", "Preț EUR\n(fără TVA)", "Preț promo\nEUR",
           "Preț MDL*\n(fără TVA)", "Notă", "Pagina\nprodusului", "Ofertă\ncomercială"]
    RAND_CAP = 4
    ws.row_dimensions[RAND_CAP].height = 28
    for i, text in enumerate(cap, start=1):
        c = ws.cell(RAND_CAP, i, text if i != 3 else SPATIU + text)
        c.font = font(bold=True, color="FFFFFF")
        c.fill = umplere(culoare)
        c.alignment = Alignment(horizontal="center" if i != 3 else "left",
                                vertical="center", wrap_text=True)
    ws.freeze_panes = f"A{RAND_CAP + 1}"
    ws.print_title_rows = f"{RAND_CAP}:{RAND_CAP}"

    # Produse pe categorii
    rand, nr, randuri_produse = RAND_CAP + 1, 0, []
    for categorie, lista in pe_categorii(produse).items():
        ws.merge_cells(start_row=rand, start_column=1, end_row=rand, end_column=9)
        c = ws.cell(rand, 1, SPATIU + (f"{categorie.upper()}  ·  {len(lista)} modele"
                                       if len(lista) > 1 else categorie.upper()))
        c.font = font(size=11, bold=True, color=culoare)
        c.fill = umplere("E3EEF6")
        c.alignment = Alignment(vertical="center")
        logo_categorie = folder / config.get("logouri_categorii", {}).get(categorie, "-")
        ws.row_dimensions[rand].height = 30 if logo_categorie.is_file() else 24
        if logo_categorie.is_file():
            pune_imagine(ws, logo_categorie, 7, rand - 1, 130, 30, taie=True, transparent=True,
                         zona_dreapta=latime_px(ws, "H") + latime_px(ws, "I"))
        rand += 1
        for p in lista:
            nr += 1
            ws.row_dimensions[rand].height = 46
            centru = Alignment(horizontal="center", vertical="center", wrap_text=True)
            scrie(ws, f"A{rand}", nr, font=font(color=GRI), alignment=centru)
            if p["imagine"]:
                pune_imagine(ws, folder / "imagini" / p["imagine"], 1, rand - 1, 96, 54, taie=True)
            scrie(ws, f"C{rand}", p["denumire"], font=font(size=11),
                  alignment=Alignment(vertical="center", wrap_text=True, indent=1))
            if p["pret_eur"] is None:
                scrie(ws, f"D{rand}", "La cerere", font=font(italic=True, color=GRI), alignment=centru)
            else:
                scrie(ws, f"D{rand}", p["pret_eur"], font=font(size=11, bold=True),
                      number_format=FORMAT_EUR, alignment=centru)
            if p["pret_promo_eur"] is not None:
                scrie(ws, f"E{rand}", p["pret_promo_eur"], font=font(size=11, bold=True, color=ROSU_PROMO),
                      number_format=FORMAT_EUR, alignment=centru)
            scrie(ws, f"F{rand}",
                  f'=IF(ISNUMBER(E{rand}),ROUND(E{rand}*Curs,0),'
                  f'IF(ISNUMBER(D{rand}),ROUND(D{rand}*Curs,0),"La cerere"))',
                  font=font(), number_format=FORMAT_MDL, alignment=centru)
            scrie(ws, f"G{rand}", p["nota"] or None, font=font(size=9, color=GRI),
                  alignment=Alignment(vertical="center", wrap_text=True))
            for col, url, text in (("H", p["link_produs"], "Vezi pe site"),
                                   ("I", p["link_oferta"], "Deschide oferta")):
                if url:
                    link(ws, f"{col}{rand}", url, text).alignment = centru
                else:
                    scrie(ws, f"{col}{rand}", "—", font=font(color=GRI_DESCHIS), alignment=centru)
            for col in "ABCDEFGHI":
                ws[f"{col}{rand}"].border = chenar_jos()
            randuri_produse.append((rand, p))
            rand += 1

    # Note
    rand += 1
    note = [
        "* Preț MDL = prețul final (promo, dacă există) × cursul EUR/MDL din celula F3. "
        "Este orientativ.",
        "Cursul din F3 vine live din Google Finance (în Google Sheets). În Excel, sau dacă Google "
        "Finance nu răspunde, se folosește cursul de rezervă "
        + f"{brand['curs_eur_mdl']:.2f}".replace(".", ",") + ".",
        "Pentru o ofertă cu unul sau mai multe produse folosiți foaia „Ofertă comercială”.",
        f"{brand['firma']} · {brand['adresa']} · {brand['email']} · "
        f"{brand['site'].replace('https://', '').rstrip('/')}",
    ]
    for text in note:
        scrie(ws, f"C{rand}", text, font=font(size=9, italic=True, color=GRI))
        rand += 1

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "Pagina &P din &N"
    ws.oddFooter.center.size = 8
    return randuri_produse


# ---------------------------------------------------------------- foaia 3

def foaia_catalog(wb, randuri_produse):
    ws = wb.create_sheet(FOAIE_CATALOG)
    cap = ["Denumire", "Link produs", "Preț final EUR", "Preț listă EUR", "Preț promo EUR",
           "Link ofertă", "Categorie"]
    for i, text in enumerate(cap, start=1):
        ws.cell(1, i, text).font = font(bold=True)
    lista = f"'{FOAIE_LISTA}'!"
    for i, (rand, p) in enumerate(randuri_produse, start=2):
        ws.cell(i, 1, f"={lista}C{rand}")
        ws.cell(i, 2, p["link_produs"] or None)
        ws.cell(i, 3, f'=IF(ISNUMBER({lista}E{rand}),{lista}E{rand},'
                      f'IF(ISNUMBER({lista}D{rand}),{lista}D{rand},"La cerere"))')
        ws.cell(i, 4, f"={lista}D{rand}")
        ws.cell(i, 5, f'=IF(ISNUMBER({lista}E{rand}),{lista}E{rand},"")')
        ws.cell(i, 6, p["link_oferta"] or None)
        ws.cell(i, 7, p["categorie"])
    for litera, latime in zip("ABCDEFG", (55, 60, 14, 14, 14, 60, 30)):
        ws.column_dimensions[litera].width = latime
    ws.sheet_state = "hidden"
    return len(randuri_produse) + 1


# ---------------------------------------------------------------- foaia 2

def foaia_oferta(wb, brand, produse, ultim_rand_catalog):
    """Pagina de start a ofertei, după designul Ozonteh („Oferta KVINT”).

    Tabelul de produse e în A:H; coloana I (ascunsă, nu se tipărește) ține linkul
    la oferta fiecărui produs ales, din care meniul „Ozonteh” face PDF-ul complet.
    Câmpurile pe care le citește scriptul au nume (Oferta*), nu adrese fixe.
    """
    ws = wb.create_sheet(FOAIE_OFERTA, 1)
    ws.sheet_view.showGridLines = False
    culoare = brand["culoare_principala"]
    titlu = brand["culoare_titlu"]
    rep = citeste_reprezentant()
    privat = RADACINA / "brand" / "privat"
    for litera, latime in {"A": 5, "B": 40, "C": 12, "D": 7, "E": 13, "F": 9, "G": 13, "H": 15,
                           "I": 40}.items():
        ws.column_dimensions[litera].width = latime
    ws.column_dimensions["I"].hidden = True

    def text(**kw):
        kw.setdefault("name", FONT_TEXT)
        kw.setdefault("bold", True)
        kw.setdefault("color", GRI_TEXT)
        return font(**kw)

    def galben_daca(interval, conditie):
        # Galben doar cât timp câmpul cere completare: oferta completată se tipărește curată.
        ws.conditional_formatting.add(
            interval, FormulaRule(formula=[conditie], fill=umplere(GALBEN_INPUT), stopIfTrue=False))

    def nume(nume_camp, interval):
        wb.defined_names[nume_camp] = DefinedName(nume_camp, attr_text=f"'{FOAIE_OFERTA}'!{interval}")

    def merge(interval, valoare, **stil):
        ws.merge_cells(interval)
        return scrie(ws, interval.split(":")[0], valoare, **stil)

    inaltimi = {1: 8, 2: 26, 3: 26, 4: 22, 5: 10, 6: 48, 7: 48, 8: 10, 9: 30, 10: 10, 11: 24,
                12: 24, 13: 24, 14: 14, 15: 22, 16: 30, 23: 24, 24: 16, 29: 10, 33: 14, 34: 16}
    for r in range(1, 40):
        ws.row_dimensions[r].height = inaltimi.get(r, 20)

    # Antet: siglă, link + QR spre site, ilustrație
    pune_imagine(ws, RADACINA / "brand" / "logo-ozonteh.png", 0, 1, 230, 72, dx=4, dy=2)
    c = merge("C2:E4", "Apasă aici ➡️, sau scanează\npentru site-ul Ozonteh!",
              font=font(name="Verdana", size=9, bold=True, color="005681", underline="single"),
              alignment=Alignment(horizontal="center", vertical="center", wrap_text=True))
    c.hyperlink = brand["site"]
    pune_imagine(ws, RADACINA / "brand" / "qr-ozonteh-net.png", 5, 1, 64, 64, dx=0, dy=6)
    pune_imagine(ws, RADACINA / "brand" / "ilustratie.png", 6, 0, 160, 92, dy=6, taie=True,
                 zona_dreapta=latime_px(ws, "G") + latime_px(ws, "H"))

    # Titlu
    for r, cuvant in ((6, "OFERTĂ"), (7, "COMERCIALĂ")):
        merge(f"A{r}:E{r}", SPATIU + cuvant, font=font(name=FONT_TITLU, size=42, bold=True, color=titlu),
              alignment=Alignment(vertical="center"))

    # Reprezentant (dreapta sus)
    merge("G6:H6", "Reprezentantul tău\nOzonteh", font=text(size=10),
          alignment=Alignment(horizontal="center", vertical="bottom", wrap_text=True))
    for col in "GH":
        ws[f"{col}6"].border = chenar_jos("595959", "medium")
    if rep.get("poza") and (privat / rep["poza"]).is_file():
        pune_imagine(ws, privat / rep["poza"], 6, 6, 120, 118, dy=4,
                     zona_dreapta=(latime_px(ws, "G") + latime_px(ws, "H") + 120) // 2 + 8)
    merge("G11:H11", rep.get("nume", ""), font=text(size=12),
          alignment=Alignment(horizontal="center", vertical="center"))
    cutie = Border(*(Side(style="medium", color="595959"),) * 4)
    for r, valoare, url in ((12, f"📞 {rep.get('telefon', '')}", None),
                            (13, rep.get("email", ""), f"mailto:{rep.get('email', '')}")):
        c = merge(f"G{r}:H{r}", valoare, font=text(size=11),
                  alignment=Alignment(horizontal="center", vertical="center"))
        if url and rep.get("email"):
            c.hyperlink = url
        for col in "GH":
            ws[f"{col}{r}"].border = cutie

    # Client, dată, număr, valabilitate
    merge("A9:F9", None, font=text(size=16), number_format='"Pentru:  "@',
          alignment=Alignment(vertical="center"))
    galben_daca("A9:F9", '$A$9=""')
    nume("OfertaClient", "$A$9")
    merge("A11:B11", "=TODAY()", font=text(size=12), number_format="dd.mmmm.yyyy",
          alignment=Alignment(horizontal="center", vertical="center"))
    nume("OfertaData", "$A$11")
    merge("C11:D11", None, font=text(size=12), number_format='"Nr. "0;;;"Nr. "@',
          alignment=Alignment(horizontal="center", vertical="center"))
    galben_daca("C11:D11", '$C$11=""')
    nume("OfertaNr", "$C$11")
    scrie(ws, "E11", "Valabilitate", font=text(size=12),
          alignment=Alignment(horizontal="center", vertical="center"))
    scrie(ws, "F11", brand["valabilitate_oferta_zile"], font=text(size=12), number_format='0" Zile"',
          alignment=Alignment(horizontal="center", vertical="center"))
    dv_zile = DataValidation(type="list", allow_blank=True,
                             formula1='"' + ",".join(str(z) for z in brand["valabilitate_variante"]) + '"')
    ws.add_data_validation(dv_zile)
    dv_zile.add("F11")
    separator = Side(style="medium", color="595959")
    for coord in ("C11", "E11"):
        ws[coord].border = Border(left=separator)

    # Produse
    merge("A15:H15", SPATIU + "PRODUSE OFERITE", font=font(size=11, bold=True, color="FFFFFF"),
          fill=umplere(culoare), alignment=Alignment(vertical="center"))
    cap = ["Nr.", "Produs (alegeți din listă)", "Pagina\nprodusului", "Cant.",
           "Preț unitar\nEUR fără TVA", "Reducere\n%", "Total EUR\nfără TVA", "Total MDL\nfără TVA"]
    for i, valoare in enumerate(cap, start=1):
        c = ws.cell(16, i, valoare)
        c.font = font(bold=True)
        c.fill = umplere("E3EEF6")
        c.border = chenar_tot()
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    scrie(ws, "I16", "Link ofertă produs (pentru meniul Ozonteh)", font=font(size=9, color=GRI))

    prima, ultima = PRIMA_LINIE_OFERTA, PRIMA_LINIE_OFERTA + LINII_OFERTA - 1
    nume_cat = f"{FOAIE_CATALOG}!$A$2:$A${ultim_rand_catalog}"
    linkuri = f"{FOAIE_CATALOG}!$B$2:$B${ultim_rand_catalog}"
    preturi = f"{FOAIE_CATALOG}!$C$2:$C${ultim_rand_catalog}"
    oferte_produse = f"{FOAIE_CATALOG}!$F$2:$F${ultim_rand_catalog}"
    dv = DataValidation(type="list", formula1=nume_cat, allow_blank=True, showErrorMessage=True,
                        errorTitle="Produs necunoscut",
                        error="Alegeți produsul din listă (săgeata din dreapta celulei).")
    ws.add_data_validation(dv)
    centru = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(prima, ultima + 1):
        ws.row_dimensions[r].height = 26
        scrie(ws, f"A{r}", f'=IF(B{r}="","",ROW()-{prima - 1})', font=font(color=GRI), alignment=centru)
        scrie(ws, f"B{r}", font=font(), alignment=Alignment(vertical="center", wrap_text=True))
        dv.add(f"B{r}")
        # Galben: primul produs cât e gol și cantitatea lipsă; rândurile goale rămân albe la tipar.
        if r == prima:
            galben_daca(f"B{r}", f'$B{r}=""')
        galben_daca(f"D{r}", f'AND($B{r}<>"",$D{r}="")')
        link_produs = f"INDEX({linkuri},MATCH(B{r},{nume_cat},0))"
        scrie(ws, f"C{r}",
              f'=IF(B{r}="","",IFERROR(IF({link_produs}&""="","",'
              f'HYPERLINK({link_produs},"Vezi produs")),""))',
              font=font(color=LINK, underline="single"), alignment=centru)
        scrie(ws, f"D{r}", font=font(), alignment=centru)
        scrie(ws, f"E{r}",
              f'=IF(B{r}="","",IFERROR(INDEX({preturi},MATCH(B{r},{nume_cat},0)),"nu e în listă"))',
              font=font(), number_format=FORMAT_EUR, alignment=centru)
        scrie(ws, f"F{r}", font=font(), number_format="0%", alignment=centru)
        scrie(ws, f"G{r}",
              f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),ROUND(D{r}*E{r}*(1-N(F{r})),2),"")',
              font=font(bold=True), number_format=FORMAT_EUR, alignment=centru)
        scrie(ws, f"H{r}", f'=IF(ISNUMBER(G{r}),ROUND(G{r}*Curs,0),"")',
              font=font(), number_format=FORMAT_MDL, alignment=centru)
        scrie(ws, f"I{r}",
              f'=IF(B{r}="","",IFERROR(INDEX({oferte_produse},MATCH(B{r},{nume_cat},0))&"",""))',
              font=font(size=9, color=GRI))
        for col in "ABCDEFGH":
            ws[f"{col}{r}"].border = chenar_tot()
    nume("OfertaProduse", f"$B${prima}:$B${ultima}")
    nume("OfertaLinkuri", f"$I${prima}:$I${ultima}")

    # Rând exemplu: primul produs din listă, o bucată.
    ws[f"B{prima}"] = produse[0]["denumire"]
    ws[f"D{prima}"] = 1

    t = ultima + 1
    merge(f"C{t}:F{t}", "TOTAL fără TVA", font=font(size=11, bold=True),
          alignment=Alignment(horizontal="right", vertical="center"))
    for col, formula, fmt in (("G", f"=SUM(G{prima}:G{ultima})", FORMAT_EUR),
                              ("H", f"=SUM(H{prima}:H{ultima})", FORMAT_MDL)):
        scrie(ws, f"{col}{t}", formula, font=font(size=11, bold=True, color="FFFFFF"),
              number_format=fmt, fill=umplere(culoare), border=chenar_tot(),
              alignment=Alignment(horizontal="center", vertical="center"))

    # Mulțumire
    p = t + 2
    ws.merge_cells(f"A{p}:H{p + 3}")
    multumire = CellRichText()
    for i, bucata in enumerate(brand["text_multumire"]):
        multumire.append(TextBlock(InlineFont(rFont=FONT_TEXT, sz=11, b=True,
                                              color="434343" if i == 1 else GRI_TEXT), bucata))
    scrie(ws, f"A{p}", multumire, alignment=Alignment(vertical="top", wrap_text=True))

    # Semnătură și ștampilă
    s = p + 5
    for i, valoare in enumerate([brand["firma"], rep.get("functie", ""), rep.get("nume", "")]):
        scrie(ws, f"C{s + i}", valoare, font=text(size=11), alignment=Alignment(vertical="center"))
    if rep.get("semnatura") and (privat / rep["semnatura"]).is_file():
        pune_imagine(ws, privat / rep["semnatura"], 5, s - 2, 190, 112, dy=0, taie=True,
                     zona_dreapta=latime_px(ws, "F") + latime_px(ws, "G") + latime_px(ws, "H"))

    # Subsol
    f = s + 5
    stanga = brand["adresa_linii"] + [brand["email"]]
    dreapta = [brand["firma"], f"C.F. {brand['cod_fiscal']}", f"TVA {brand['cod_tva']}"]
    linie = Side(style="thin", color="595959")
    for i in range(3):
        ws.row_dimensions[f + i].height = 16
        scrie(ws, f"A{f + i}", stanga[i], font=text(size=9))
        scrie(ws, f"H{f + i}", dreapta[i], font=text(size=9), alignment=Alignment(horizontal="right"))
    for col in "AB":
        ws[f"{col}{f}"].border = Border(top=linie)
        ws[f"{col}{f + 2}"].border = Border(bottom=linie)
    for col in "GH":
        ws[f"{col}{f}"].border = Border(top=linie)
        ws[f"{col}{f + 2}"].border = Border(bottom=linie)
    pune_imagine(ws, RADACINA / "brand" / "logo-ozonteh-complet.png", 2, f - 2, 90, 72, dy=0, taie=True,
                 zona_dreapta=(latime_px(ws, "C") + latime_px(ws, "D") + latime_px(ws, "E")
                               + latime_px(ws, "F") + 90) // 2 + 8)
    ultim_rand = f + 2

    ws.print_area = f"A1:H{ultim_rand}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.4


def citeste_reprezentant():
    """Datele reprezentantului stau în brand/privat/ (nu se urcă pe GitHub)."""
    cale = RADACINA / "brand" / "privat" / "reprezentant.json"
    if not cale.is_file():
        print("  atenție: lipsește brand/privat/reprezentant.json; pagina de start "
              "iese fără reprezentant, poză și ștampilă (vezi brand/reprezentant.exemplu.json)")
        return {}
    return json.loads(cale.read_text(encoding="utf-8"))


# ---------------------------------------------------------------- foaia 4

def foaia_instructiuni(wb, brand):
    """Instrucțiunile stau pe foaia lor: Google Sheets tipărește tot ce e pe o foaie."""
    ws = wb.create_sheet(FOAIE_INSTRUCTIUNI, 2)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 110
    sectiuni = [
        ("OFERTA COMERCIALĂ (pagina de start)", [
            "1. Completați câmpurile galbene: clientul („Pentru:”), numărul ofertei; după completare "
            "devin albe, ca PDF-ul să iasă curat. Data e cea de azi; valabilitatea se alege din listă.",
            "2. În coloana „Produs” apăsați săgeata și alegeți produsul (până la "
            f"{LINII_OFERTA} produse). Prețul (promo, dacă există) și linkul se completează singure "
            "din „Lista de prețuri”.",
            "3. Scrieți cantitatea și, dacă e cazul, reducerea (ex. 5%).",
            "4. Totalul în lei folosește cursul din „Lista de prețuri” (F3).",
            "5. Primul rând e un exemplu: înlocuiți-l sau ștergeți-l.",
            "6. Reprezentantul, poza și ștampila vin din brand/privat/ (le schimbă cine generează "
            "fișierul).",
        ]),
        ("PDF: DOAR OFERTA", [
            "Pe foaia „Ofertă comercială”: Fișier → Descărcare → PDF → Foaia curentă, A4, portret, "
            "„Potrivire la lățime”.",
        ]),
        ("PDF: OFERTA + FIȘELE PRODUSELOR ALESE (meniul Ozonteh, doar în Google Sheets)", [
            "Meniul „Ozonteh → Descarcă oferta completă (PDF)” face un singur PDF: oferta comercială, "
            "apoi oferta (fișa tehnică) fiecărui produs ales, în ordinea din tabel.",
            "PDF-ul se salvează și în Google Drive, în dosarul „Oferte Ozonteh (PDF)”.",
            "",
            "Instalare, o singură dată pentru fiecare fișier:",
            "  a. Fișier → Salvează ca Foi de calcul Google (dacă în titlu apare .XLSX).",
            "  b. Extensii → Apps Script. Ștergeți tot ce e acolo și lipiți codul din fișierul "
            "oferta_completa.gs. Salvați (Ctrl+S).",
            "  c. Reîncărcați foaia de calcul. Apare meniul „Ozonteh”. La prima folosire Google cere "
            "permisiunea: Continuați → contul Ozonteh → Permiteți.",
            "",
            "Când se schimbă prețurile: Fișier → Import → Încărcare → fișierul .xlsx nou → "
            "„Înlocuiți foaia de calcul”. Meniul rămâne.",
        ]),
        ("LISTA DE PREȚURI", [
            "Prețurile sunt în EUR, fără TVA. Prețul în lei = prețul final × cursul din F3.",
            "Cursul din F3 se ia live din Google Finance (GOOGLEFINANCE, doar în Google Sheets). "
            "În Excel apare cursul de rezervă din brand/ozonteh.json. Pentru un curs fix, "
            "scrieți numărul direct în F3.",
        ]),
    ]
    rand = 2
    scrie(ws, "B2", "CUM SE FOLOSEȘTE", font=font(size=16, bold=True, color=brand["culoare_principala"]))
    rand = 4
    for titlu, randuri in sectiuni:
        scrie(ws, f"B{rand}", titlu, font=font(size=11, bold=True, color=brand["culoare_principala"]))
        ws[f"B{rand}"].border = chenar_jos(brand["culoare_principala"])
        rand += 1
        for text in randuri:
            scrie(ws, f"B{rand}", text or None, font=font(),
                  alignment=Alignment(vertical="top", wrap_text=True))
            rand += 1
        rand += 1
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ---------------------------------------------------------------- main

def genereaza(folder):
    folder = Path(folder)
    brand = json.loads((RADACINA / "brand" / "ozonteh.json").read_text(encoding="utf-8"))
    config, produse = citeste_furnizor(folder)
    wb = Workbook()
    randuri_produse = foaia_lista(wb, brand, config, produse, folder)
    ultim_rand_catalog = foaia_catalog(wb, randuri_produse)
    foaia_oferta(wb, brand, produse, ultim_rand_catalog)
    foaia_instructiuni(wb, brand)
    wb.active = 0
    iesire = RADACINA / "output" / f"Lista de preturi - {config['furnizor']}.xlsx"
    iesire.parent.mkdir(exist_ok=True)
    wb.save(iesire)
    print(f"{config['furnizor']}: {len(produse)} produse -> {iesire.relative_to(RADACINA)}")
    return iesire


if __name__ == "__main__":
    if sys.argv[1:] == ["--toti"]:
        for f in sorted((RADACINA / "furnizori").iterdir()):
            if f.is_dir() and not f.name.startswith("_"):
                genereaza(f)
    elif len(sys.argv) == 2:
        genereaza(sys.argv[1])
    else:
        raise SystemExit(__doc__)
