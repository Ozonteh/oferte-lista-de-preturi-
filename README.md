# Liste de prețuri și oferte Ozonteh

Un singur model de listă de prețuri pentru toți furnizorii. Fiecare furnizor are
un folder cu datele lui, iar scriptul face din ele fișierul Excel final.

## Ce conține fișierul Excel

`output/Lista de preturi - <Furnizor>.xlsx`, cu două foi vizibile:

1. **Lista de prețuri**: produsele grupate pe categorii, fiecare cu:
   Nr. · Foto · Denumire produs · Preț EUR (fără TVA) · Preț promo EUR ·
   Preț MDL · Notă · link „Vezi pe site” · link „Deschide oferta”.
   Cursul EUR/MDL e în celula **F3**, iar prețul în lei se recalculează singur.
2. **Ofertă comercială**: ofertă pentru client, de o pagină A4. Produsul se
   alege din listă, iar prețul, linkul și totalurile (fără TVA, TVA, cu TVA, în
   EUR și MDL) se completează singure. Câmpurile galbene sunt de completat și
   devin albe după completare, ca PDF-ul pentru client să iasă curat.

A treia foaie, **Catalog**, e ascunsă: din ea citește oferta. Nu se editează.

## Structura

```
brand/                  sigla Ozonteh, QR-ul site-ului, datele firmei (ozonteh.json)
furnizori/
  _sablon/              punct de pornire pentru un furnizor nou
  shaktiman/            un furnizor = un fișier Excel
  tractoare/            listă cu mai multe mărci (LS, Farmtrac, Carraro, Solis)
    furnizor.json       nume, titlu, sigla, data actualizării
    produse.csv         lista de produse (sursa de adevăr)
    imagini/            pozele produselor
    logo.png            sigla furnizorului
    VERIFICARI.md       ce s-a corectat la import și ce rămâne de verificat
output/                 fișierele Excel generate
scripts/
  genereaza_lista.py    produse.csv -> Excel
  importa_lista_veche.py  lista veche din Google Sheets -> produse.csv + imagini
```

## `produse.csv`

| coloană | ce conține |
|---|---|
| `categorie` | ex. „Tocătoare de resturi vegetale”. Ordinea categoriilor e ordinea din fișier. |
| `denumire` | `<Tip> <Marcă> <Model>`, cu diacritice. Unică: oferta caută prețul după denumire. |
| `link_produs` | pagina produsului pe ozonteh.net |
| `pret_eur` | preț de listă în EUR, fără TVA. Gol = „La cerere”. |
| `pret_promo_eur` | opțional; dacă există, el intră în oferta comercială |
| `nota` | opțional, text scurt (ex. „Cu 2 burghie: 2 300 €”) |
| `link_oferta` | opțional, oferta/fișa tehnică pe Google Drive |
| `imagine` | numele fișierului din `imagini/` |

## Furnizor nou, pas cu pas

1. Copiați `furnizori/_sablon` în `furnizori/<nume-furnizor>` și completați `furnizor.json`.
2. Dacă furnizorul are deja o listă în formatul vechi (Denumire / Poze / Link la sait /
   Descriere / Preț sau Pret Vinzare / Preț Promo / Oferta), descărcați-o ca .xlsx și
   importați-o. Merge și cu foi care au mai multe secțiuni (ex. câte una pe marcă).
   Coloanele „Achiziție” și „SKU” se ignoră.
   ```
   python scripts/importa_lista_veche.py lista.xlsx furnizori/<nume-furnizor>
   ```
   Altfel completați `produse.csv` de mână și puneți pozele în `imagini/`.
3. Completați categoriile și corectați denumirile în `produse.csv`.
4. Generați fișierul:
   ```
   python scripts/genereaza_lista.py furnizori/<nume-furnizor>
   ```
   Scriptul oprește generarea la erori (denumiri dublate, prețuri nenumerice) și
   afișează atenționări (produse fără preț, fără link sau fără poză).

Pentru toți furnizorii odată: `python scripts/genereaza_lista.py --toti`.

## În Google Sheets

Fișier → Import → Încărcare → alegeți fișierul .xlsx. Pozele, linkurile, lista
de produse din ofertă și formulele rămân. Cursul din F3 se poate lega automat
la Google Finance cu `=GOOGLEFINANCE("CURRENCY:EURMDL")`.

## `furnizor.json`

| câmp | ce conține |
|---|---|
| `furnizor` | numele din numele fișierului: `Lista de preturi - <furnizor>.xlsx` |
| `titlu` | rândul de sub „LISTA DE PREȚURI” |
| `logo` | sigla din dreapta sus (opțional) |
| `logouri_categorii` | opțional: sigla afișată pe rândul fiecărei categorii, pentru liste cu mai multe mărci |
| `data_actualizare` | data din antet, `AAAA-LL-ZZ` |

## Setări comune

`brand/ozonteh.json`: datele firmei, culoarea, cursul implicit, cota TVA,
valabilitatea ofertei și condițiile din ofertă (garanție, livrare, finanțare,
service). Se schimbă o dată aici și apar la toți furnizorii la următoarea generare.

Instalare: `pip install -r requirements.txt`
