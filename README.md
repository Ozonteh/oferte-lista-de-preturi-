# Liste de prețuri și oferte Ozonteh

Un singur model de listă de prețuri pentru toți furnizorii. Fiecare furnizor are
un folder cu datele lui, iar scriptul face din ele fișierul Excel final.

## Ce conține fișierul Excel

Un singur fișier pentru toți furnizorii: `output/Lista de preturi Ozonteh.xlsx`
(`python scripts/genereaza_lista.py --toti`). Foile, în ordine: „Ofertă comercială”,
„Instrucțiuni”, câte o foaie pe furnizor („Shaktiman”, „Tractoare”, …) și
„Catalog” (ascunsă). Un furnizor nou = o foaie nouă în același fișier.

1. **Foaia fiecărui furnizor** (lista de prețuri): produsele grupate pe categorii, fiecare cu:
   Nr. · Foto · Denumire produs · Preț EUR (fără TVA) · Preț promo EUR ·
   Preț MDL · Notă · link „Vezi pe site” · link „Deschide oferta”.
   Cursul EUR/MDL e în celula **F3** a primei foi de furnizor: în Google Sheets vine
   live din Google Finance (`GOOGLEFINANCE("CURRENCY:EURMDL")`), în Excel e cursul de
   rezervă din `brand/ozonteh.json`. Celelalte foi îl preiau (`=Curs`). Prețul în lei
   se recalculează singur.
2. **Ofertă comercială**: pagina de start a ofertei, după designul Ozonteh
   („Oferta KVINT”): titlu, „Pentru: client”, dată, număr, valabilitate,
   reprezentantul cu poză și contacte, tabelul de produse (până la 6, cu poza
   produsului ales), textul de mulțumire, semnătura cu ștampila. Produsul se
   alege din listă (produsele tuturor furnizorilor); poza, prețul, linkul și totalul fără TVA (EUR și MDL) se
   completează singure. Câmpurile galbene sunt de completat și devin albe după
   completare.

   Poza din ofertă se încarcă în Google Sheets cu `IMAGE()` de pe GitHub (adresa e
   în „Catalog”, coloana H, legată de commitul pozei). De aceea: o poză nouă se urcă
   (commit + push) **înainte** de generare, iar dacă repo-ul devine privat, pozele din
   ofertă nu se mai văd (cele din „Lista de prețuri” rămân, sunt în fișier).

3. **Instrucțiuni**: cum se completează oferta și cum se scoate PDF-ul.

**Catalog** e ascunsă: din ea citește oferta. Nu se editează.

## Oferta completă într-un PDF (Google Sheets)

Meniul **Ozonteh** are două comenzi: **Descarcă oferta completă (PDF)** și
**Tipărește oferta completă**. Amândouă fac un singur PDF: pagina de start, apoi
oferta (fișa tehnică) fiecărui produs ales, luată din linkul „Ofertă comercială”
al produsului. PDF-ul se salvează și în Drive, în dosarul „Oferte Ozonteh (PDF)”.
Codul e în `scripts/apps_script/oferta_completa.gs`.

Instalare, o dată pe fiecare foaie de calcul:
1. Fișier → Salvează ca Foi de calcul Google (dacă titlul are .XLSX).
2. Extensii → Apps Script → lipiți tot `oferta_completa.gs` → Salvați.
3. Reîncărcați foaia: apare meniul „Ozonteh”. La prima folosire Google cere permisiunea.

Prețuri noi: Fișier → Import → fișierul .xlsx nou → „Înlocuiți foaia de calcul”.
Meniul rămâne.

## Structura

```
brand/                  sigla Ozonteh, QR-ul site-ului, datele firmei (ozonteh.json)
furnizori/
  _sablon/              punct de pornire pentru un furnizor nou
  adrimak/              scarificatoare, grape cu discuri, distribuitoare Adrimak
  agroremmasch/         grape cu discuri AGD și cultivatoare KSO Agroremmasch
  shaktiman/            un furnizor = un fișier Excel
  mcms-warka/           cositori, platforme și utilaje de plantare MCMS Warka
  pluguri/              pluguri Namyslo, Aydin, Graecus (o categorie pe marcă)
  tractoare/            listă cu mai multe mărci (LS, Farmtrac, Carraro, Solis, Lansu)
    furnizor.json       nume, titlu, sigla, data actualizării
    produse.csv         lista de produse (sursa de adevăr)
    imagini/            pozele produselor
    logo.png            sigla furnizorului
    VERIFICARI.md       ce s-a corectat la import și ce rămâne de verificat
output/                 fișierele Excel generate
scripts/
  genereaza_lista.py    produse.csv -> Excel
  importa_lista_veche.py  lista veche din Google Sheets -> produse.csv + imagini
  apps_script/oferta_completa.gs  meniul „Ozonteh” din Google Sheets
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

Fișierul final, cu toți furnizorii: `python scripts/genereaza_lista.py --toti`.
Denumirile trebuie să fie unice și între furnizori (scriptul se oprește altfel).

## În Google Sheets

Fișier → Import → Încărcare → alegeți fișierul .xlsx. Pozele, linkurile, lista
de produse din ofertă și formulele rămân. Cursul din F3 se actualizează singur.

## `furnizor.json`

| câmp | ce conține |
|---|---|
| `furnizor` | numele foii din fișierul comun (ex. „Shaktiman”) |
| `titlu` | rândul de sub „LISTA DE PREȚURI” |
| `logo` | sigla din dreapta sus (opțional) |
| `logouri_categorii` | opțional: sigla afișată pe rândul fiecărei categorii, pentru liste cu mai multe mărci |
| `data_actualizare` | data din antet, `AAAA-LL-ZZ` |

## Date private (nu se urcă pe GitHub)

Repo-ul e **public**. Poza și datele reprezentantului, semnătura și ștampila stau în
`brand/privat/` (ignorat de git; model de completat: `brand/reprezentant.exemplu.json`),
iar fișierele generate din `output/` (care le conțin) nu se urcă: se trimit direct.
Fără `brand/privat/`, pagina de start iese fără reprezentant și ștampilă.

## Setări comune

`brand/ozonteh.json`: datele firmei, culorile, cursul de rezervă (pentru Excel),
valabilitatea implicită a ofertei și variantele din listă, textul de mulțumire.
Se schimbă o dată aici și apar la toți furnizorii la următoarea generare.

Instalare: `pip install -r requirements.txt`
