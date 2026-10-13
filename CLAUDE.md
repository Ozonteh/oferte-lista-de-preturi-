# Reguli pentru listele de prețuri Ozonteh

Repo-ul ține **un singur model** de listă de prețuri + ofertă comercială, aplicat la
toți furnizorii. Citește `README.md` pentru structură și comenzi.

## Când utilizatorul cere lista pentru un furnizor nou

1. Ia sursa: link Google Sheets (descarcă .xlsx prin Google Drive,
   `exportMimeType` = xlsx), PDF sau listă de la furnizor.
2. Format vechi (Denumire / Poze / Link la sait / Preț / Preț Promo / Oferta):
   `scripts/importa_lista_veche.py`. Altfel construiește `produse.csv` direct.
3. Curăță datele în `furnizori/<furnizor>/produse.csv`:
   - denumire `<Tip> <Marcă> <Model>`, cu diacritice, unică;
   - categorii scurte, la plural, aceleași ca la ceilalți furnizori acolo unde se
     potrivesc (ex. „Tocătoare de resturi vegetale”, „Freze de sol”, „Prese de balotat”);
   - pozele numite după serie (`bmf.jpg`, `srp.jpg`).
   - foaie fără poze de produs: ia poza principală din oferta fiecărui produs
     (Google Drive, export xlsx), verifică pe o planșă că modelul se potrivește și
     pune-o pe fundal alb (pozele PNG transparente devin altfel negre);
   - listă cu mai multe mărci: o categorie pe marcă („Tractoare LS”) și siglele în
     `logouri_categorii` din `furnizor.json`.
4. `python scripts/genereaza_lista.py --toti`: **un singur fișier**,
   `output/Lista de preturi Ozonteh.xlsx`, cu oferta, instrucțiunile și câte o foaie
   pe furnizor. Furnizorul nou devine o foaie nouă acolo, nu un fișier separat.
5. Verifică: recalculează o **copie** cu `recalc.py` din skill-ul xlsx (0 erori),
   randează în PDF cu LibreOffice și uită-te la pagini. Nu recalcula fișierul din
   `output/` direct: LibreOffice îl rescrie.
6. Commit pe branch-ul de lucru, cu `produse.csv` și `imagini/`. Fișierul din `output/`
   nu se urcă (conține ștampila; repo-ul e public): îl trimiți utilizatorului direct.

## Ce nu se face

- **Nu inventa** prețuri, modele sau linkuri. Ce lipsește rămâne gol („La cerere”,
  „—”) și se raportează utilizatorului.
- **Prețurile de achiziție, codurile de fabrică și datele de container nu intră
  niciodată în listă.** Listele vechi le au uneori mai jos în foaie; importul se
  oprește după 20 de rânduri goale tocmai ca să nu le prindă. Verifică.
- Linkuri: când textul din celulă e o adresă diferită de hyperlink, importul ia
  textul (hyperlinkul e adesea vechi sau o previzualizare de ciornă). Treci fiecare
  diferență în `VERIFICARI.md`.
- Nu „repara” în tăcere neconcordanțele dintre listă, oferte și site (model diferit,
  link la alt produs). Alege varianta susținută de cele mai multe surse, scrie-o în
  `furnizori/<furnizor>/VERIFICARI.md` și spune-i utilizatorului.
- Nu schimba modelul (coloane, foi, culori) doar pentru un furnizor. O schimbare
  de model se face în `scripts/genereaza_lista.py` și se regenerează toți (`--toti`).

## Detalii de model de păstrat

- Ordinea foilor: Ofertă comercială, Instrucțiuni, foile furnizorilor (numele din
  `furnizor.json`), Catalog (ascunsă, toate produsele). Denumiri unice între furnizori.
- Prețurile sunt în EUR, fără TVA. Lei = preț final (promo dacă există) × curs (F3).
  Cursul live e pe prima foaie de furnizor (numele `Curs`); celelalte au `=Curs`.
- F3 = `IFERROR(GOOGLEFINANCE("CURRENCY:EURMDL"), <curs_eur_mdl>)`: live în Google
  Sheets, curs de rezervă în Excel/LibreOffice (unde GOOGLEFINANCE dă #NAME?).
- Foaia „Catalog” e ascunsă; oferta citește din ea prin INDEX/MATCH după denumire.
- Câmpurile de completat din ofertă sunt galbene doar cât sunt goale (formatare
  condițională), ca PDF-ul pentru client să iasă curat.
- Font Arial, culoarea principală din `brand/ozonteh.json` (`00497E`, din siglă).
- Google Sheets ignoră zona de tipar din Excel și tipărește tot ce e pe foaie:
  pe foile care se tipăresc nu pune nimic în afara tabelului (instrucțiunile stau
  pe foaia „Instrucțiuni”; datele ajutătoare în coloane ascunse).
- „Ofertă comercială” e pagina de start după designul Ozonteh (titlu League Spartan
  42 `#053B57`, text Open Sans gri). `oferta_completa.gs` citește câmpurile prin
  nume (`OfertaProduse`, `OfertaLinkuri`, `OfertaClient`, `OfertaNr`, `OfertaData`);
  dacă muți tabelul, numele se mută odată cu el.
- Poza din tabelul ofertei = `IMAGE()` cu adresa raw GitHub a pozei (Catalog, coloana H,
  `url_poze` din `brand/ozonteh.json`, legată de commitul pozei). Pozele noi: commit +
  push, apoi generează din nou; scriptul avertizează dacă o poză nu e încă urcată.
- **Repo public**: semnătura, ștampila, poza și datele reprezentantului stau doar în
  `brand/privat/` (ignorat de git). Nu le urca niciodată și nu le pune în alt dosar.
- Google Sheets nu are indentare: textul de pe benzi primește spații în față.
