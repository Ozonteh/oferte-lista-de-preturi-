# Pluguri: de verificat

Sursa: Google Sheets „lista de preturi pluguri”, foaia „Pluguri”, importată pe 14.10.2026.
Secțiuni: Namyslo (actualizat 20.01.26), Aydin (04.07.24), Aydin reversibile (04.07.24),
Shaktiman reversibile (04.07.24), Graecus (16.09.24). Prețurile sunt „Pret EURO”;
toate cele 17 oferte de pe Drive au exact prețul din listă.

## Nu au intrat în listă

- **Rândurile 10–52 (ascunse)**: o listă Aydin mai veche, fără prețuri
  („De actualizat”), cu pluguri reversibile de 30/35/40/45/50 cm care nu apar
  în secțiunea actuală.
- **Plugurile reversibile Shaktiman SMP-2AX și SMP-3AX** (MB Plough 2MB / 3MB): sunt
  deja în foaia „Shaktiman” (categoria „Pluguri și scarificatoare”), iar o denumire
  nu poate apărea de două ori. Prețul diferă: lista de pluguri spune **3 100 €** și
  **3 400 €**, dar foaia Shaktiman și ofertele de pe Drive spun **3 200 €** și
  **3 800 €**. Am păstrat 3 200 / 3 800 (două surse din trei). De corectat în foaia
  sursă de pluguri.
- Coloana V: ofertele „Plug Ayden 3.30” și „Plug Ayden 3.35” (lângă Aydin 3-30 și
  3-35 Clasic) nu se mai găsesc pe Drive. În listă am pus ofertele din coloana J,
  „Plug AYDIN 3-30 Clasic” și „Plug AYDIN 3-35 Clasic”.

## Fără preț („La cerere”), fără ofertă și fără poză

| Produs | În sursă |
|---|---|
| Plug Aydin 3-25 Lamelat | preț gol |
| Plug Aydin 3-30 Lamelat | preț gol |
| Plug Aydin 3-35 Lamelat | preț gol |
| Plug Aydin 4-35 Clasic | „nui” |
| Plug Aydin 4-35 Lamelat | „-” |
| Plug reversibil Aydin 5-35 Clasic | „La comanda” (trecut la Notă) |
| Plug reversibil Aydin 6-35 Clasic | „La comanda” (trecut la Notă) |

Dacă unele nu se mai vând (ex. 4-35 Clasic, „nui”), scoateți-le din `produse.csv`.

## Linkuri

- Rândurile de mai sus nu aveau link la site. Am pus pagina seriei, ca la celelalte
  modele din serie (`/product/pluguri-aydin/`, `/product/plug-reversibil-aydin/`).
- Namyslo: în celulă scrie „Pluguri Namyslo – Ozonteh”, linkul e hyperlinkul din
  spate (`/product/pluguri-namyslo/`). N-am putut deschide ozonteh.net din sesiune.

## Denumiri

- „Pluguri Reversibile Aydin”: „3 corpi x 35cm Clasic” → „Plug reversibil Aydin 3-35
  Clasic” (ca în titlul ofertei); la fel 4, 5 și 6 corpi.
- „AYDIN” scris „Aydin”, ca Namyslo și Graecus.

## Pozele

Foaia nu are poze. Pozele vin din ofertele de pe Drive (poza principală), verificate
pe planșă: numărul de trupițe se potrivește cu modelul. Ofertele din aceeași serie
au aceeași poză (Aydin 2-x, Aydin 3-x, Namyslo 3-x, Graecus DIA8i, Graecus TRA8i),
deci e câte o poză pe serie. La Aydin fundalul transparent a fost pus pe alb.

## Sigle

Foaia și ofertele nu au siglele Namyslo, Aydin și Graecus. Pentru siglă pe rândul
fiecărei categorii, puneți fișierele în acest folder și adăugați
`logouri_categorii` în `furnizor.json` (ca la Tractoare).

## Alte observații din sursă

- Graecus 3-25 TRA8i: puterea minimă e scrisă „35-35” CP (nu intră în listă).
- Foaia sursă are și coloane „SKU” și „Pret Vinzare” (Q–V), goale.
