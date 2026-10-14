# Adrimak: de verificat

Sursa: Google Sheets „Adrimak”, foaia „Adrimak” (rândurile 2–48), importată pe 14.10.2026.
Prețurile sunt „Pret Vinzare” (coloana D), în EUR fără TVA (așa scriu toate ofertele).

## Date interne în foaia sursă

Coloanele G–I au **prețurile de achiziție**, codurile furnizorului și un preț calculat
(„+10% apoi +25%”). Nu au intrat în listă. Foaia are link de partajare: nu o trimiteți
clienților.

## Prețuri: listă ≠ ofertă

| Produs | În listă (pus) | În oferta de pe Drive |
|---|---|---|
| PIYANO 16 | 2 800 € | 2 600 € |
| PIYANO 20 | 3 400 € | 3 200 € |
| PIYANO 24 | 3 800 € | 3 400 € |
| LTD-12 | 2 400 € | 2 200 € (în listă: „2200 doar modelul din stok”) |

Am pus prețul din listă. Celelalte 15 oferte găsite au exact prețul din listă.
LTD-12: nota „Din stoc: 2 200 €; rolă opțională +600 €” (din listă: „OPTIONAL:
Roller + 600 euro”).

## Ofertele: suma în lei e greșită

16 oferte au în celula „Lei” numărul fix **108 000**, nu formula preț × curs:
Pinokyo PNK 5 și 7, KVC 5 și 7, ATOM 14 și 16, KBO 20/24/28/32/40, PIYANO 16/20/24,
LTD, ALİNDA. Clientul vede 108 000 lei oricare ar fi prețul. De corectat în ofertele
de pe Drive (formula din TTC 5M, HGM și HTG e bună).

## Oferte care nu se deschid

Linkurile de ofertă de mai jos nu se pot deschide din contul marketing@ozonteh.md
(fișiere șterse sau nepartajate) și nu le-am găsit pe Drive sub alt nume. Le-am lăsat
în listă; verificați-le:

- TTC 7M (pus la TTC-7), TTC 9M (pus la TTC-9);
- BLG-16, BLG-20, BLG-24, BLG-28, BLG-32;
- CBG 1.5-2.0, CBG 1.7, CBG 2-2.5, CBG 2.5-3.0.

Fără ofertă în listă: TTC-7M, TTC-9M, TTC-11, TTC-11M.

## Linkuri la site

| Produs | În listă | Pus în listă | De ce |
|---|---|---|---|
| Pinokyo PNK 5, 7 | pagina TTC | `/product/scarificator-adrimak-pinokyo-pnk/` | așa scrie în ambele oferte; pagina TTC e alt produs |
| KBO 20–40 | `/product/grapa-cu-discuri-adrimak-kb/` | `/product/grapa-cu-discuri-adrimak-kbo/` | așa scrie în toate cele 5 oferte |

**PIYANO** și **LTD** au, și în listă și în oferte, o previzualizare de ciornă
(`?post_type=product&p=10890&preview=true`, `p=10900`). Le-am lăsat: funcționează doar
dacă pagina e publicată. N-am putut deschide ozonteh.net din sesiune; verificați.

## Denumiri

- **TTC**: ofertele „TTC 5M/7M/9M” sunt modelele simple (TTC-5, TTC-7, TTC-9, aceleași
  prețuri). În listă „M” (TTC-7M, 9M, 11M) înseamnă varianta cu rolă dublă: am pus
  nota „Cu rolă dublă”.
- **ATOM 140 / 160**: ofertele scriu „ATOM 14 / 16”. Am păstrat denumirea din listă.
- **CBG 2,50-3,00 m**: în listă scrie „2.40 - 3.00m”, dar linkul de ofertă („CBG 2.5-
  3.00”) și modelul furnizorului spun 2,5–3 m. Am pus 2,50–3,00.
- KVC: „(55 CM)” din denumire → nota „Adâncime de lucru până la 55 cm” (așa scrie oferta).
- „PİYANO”, „ALİNDA” scrise cu I simplu; „LTD-12 51” → „LTD-12”.

## Pozele

Foaia nu are poze. Pozele vin din ofertele de pe Drive, verificate pe planșă.
Câte o poză pe serie unde ofertele au aceeași poză (ATOM, KBO, PIYANO). TTC: o singură
ofertă (TTC 5M), deci toate cele 7 TTC au poza ei (5 ancore). **BLG și CBG nu au poză**
(nicio ofertă nu se deschide).

## Sigla

Ofertele și foaia nu au sigla Adrimak. Puneți-o ca `logo.png` și completați
`"logo": "logo.png"` în `furnizor.json`.
