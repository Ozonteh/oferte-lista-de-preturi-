# MCMS Warka: de verificat

Sursa: Google Sheets „MCMS Warka”, foaia „MCMS Warka”, importată pe 14.10.2026.
3 secțiuni: cositori RG, platforme pentru recoltare PSN, utilaje pentru plantare.

## TVA: de confirmat

Foaia sursă scrie doar „Pret”. Cele 8 oferte de pe Drive scriu **„Euro cu TVA”** și
**„Lei cu TVA”**, cu aceeași sumă ca în listă. Lista Ozonteh e în EUR **fără TVA**,
iar ofertele celorlalți furnizori (pluguri, tractoare) scriu „fără TVA”.
Am pus prețurile din foaie așa cum sunt. Dacă ele conțin TVA, trebuie scăzut
TVA-ul în `produse.csv` (sau corectată eticheta din oferte, dacă ea e greșită).

## Prețuri scrise ca text în sursă

„12.500 Euro” → 12 500, „15.200 Euro” → 15 200, „3300 Euro” → 3 300,
„4900 Euro” → 4 900. Toate cele 8 oferte găsite au exact prețul din listă.

## Note

- Cositorile RG: „Cardan 6x6 inclus in pret” (sub tabel) → nota „Cardan 6×6 inclus”.
- „Cardan T6 6x6 inclus” (sub secțiunea de plantare) → doar la freza FSS-35, care
  merge de la priza de putere. Plantatorul GZ/M + SD se prinde la subsolier și nu
  are antrenare (așa scrie în oferta lui).
- RG-200/2: „(cu 2 rotoare)” din coloana Model → nota „Cu 2 rotoare”.

## Oferta lipsă

**Platformă PSN-260/2 cu role și sistem de descărcare** (15 200 €): linkul de ofertă
din foaie („MCMS warka PSN260/2”) nu se poate deschide din contul
marketing@ozonteh.md (fișier șters sau nepartajat). L-am lăsat în listă; verificați
că se deschide. Poza e cea a platformei PSN-260/2 (același model).

## Denumiri

- Denumirile sunt titlurile ofertelor, cu modelul scris ca în listă (RG-200, nu RG 200).
- „Platforma cu 2 axe cu rollere si sistem de descarcare” → „Platformă hidraulică MCMS
  Warka PSN-260/2 cu role și sistem de descărcare”.

## Pozele

Pozele din foaie au filigranul altui dealer (AgroHor) și nu stau pe rândurile
produselor, așa că nu le-am folosit. Pozele vin din ofertele de pe Drive (poza
principală), verificate pe planșă. RG-200/2, 220, 240 și 270 au aceeași poză în
oferte (`rg.jpg`); RG-200, cu un singur rotor, are alta (`rg-200.jpg`).

## Sigla

Foaia și ofertele nu au sigla MCMS Warka. Puneți-o ca `logo.png` în acest folder și
completați `"logo": "logo.png"` în `furnizor.json`.

## Alte observații din sursă

Caracteristicile tehnice nu intră în listă, dar diferă între listă și ofertă la RG-200:
lista spune 40 CP și 500 kg, oferta 25 kW (≈ 34 CP) și 510 kg.
