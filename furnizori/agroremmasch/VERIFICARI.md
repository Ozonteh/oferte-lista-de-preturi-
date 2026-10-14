# Agroremmasch: de verificat

Sursa: Google Sheets „Agroremmasch”, foaia „Agroremmasch”, importată pe 14.10.2026.
Prețurile sunt „Pret Vinzare” din tabelul de sus (rândurile 2–15), în EUR fără TVA
(așa scriu ofertele). Regula: se folosește prețul din listă.

## Al doilea tabel din foaie (rândurile 38–53): alte prețuri

Mai jos în foaie sunt un tabel cu caracteristici („Pret EURO”) și câteva rânduri
separate. Unele prețuri diferă de lista de sus:

| Produs | Lista de sus (pus) | Tabelul de jos |
|---|---|---|
| AGD-1,3 | 1 900 € | 1 800 € |
| AGD-2,5 | 3 900 € | 3 550 € |
| KSO 1.9 („CSO 1,9”) | 1 750 € | 1 700 € |
| KSO 2.5 („CSO 2,5”) | 1 850 € | 1 750 € |
| KSON 3.5 („CSON- 3,5”) | 3 000 € | 2 900 € |

Ofertele găsite (KSO 1.9, KSO 2.5, KSON 3.5, AGD-1,8 L, AGD-2,5 H) au prețul din
lista de sus. **AGD-2,1 X** (3 600 €) apare doar în tabelul de jos, fără link și fără
ofertă: nu l-am pus în listă. AGD-1,8 L are acolo „3000 (cu discuri)”.

## Oferte care nu se deschid

Ofertele AGD-1,0, AGD-1,3, AGD-1,6, AGD-1,8, AGD-2,1 și AGD-2,5 nu se pot deschide din
contul marketing@ozonteh.md și nu le-am găsit pe Drive sub alt nume. Le-am lăsat în
listă; verificați-le.

## Ofertele: suma în lei e greșită

AGD-1,8 L, KSO 1.9, KSO 2.5 și KSON 3.5 au în celula „Lei” numărul fix **108 000**, nu
formula preț × curs (aceeași greșeală ca la Adrimak).

## Linkuri

- Grapele AGD nu au link la site în listă (celula scrie „link”, fără adresă). Am pus
  pagina seriei din oferta AGD-1,8 L: `/product/grapa-cu-discuri-agroremmash-agd/`.
- KSO și KSON: linkurile din listă, aceleași ca în oferte.

## Denumiri

- „KSOH 3.5” → **KSON 3.5**: așa scriu oferta („KSON-3,5”), pagina de pe site
  (`…-kson-35/`) și tabelul de jos („CSON”).
- Marca scrisă „Agroremmasch” (titlul foii, ofertele KSO, site-ul); oferta AGD-1,8 L și
  pagina AGD scriu „Agroremmash”.

## Pozele

Foaia are 7 poze AGD într-un rând, fără legătură cu rândurile produselor.
- AGD-2,1: poza din foaie cu eticheta „АГД-2,1”.
- AGD-1,8 L, AGD-2,5 H, KSO, KSON: poza principală din ofertă.
- AGD-1,0, 1,3, 1,6, 1,8 și 2,5: aceeași poză de serie (`agd.jpg`, prima din foaie),
  pentru că ofertele lor nu se deschid. Înlocuiți-o cu poza fiecărui model, dacă o aveți.

## Sigla

Foaia și ofertele nu au sigla Agroremmasch. Puneți-o ca `logo.png` și completați
`"logo": "logo.png"` în `furnizor.json`.
