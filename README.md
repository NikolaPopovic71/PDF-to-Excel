# PDF ➜ Excel (rezultati plivanja)

Pretvara izveštaj „Medalists by event“ (Splash Meet Manager) u Excel:
naslov discipline preko svih kolona, pa kolone **R.br. | Ime i prezime | Godište | Klub | Vreme**
(godište i vreme su formatirani kao tekst).

## Pravljenje .exe fajla

**Opcija A — bez Pythona (GitHub):** postavite folder u GitHub repozitorijum.
Workflow `.github/workflows/build.yml` automatski napravi `PdfUExcel.exe`
(Actions → poslednji run → Artifacts → PdfUExcel).

**Opcija B — lokalno na Windowsu sa Pythonom:** dvoklik na `build.bat`.
Rezultat: `dist\PdfUExcel.exe`.

## Korišćenje
Pokrenite `PdfUExcel.exe` → *Izaberi PDF…* → *Konvertuj u Excel* → izaberite gde da sačuva.
Na drugom laptopu nije potrebno ništa instalirati — samo kopirajte .exe.

> Windows SmartScreen može prikazati upozorenje za nepotpisan .exe: *More info → Run anyway*.
