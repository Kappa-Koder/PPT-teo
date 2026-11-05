# Vi presentiamo Giove

Presentazione animata su Giove per la ricerca di Razvan, Sofia, Riccardo e Idir.
15 slide con illustrazioni in movimento: il pianeta che ruota, le lune che girano, la Grande Macchia Rossa che vortica, 1.321 Terre che riempiono Giove, la sonda Juno in orbita.

## Come si usa

1. Apri **`index.html`** con Chrome, Edge o Firefox (doppio clic sul file). Funziona anche **senza internet**.
2. Premi **F** (o il pulsante "Schermo intero") per proiettarla a tutto schermo.
3. Comandi:
   - **→**, **barra spaziatrice**, **PagGiù** o un clic: slide successiva. Funziona anche il telecomando da presentazione.
   - **←** o **PagSu**: slide precedente.
   - **N**: mostra le **note per chi parla** (il testo completo della ricerca, slide per slide).
   - Sulla LIM o sul tablet si cambia slide scorrendo con il dito.
4. Slide interattive:
   - **04 Simile a una stella**: tocca × 1, × 13, × 75 per trasformare Giove in nana bruna o in stella.
   - **05 Le dimensioni**: "Ripeti l'animazione" fa ripartire le Terre.
   - **14 Quiz lampo**: tocca una carta per girarla e vedere la risposta.

## Versione PowerPoint con testi modificabili

**`giove-animato.pptx`** è la stessa presentazione per PowerPoint 2016 o 365:

- le animazioni (Giove, lune, Macchia Rossa, Terre, anelli, Juno) sono video di sfondo che partono da soli;
- tutti i testi sono caselle di PowerPoint: si correggono come in qualsiasi presentazione;
- i testi compaiono in dissolvenza; nel quiz ogni clic fa comparire una risposta;
- le note per chi parla sono sotto ogni slide.

Cosa cambia rispetto all'HTML:

- il carattere è Calibri, presente con Office;
- alcune scritte fanno parte del video e non si modificano: il contatore "× 1 / × 13 / × 75" e lo stato "Giove / Nana bruna / Stella", i numeri 11 e 1.321 che contano le Terre, i nomi delle lune in orbita, "Juno", "Terra" e le misure in km sulla Macchia Rossa;
- i video delle slide 05, 06 e 09 (Terre, interno di Giove, Macchia Rossa) si fermano sull'ultima immagine; gli altri ripartono da soli;
- se un testo diventa più lungo, la casella non si allarga da sola sopra i riquadri disegnati nel video.

Per rigenerarla dopo una modifica a `src/giove.html`: `python3 build.py`, poi
`node tools/ppt/record.js index.html <cartella>` e `python3 tools/ppt/build_pptx.py <cartella> giove-animato.pptx`
(servono Playwright, ffmpeg e python-pptx).

## Versione di riserva (senza animazioni)

Nella cartella `riserva/` ci sono le stesse 15 slide, ferme, con le note per chi parla:

- `giove-statico.pptx` per PowerPoint;
- `giove-statico.pdf` da aprire con qualsiasi lettore PDF.

Usale se il computer della scuola è molto vecchio o non apre i file HTML.

## Cosa è stato corretto rispetto ai materiali raccolti

| Nel fascicolo | Nella presentazione | Perché |
|---|---|---|
| Grande Macchia Rossa "visibile da circa quattro secoli" | "La osserviamo da almeno 190 anni" | Uno studio del 2024 (Sánchez-Lavega e altri, *Geophysical Research Letters*) ha mostrato che la macchia vista da Cassini nel 1665 era un'altra tempesta. |
| Anello "spesso 4 km", formato da "collisioni fra meteoriti" | Polvere che nasce quando piccoli meteoriti colpiscono le lune vicine; anello principale a più di 50.000 km dalle nubi | L'anello principale è tra circa 51.000 e 57.000 km sopra le nubi e spesso decine di km, non 4. |
| Nucleo "si ritiene che sia solido" | Nucleo denso, che secondo i dati della sonda Juno è probabilmente "sfumato" | I dati di gravità di Juno indicano un nucleo diluito, mescolato con l'idrogeno. |
| Campo magnetico "10 volte" e "polo 20.000 volte più forte" | "20.000 volte" per il magnete nel complesso, "più di 10 volte" vicino alle nubi | Il numero 20.000 riguarda la forza complessiva del magnete (momento magnetico), non il polo. |
| Temperatura e gravità "superficiali" | "In cima alle nubi" | Giove non ha una superficie solida (lo dice la stessa ricerca). |
| Foto di copertina con un anello largo e brillante | Giove disegnato senza anello vistoso | Quello è l'aspetto di Saturno: gli anelli di Giove sono quasi invisibili. |

Il conteggio "più di 100 lune" è corretto: nel marzo 2026 l'Unione Astronomica Internazionale ne riconosceva 101, e altre scoperte di aprile 2026 hanno portato il totale a circa 115.

## Per chi vuole modificarla

Il file da modificare è `src/giove.html`. Dopo ogni modifica esegui `python3 build.py`, che rigenera `index.html` con i caratteri incorporati (Bricolage Grotesque e Atkinson Hyperlegible, licenza SIL Open Font License).
Giove, le lune e le animazioni sono disegnati dal codice (WebGL e Canvas): non ci sono immagini protette da diritti d'autore.
