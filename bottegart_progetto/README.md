# Come portare il lavoro su Claude Code

## 1. Cosa trovi in questa cartella
| Cosa | Dove |
|---|---|
| Istruzioni per Claude Code | `CLAUDE.md` |
| Script che costruiscono Excel e flipbook | `build_xlsx.py`, `make_flipbook.py`, `flip_template.html` |
| Dati estratti dai cataloghi, per editore | `*_data.py`, `uo_cat.py`, `*.json` |
| Copertine già estratte | `covers.json` (base64) |
| Loghi e locandine della copertina | `assets/` |
| Risultati attuali | `output/` (flipbook HTML, Excel catalogo, Excel "da completare") |
| Cataloghi sorgente | `cataloghi/` (qui ci sono solo Excel Logos e immagini; vedi punto 3) |

## 2. Preparare il computer (una volta sola)
1. Installa Python 3.10+ e Claude Code (`npm install -g @anthropic-ai/claude-code`, serve Node.js; oppure dall'app desktop).
2. Nella cartella del progetto: `pip install -r requirements.txt`.
3. (Facoltativo) installa LibreOffice per ricalcolare le formule dell'Excel da terminale; altrimenti si ricalcolano aprendo il file.

## 3. Rimettere i cataloghi PDF
Per motivi di dimensione non li ho inclusi. Copia in `cataloghi/` questi file (con questi nomi):
Catalogo_TDM_generale_1_SEM_2026.pdf, minibombo_catalogo_2026.pdf, Catalogo-Pelledoca-2025.pdf,
CATALOGO-PRIMAVERA-2026-ridotto.pdf (Corsare), catalogo-autunno-2026_a5_layout.pdf (L'Ippocampo),
catalogo-camelozampa-2025.pdf, uovonero_catalogo_librai.pdf, Catalog_BeccoGiallo_Fall-2024.pdf,
ioleggoperche-2026locandinagemellaggia4.pdf, rebranding_Bottegart_adesivo.pdf.
Servono solo se vuoi rifare le estrazioni: i dati già estratti sono nei file `.py`/`.json`.

## 4. Aprire il progetto in Claude Code
```
cd percorso/bottegart_progetto
claude
```
Claude Code legge da solo `CLAUDE.md`. Come primo messaggio puoi scrivere:
> Leggi CLAUDE.md. Esegui `python build_xlsx.py` e `python make_flipbook.py` per verificare che tutto funzioni, poi dimmi quanti libri hanno sinossi, prezzo, età e copertina per ogni casa editrice.

## 5. Il ciclo di lavoro per ogni editore (uno alla volta)
1. Metti il catalogo (PDF/Excel/ZIP di copertine) in `cataloghi/`.
2. Scrivi a Claude Code: "Ecco il catalogo di <editore> in `cataloghi/…`. Completa solo i titoli già in selezione: sinossi, prezzo, età, ISBN, copertina."
3. Claude Code crea o aggiorna un `<editore>_data.py`, lo collega a `build_xlsx.py` e `make_flipbook.py`, estrae le copertine in `covers.json` e controlla a vista un campione.
4. Rilancia `python build_xlsx.py` e `python make_flipbook.py`.
5. Apri `output/flipbook_bottegart.html` nel browser e controlla (pagine, schede, scheda desiderata).
Per le età e i testi puoi anche modificare a mano il blocco "DATI DEI LIBRI" nell'HTML (istruzioni in testa al file).

## 6. Cosa resta da fare (backlog)
Apri `output/da_completare_per_editore.xlsx`: ha un foglio riepilogo e una riga per ogni dato mancante. In sintesi:
- Sinossi: Logos (10), BeccoGiallo (14), Camelozampa (7), La Meridiana (9), Lapis (9), effequ (4), People, Cronache Ribelli, Gallucci e pochi altri.
- Copertine: mancano quasi per tutti, tranne Terre di Mezzo, Corsare, Pelledoca, L'Ippocampo e in parte Uovonero.
- Prezzi: BeccoGiallo (7), Lapis (10), Camelozampa e pochi altri.
- Chiedere agli editori: ZIP di copertine, listino prezzi, catalogo completo (la mail modello è nel foglio "Email modello" dell'Excel).

## 7. Pubblicare il flipbook per le insegnanti
Il flipbook è un solo file HTML con tutto incorporato. Opzioni:
- caricarlo sul sito di Bottegart (o su Netlify / GitHub Pages / Google Drive con link pubblico) e mandare il link;
- se è ospitato fuori dall'anteprima di Claude, puoi lasciare `COPERTINE_ONLINE = true` nel blocco dati: per i libri senza copertina incorporata prova a caricarla da internet con l'ISBN (da verificare).
La scheda desiderata usa un link `mailto:` verso bottegart@gmail.com: apre la posta dell'insegnante, che deve premere Invio.

## 8. Limiti da ricordare
- Le copertine estratte dai PDF sono piccole (circa 100 dpi): vanno bene a video, non per stampa.
- Alcune sinossi sono descrizioni sintetiche (segnate "scheda sintetica: trama completa in arrivo") finché non si trova la scheda vera.
- Prima di inviare il link alle scuole serve un controllo umano a campione su età, prezzi e diritti d'uso di testi e immagini (chiedere l'autorizzazione agli editori).
