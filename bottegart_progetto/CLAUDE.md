# Progetto Bottegart / So'Stare – Flipbook "Libri per crescere"

## Obiettivo
Catalogo online (un solo file HTML) per insegnanti, con libri di case editrici indipendenti divisi per età,
con titolo, autore, editore, copertina, sinossi, prezzo e una scheda "desiderata" che apre una mail a bottegart@gmail.com.
Il catalogo completo vive nell'Excel; nel flipbook entra solo una SELEZIONE (circa 360 titoli: gli insegnanti non leggono mille titoli).

## Pipeline (da eseguire dalla radice del progetto)
1. `python build_xlsx.py`  -> output/catalogo_case_editrici_indipendenti.xlsx (+ isbn.json)
2. `python make_flipbook.py` -> output/flipbook_bottegart.html (usa covers.json e flip_template.html)
Le formule dell'Excel si ricalcolano aprendo il file (o con LibreOffice headless).

## File principali
- build_xlsx.py: dataset di base (righe scritte a mano per editore) + blocchi che applicano i dati dei cataloghi.
  Dentro c'è SELECT = insiemi di titoli ammessi nel flipbook per gli editori con cataloghi grandi (Terre di Mezzo, Logos, Minibombo, Pelledoca, Dalia).
- make_flipbook.py: legge l'Excel (foglio "Tutti i titoli", colonna "Nel flipbook" = Sì) e produce l'HTML.
- flip_template.html: copertina, pagine (>=4 libri per pagina), indice, scheda desiderata. I dati dei libri sono nel blocco "DATI DEI LIBRI" dell'HTML generato.
- *_data.py / uo_cat.py: dati estratti dai cataloghi (prezzi, età, ISBN, sinossi) per editore.
- covers.json: copertine in base64 (chiave = ISBN, oppure "t:" + titolo normalizzato). Estratte dai PDF con extract_covers.py.
- cataloghi/: PDF/Excel/immagini caricati dall'utente (i PDF grandi vanno rimessi qui, vedi README).

## Regole da rispettare
- Non inventare trame, prezzi, età: se un dato non c'è nella fonte resta vuoto ("da verificare").
- Le sinossi dagli PDF si riscrivono/puliscono ma non si cambia il contenuto.
- L'utente carica un catalogo alla volta: usare il catalogo SOLO per completare i titoli già in selezione (non aggiungere tutto il catalogo). Si può proporre di aggiungere titoli con temi sociali o CAA, chiedendo prima.
- Priorità tematiche: libri in CAA e temi sociali (bullismo, inclusione, ambiente, diritti, genere, migrazione).
- Controllare a vista (contact sheet) le copertine estratte prima di tenerle; scartare QR code, crop con testo, abbinamenti scambiati.
- Tutto il testo visibile in italiano con apostrofi tipografici (’) e virgolette “ ”.
- Il flipbook finale deve restare un SINGOLO file HTML autosufficiente (immagini incorporate).
