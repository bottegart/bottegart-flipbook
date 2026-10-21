# Mettere online il catalogo e la sezione "La nostra libreria"

## 1. Il link del catalogo (GitHub Pages) – una volta sola
1. Apri https://github.com/bottegart/bottegart-flipbook/settings/pages
2. **Source**: "Deploy from a branch".
3. **Branch**: `main`, cartella `/docs`, poi **Save**.
4. Dopo 1-2 minuti il catalogo è online a https://bottegart.github.io/bottegart-flipbook/

Ogni volta che rigeneri il flipbook (`python make_flipbook.py`) e lo carichi su GitHub, il sito si aggiorna da solo.

## 2. La sezione in homepage su Flazio

### Se il tuo piano ha il componente Codici (piano Professional)
1. Apri l'editor di Flazio sulla homepage.
2. **Elementi → Utility → Codici** e trascina il componente nella pagina, dove vuoi la sezione.
3. Apri il file https://bottegart.github.io/bottegart-flipbook/sezione-libreria.html,
   oppure `docs/sezione-libreria.html` nel repository (pulsante "Raw"), seleziona tutto e copia.
4. Incolla il testo nel componente Codici e salva.
5. Se il componente mostra la sezione tagliata, allunga il riquadro nell'editor finché si vede tutta
   (al computer serve circa 900 px di altezza, sul telefono di più).
6. Pubblica il sito e prova il pulsante "Sfoglia il catalogo".

I numeri (libri, case editrici, CAA, fasce) si aggiornano nel file quando rigeneri il flipbook:
se cambiano, ricopia e reincolla la sezione.

### Se il tuo piano non ha il componente Codici
Rifai la sezione con i blocchi normali di Flazio:
- **Immagine**: https://bottegart.github.io/bottegart-flipbook/copertine.jpg
  (collegamento: https://bottegart.github.io/bottegart-flipbook/, apri in una nuova finestra)
- **Titolo**: "La nostra libreria"
- **Testo**: "Scegliamo libri di case editrici indipendenti per bambine, bambini e ragazzi. Tanti titoli sono in CAA,
  con il testo accompagnato da simboli, e molti parlano di inclusione, ambiente, diritti e di come si cresce insieme.
  Per le scuole abbiamo preparato un catalogo da sfogliare online, diviso per età: le insegnanti scelgono i libri,
  compilano la lista desiderata e ce la mandano con un clic."
- **Pulsante**: "Sfoglia il catalogo →", collegamento https://bottegart.github.io/bottegart-flipbook/ (nuova finestra)
- Colori del catalogo: blu #1F8FD6 (pulsante), testo #121A26, sfondo #FBFBF9;
  arcobaleno delle fasce: #EE4A5F, #F07A2A, #F2B705, #3DAE5A, #1F9D8B, #1F8FD6, #4A5FC1, #9B3FA0.
- Caratteri (se Flazio permette i Google Fonts): Bricolage Grotesque per i titoli e il testo, Instrument Serif corsivo per la parola "libreria".
