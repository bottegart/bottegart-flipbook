import openpyxl
# -*- coding: utf-8 -*-
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

F0, F3, FP, F8, F9, F10, FM, FA, FN = (
    "0-3 anni", "3-6 anni", "Prime letture (6-8)", "8+", "9+", "10+",
    "Scuola media (11-13)", "Adolescenza (14+)", "n.d. / tutte le età")
FASCE = [F0, F3, FP, F8, F9, F10, FM, FA, FN]
TAGF = {F0: "0-3 anni", F3: "3-6 anni", FP: "prime letture", F8: "8+", F9: "9+", F10: "10+",
        FM: "prima/seconda/terza media", FA: "adolescenza", FN: ""}

CAT, SIM = "Libro in simboli (CAA)", "Libro in simboli (CAA)"
U_URL = "https://www.uovonero.com/libri"
S_URL = "https://www.settenove.it/catalogo/narrativa/5"
S_AGE = "https://www.informafamiglie.it/bibliografia-e-filmografia/bibliografia/identita-e-stereotipi-di-genere"
C_URL = "https://www.camelozampa.com/gruppi-di-lettura/"
M_URL = "https://www.lameridiana.it/categorie/libri-per-una-cultura-accessibile/libri-accessibili.html"

# (editore, titolo, autori, collana, età testo catalogo, fascia, genere, tema sociale, CAA, verifica, fonte)
R = []
ISBNS = {}
def add(pub, t, a, col, eta, fascia, gen, tema="", caa="No", ver="Catalogo", src=""):
    R.append([pub, t, a, col, eta, fascia, gen, tema, caa, ver, src])

# ---------------- UOVONERO ----------------
P = "Uovonero"
for t, a, col, eta, f, tema in [
    ("Il brutto anatroccolo", "a cura di E. Crivelli; ill. A. Papini", "Pesci parlanti", "0/3, 3/6", F0, ""),
    ("Il gatto con gli stivali", "a cura di E. Crivelli; ill. R. Angeletti", "Pesci parlanti", "0/3, 3/6", F0, ""),
    ("I vestiti nuovi dell'imperatore", "a cura di E. Crivelli; ill. N. Corfini", "Pesci parlanti", "3/6", F3, ""),
    ("I musicanti di Brema", "a cura di E. Crivelli; ill. A. Manca", "Pesci parlanti", "0/3, 3/6", F0, "Cooperazione"),
    ("Biancaneve", "a cura di E. Crivelli; ill. T. D'Incalci", "Pesci parlanti", "0/3, 3/6", F0, ""),
    ("Il lupo e i sette capretti", "a cura di E. Crivelli; ill. A. Alemanno", "Pesci parlanti", "0/3, 3/6", F0, ""),
    ("Raperonzolo", "a cura di E. Crivelli; ill. A. Boffa", "Pesci parlanti", "0/3, 3/6", F0, ""),
    ("Che rabbia!", "Mireille D'Allancé", "I libri di Camilla", "0/3, 3/6", F0, "Emozioni (rabbia)"),
    ("Questo posso farlo", "Satoe Tone", "I libri di Camilla", "0/3, 3/6", F0, "Disabilità"),
    ("Ninna nanna per una pecorella", "E. Bellini, M. Caccia", "I libri di Camilla", "0/3, 3/6", F0, ""),
    ("Una scatola gialla", "Pieter Gaudesaboos", "I libri di Camilla", "0/3, 3/6", F0, ""),
    ("Lei ci sarà sempre", "Thierry Lenain; ill. Manon Gauthier", "I libri di Camilla", "0/3, 3/6", F0, ""),
    ("Il piccolo coniglio bianco", "Xosé Ballesteros; ill. Óscar Villán", "I libri di Camilla", "0/3, 3/6", F0, ""),
    ("Lindo Porcello", "Eric Battut", "I libri di Camilla", "0/3", F0, ""),
    ("Piccolo Uovo", "Francesca Pardi; ill. Altan", "I libri di Camilla", "0/3, 3/6", F0, "Famiglie diverse (tema noto del libro, da verificare)"),
    ("Le parole di Bianca sono farfalle", "Chiara Lorenzoni; ill. Sophie Fatus", "I libri di Camilla", "3/6, 6/9", F3, "Disabilità / comunicazione"),
]:
    add(P, t, a, col, eta, f, SIM, tema, "Sì", "Catalogo", U_URL)
add(P, "La principessa sul pisello", "a cura di E. Crivelli; ill. F. Corso", "Pesci parlanti", "n.d.", FN, SIM, "", "Sì", "Da verificare", U_URL)
for t in ("Domino i verbi con i simboli", "Domino gli aggettivi con i simboli"):
    add(P, t, "E. Crivelli; ill. A. Manca", "Altrimenti (giochi)", "0/3, 3/6", F0, "Gioco/Activity book", "", "Sì", "Catalogo", U_URL)
for t in ("I quaderni di #intantofaccioqualcosa 4 – Attività con la CAA", "I quaderni di #intantofaccioqualcosa 6 – Attività con la CAA 2"):
    add(P, t, "a cura di Enza Crivelli", "I quaderni di #IFQ", "n.d.", FN, "Gioco/Activity book", "", "Sì", "Da verificare", "https://www.uovonero.com/libri/i-quaderni")

for t, a, col, eta, f, g, tema in [
    ("Dove sei, piccolo Giulio?", "Giovanni Colaneri", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Autismo"),
    ("Ad abbracciar nessuno", "Arianna Papini", "I Geodi", "3/6", F3, "Albo illustrato", "Autismo"),
    ("La distanza dei pesci", "C. Lorenzoni, G. Conoscenti", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Autismo"),
    ("Che bambino fortunato!", "L. Schimel, J. C. Mayorga", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Cecità / valore della differenza"),
    ("Che cos'è una sindrome?", "Giovanni Colaneri", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Disabilità"),
    ("La bambina che andava a pile", "Monica Taini", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Disabilità / identità"),
    ("Una canzone", "E. Vincenzi, I. Braiotta", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Malattia (Alzheimer) e nonni"),
    ("In cielo, ma dove?", "A. Ossorio, A. Ferrara", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Lutto"),
    ("Ti cerco", "Britta Teckentrup", "I Geodi", "3/6", F3, "Albo illustrato", "Emozioni"),
    ("Il mondo è rosso", "Britta Teckentrup", "I Geodi", "3/6, 6/9", F3, "Albo illustrato", "Emozioni (rabbia)"),
    ("La chitarra di Django", "Fabrizio Silei, Alfred", "I Geodi", "6/9", FP, "Albo illustrato", "Disabilità (Django Reinhardt)"),
    ("L'orco del piano di sotto", "Olivier Dupin; ill. Barroux", "I Geodi", "6/9", FP, "Albo illustrato", "Memoria della Shoah / Giusti"),
    ("La vecchia casa sul canale", "Britta Teckentrup", "I Geodi", "6/9, 9/12", FP, "Albo illustrato", "Memoria (la casa di Anne Frank)"),
    ("La scuola", "Britta Teckentrup", "I Geodi", "6/9, 9/12", FP, "Albo illustrato", "Comunità scolastica"),
    ("Rumple Buttercup", "Matthew Gray Gubler", "I Geodi", "6/9", FP, "Albo illustrato", "Unicità e amicizia"),
    ("Lilo", "Inés Garland; ill. Maite Mutuberria", "I Geodi", "6/9, 9/12", FP, "Narrativa", "Dipendenza dallo smartphone"),
    ("Vi presento Hank 1 – Un segnalibro in cerca d'autore", "H. Winkler, L. Oliver; ill. G. Orecchia", "Abbecedanze (alta leggibilità)", "6/9", FP, "Narrativa", "DSA (dislessia)"),
    ("Può succedere ancora?", "Jesper Lundqvist", "I Geodi", "9/12", F9, "Divulgazione/Saggio", "Memoria della Shoah"),
    ("Una specie di scintilla", "Elle McNicoll", "I Geodi", "9/12, 12+", F9, "Romanzo", "Autismo (autrice autistica)"),
    ("Castelli di fiammiferi", "Bettina Obrecht", "I Geodi", "9/12, 12+", F9, "Romanzo", "Autismo"),
    ("Amiche d'ombra", "Arianna Papini", "I Geodi", "9/12", F9, "Narrativa", "Cecità / pregiudizi"),
    ("Farfariel – Il libro di Micù", "Pietro Albì", "I Geodi", "9/12, 12+", F9, "Romanzo", "Disabilità"),
    ("Garrincha. L'angelo dalle gambe storte", "Antonio Ferrara", "I Geodi", "9/12, 12+", F9, "Graphic novel/Fumetto", "Disabilità"),
    ("Hank Zipzer e le cascate del Niagara (vol. 1)", "H. Winkler, L. Oliver", "Abbecedanze (alta leggibilità)", "9/12", F9, "Narrativa", "DSA (dislessia)"),
    ("Un pesce sull'albero", "Lynda Mullaly Hunt", "Abbecedanze (alta leggibilità)", "9/12, 12+", F9, "Romanzo", "DSA (dislessia)"),
    ("Non è colpa della pioggia", "Lynda Mullaly Hunt", "Abbecedanze (alta leggibilità)", "9/12, 12+", F9, "Romanzo", "Famiglia e comunità"),
    ("La più bella nuotata della mia vita", "Anne Becker", "I Geodi", "9/12, 12+", F9, "Romanzo", "DSA"),
    ("30 giorni per capire l'autismo", "Babe, Baroukh, Dunana (cura E. Crivelli)", "Altrimenti", "9/12, 12+, 14+", F9, "Divulgazione/Saggio", "Autismo"),
    ("30 giorni per capire i disturbi visivi", "Babe, Baroukh, Dunana (cura E. Crivelli)", "Altrimenti", "9/12, 12+, 14+", F9, "Divulgazione/Saggio", "Cecità / disabilità"),
    ("30 giorni per capire i disturbi dell'apprendimento", "Babe, Baroukh, Dunana (cura E. Crivelli)", "Altrimenti", "9/12, 12+, 14+", F9, "Divulgazione/Saggio", "DSA"),
    ("Keedie", "Elle McNicoll", "I Geodi", "12+", FM, "Romanzo", "Autismo"),
    ("Come un incantesimo", "Elle McNicoll", "I Geodi", "12+", FM, "Romanzo", "Neurodivergenza"),
    ("Mostraci chi sei", "Elle McNicoll", "I Geodi", "12+, 14+", FM, "Romanzo", "Autismo"),
    ("Il mistero del London Eye", "Siobhan Dowd", "I Geodi", "12+", FM, "Giallo/Mistero", "Autismo (Premio Andersen 2012)"),
    ("Il mistero del Guggenheim", "Robin Stevens (idea di S. Dowd)", "I Geodi", "12+, 14+", FM, "Giallo/Mistero", "Autismo"),
    ("Tutt'altro che tipico", "Nora Raleigh Baskin", "I Geodi", "12+", FM, "Romanzo", "Autismo"),
    ("La fisica degli abbracci", "Anna Vivarelli", "I Geodi", "12+, 14+", FM, "Romanzo", "Autismo"),
    ("Gabbie", "A. Vivarelli, G. Quarzo", "I Geodi", "12+, 14+", FM, "Giallo/Mistero", "Autismo"),
    ("Una piccola cosa senza importanza", "Catherine Fradier", "I Geodi", "12+, 14+", FM, "Avventura", "Autismo"),
    ("Dacca toxic", "Catherine Fradier", "I Geodi", "12+", FM, "Avventura", "Autismo"),
    ("Mio fratello è un Custode", "Marine Carteron", "I Geodi", "12+, 14+", FM, "Avventura", "Autismo"),
    ("Semplicemente Maria", "Jay Hardwig", "I Geodi", "12+", FM, "Romanzo", "Cecità / disabilità"),
    ("Professor papà", "Gunnar Helgason", "I Geodi", "12+", FM, "Romanzo", "Disabilità"),
    ("Una mamma svitata", "Gunnar Helgason", "I Geodi", "12+", FM, "Romanzo", "Famiglia e adolescenza"),
    ("Tutta la rabbia che ho dentro", "Inés Garland", "I Geodi", "12+", FM, "Romanzo", "Emozioni e adolescenza"),
    ("Tre dita", "Massimo Canuti", "I Geodi", "12+", FM, "Romanzo", "Memoria (Seconda guerra mondiale)"),
    ("Thornhill", "Pam Smy", "I Geodi", "12+, 14+", FM, "Narrativa", "Bullismo (tema noto del libro, da verificare)"),
    ("Il pavee e la ragazza", "Siobhan Dowd; ill. Emma Shoard", "I Geodi", "9/12, 12+, 14+", FM, "Narrativa", "Minoranze (comunità Pavee)"),
    ("Con le ali sbagliate", "Gabriele Clima", "I Geodi", "12+, 14+", FM, "Romanzo", "Identità sessuale e pregiudizi"),
    ("Divergente", "Victoria Grondin", "I Geodi", "14+", FA, "Romanzo", "Autismo / diversità"),
    ("L'aquilone di Noah", "Rafael Salmerón", "I Geodi", "14+", FA, "Romanzo", "Autismo; Seconda guerra mondiale"),
    ("Le degenerate", "J. Albert Mann", "I Geodi", "14+", FA, "Romanzo", "Disabilità (storia delle istituzioni, 1928)"),
    ("Il mio posto", "Daniele Bergesio", "I Geodi", "14+", FA, "Romanzo", "Emozioni e adolescenza"),
]:
    add(P, t, a, col, eta, f, g, tema, "No", "Catalogo", U_URL)

# ---------------- SETTENOVE ----------------
P = "Settenove"
for t, a, eta, f, tema, ver, src in [
    ("Ettore. L'uomo straordinariamente forte", "Magali Le Huche", "5 anni", F3, "Stereotipi di genere", "Catalogo", S_AGE),
    ("Mi piace Spiderman... e allora?", "G. Vezzoli, M. Di Lauro", "5 anni", F3, "Stereotipi di genere", "Catalogo", S_AGE),
    ("Io sono così", "F. Degl'Innocenti, A. Ferrara", "4 anni", F3, "Identità e stereotipi di genere", "Catalogo", S_AGE),
    ("C'è qualcosa di più noioso che essere una principessa rosa?", "Raquel Díaz Reguera", "5 anni", F3, "Stereotipi di genere", "Catalogo", S_AGE),
    ("June e Lea", "Sandrine Bonini; ill. Sandra Desmazières", "da 6 anni", FP, "Stereotipi di genere / passaggio all'adolescenza", "Catalogo", "https://www.settenove.it/articoli/june-e-lea/303"),
    ("Papà aspetta un bimbo!", "F. Loew; ill. S.-Y. Barroux", "n.d.", FN, "Ruoli familiari e genitorialità", "Da verificare", S_URL),
    ("Selvaggia", "Emily Hughes", "n.d.", FN, "Libertà e identità", "Da verificare", S_URL),
    ("Il piccolo giardiniere", "Emily Hughes", "n.d.", FN, "Libertà e identità", "Da verificare", S_URL),
    ("Storia di Giulia che aveva un'ombra da bambino", "C. Bruel; ill. A. Bozellec", "n.d.", FN, "Identità di genere", "Da verificare", S_URL),
    ("Io sono Adila. Storia illustrata di Malala Yousafzai", "F. Degl'Innocenti; ill. A. Forlati", "n.d.", FN, "Diritto all'istruzione", "Da verificare", S_URL),
    ("La nudità che male fa?", "Rosie Haine", "n.d.", FN, "Corpo e nudità", "Da verificare", S_URL),
    ("Il bosco in casa", "Sarah Vegna; ill. Astrid Tolke", "da 4 anni", F3, "Famiglie omogenitoriali, riciclo e natura", "Catalogo", "https://www.ibs.it/bosco-in-casa-ediz-a-libro-sarah-vegna/e/9788898947355"),
    ("In famiglia!", "A. Maxeiner; ill. A. Kuhl", "n.d.", FN, "Famiglie ricomposte", "Da verificare", S_URL),
    ("Cosa faremo da grandi? Prontuario di mestieri per bambine e bambini", "Irene Biemmi; ill. L. Terranera", "n.d.", FN, "Stereotipi nei mestieri", "Da verificare", S_URL),
    ("Giorgia e Giorgio. W i nonni!", "C. Obber; ill. S. Vinciguerra", "n.d.", FN, "Parità e famiglia (serie Giorgia e Giorgio)", "Da verificare", S_URL),
    ("Se un bambino", "Davide Musso; ill. Anna Forlati", "n.d.", FN, "Diritti dell'infanzia", "Da verificare", S_URL),
    ("La strana storia di Cappuccetto Blu", "Guia Risari; ill. Clémence Pollet", "n.d.", FN, "Stereotipi nelle fiabe", "Da verificare", S_URL),
    ("Una damigella NON in pericolo!", "Bethan Stevens", "n.d.", FN, "Stereotipi nelle fiabe", "Da verificare", S_URL),
    ("Ada al contrario", "Guia Risari; ill. Francesca Buonanno", "da 4 anni", F3, "Autismo (Asperger) e diritto alla diversità", "Catalogo", "https://www.ibs.it/ada-al-contrario-ediz-a-libro-guia-risari/e/9788898947423"),
    ("Beaver, come Castoro", "Serena Ballista; ill. Martina Paderni", "n.d.", FN, "Identità", "Da verificare", S_URL),
    ("Una stanza tutta per me", "Serena Ballista; ill. Chiara Carrer", "n.d.", FN, "Autonomia delle donne (da Virginia Woolf)", "Da verificare", S_URL),
    ("Greta e le altre. Un pianeta da salvare", "F. Degl'Innocenti; ill. F. Rizzato", "n.d.", FN, "Ambiente e attivismo giovanile", "Da verificare", S_URL),
    ("Il primo voto di Matilde", "F. Degl'Innocenti; ill. G. Marchegiani", "n.d.", FN, "Diritto di voto delle donne", "Da verificare", S_URL),
    ("Come un fiume. Una storia chiamata femminismo", "Emilia Zazza, Anna Arianna Melone", "n.d.", FN, "Storia del femminismo", "Da verificare", S_URL),
    ("Period Girl", "Giorgia Vezzoli", "n.d.", FN, "Ciclo mestruale / corpo", "Da verificare", S_URL),
    ("Il mio super eserciziario femminista", "Claire Cantais", "n.d.", FN, "Femminismo", "Da verificare", S_URL),
    ("E io qui, nuda", "Cristina Obber", "n.d.", FA, "Violenza e consenso", "Da verificare", S_URL),
    ("Il mondo ha bisogno delle ragazze (di me)", "Alessandra Spada", "n.d.", FA, "Parità di genere", "Da verificare", S_URL),
    ("Rompi le scatole! Idee e storie per fare a pezzi stereotipi e pregiudizi", "Pina Caporaso", "n.d.", FN, "Stereotipi e pregiudizi", "Da verificare", "https://www.ibs.it/libri/editori/settenove"),
    ("Leggere senza stereotipi. Percorsi educativi 0-6 anni (per adulti)", "SCOSSE APS", "0-6 (per educatori)", FN, "Educazione al genere (strumento per insegnanti)", "Catalogo", "https://www.settenove.it/articoli/leggere-senza-stereotipi/317"),
]:
    gen = "Albo illustrato"
    if t.startswith(("Come un fiume", "Il mio super", "Rompi", "Leggere senza")): gen = "Divulgazione/Saggio"
    if t.startswith(("E io qui", "Il mondo ha bisogno", "June e Lea")): gen = "Narrativa"
    add(P, t, a, "Albi illustrati / Narrativa / Educazione al genere", eta, f, gen, tema, "No", ver, src)

# ---------------- CAMELOZAMPA ----------------
P = "Camelozampa"
for t, eta, f, g, tema in [
    ("Dove sono finiti i bambini?", "6-10 anni", FP, "Poesia", "Diritto dei bambini a essere ascoltati"),
    ("Isidoro e la luna", "7-11 anni", FP, "Avventura", "Potere della parola"),
    ("70 trucchi per non fare il bagno", "8-11 anni", F8, "Romanzo", "Ansia da prestazione"),
    ("Furia Mascherata", "8-11 anni", F8, "Avventura", "Amicizia"),
    ("Felice abbastanza", "9-12 anni", F9, "Romanzo", "Crescita / preadolescenza"),
    ("La mia vita imprevista", "9-12 anni", F9, "Romanzo", "Famiglia / affido"),
    ("La figlia della biblioteca", "10-14 anni", FM, "Avventura", "Famiglie che si scelgono"),
    ("Oggi non si torna a casa", "10-14 anni", FM, "Romanzo", "Identità e appartenenza"),
    ("Quell'estate con Yente", "10-14 anni", FM, "Romanzo", "Relazioni tossiche"),
    ("Il gruppo", "11-14 anni", FM, "Romanzo", "Essere sé stessi / pressione del gruppo"),
    ("Lemming e fiocchi di neve", "14-18 anni", FA, "Romanzo", "Salute mentale"),
    ("Vulcano", "14-18 anni", FA, "Romanzo", "Ricerca di sé"),
    ("Figlia del destino", "14-18 anni", FA, "Romanzo", "Condizione femminile (storico)"),
    ("Senza titolo", "14-18 anni", FA, "Romanzo", "Adolescenza"),
]:
    add(P, t, "(vedi scheda)", "Narrativa / Gruppi di lettura", eta, f, g, tema, "No", "Catalogo", C_URL)
for t, tema in [
    ("Dove nasco ogni giorno", "Famiglie di molte forme"),
    ("Il giardino di Rita", "Lutto e trasformazione"),
    ("Cielo di ciliegie", "Differenze sociali"),
    ("La sorpresa della volpe", "Amicizia e pregiudizi"),
    ("Avrò cura di te", "Cura e solidarietà"),
    ("Felicità è una parola semplice", "Elaborazione del dolore collettivo (pandemia)"),
    ("Il cortile selvaggio", "Natura e legami"),
]:
    add(P, t, "(vedi scheda)", "Le Piume (albi illustrati)", "collana prevalentemente da 4 anni", F3, "Albo illustrato", tema, "No", "Da verificare", "https://www.camelozampa.com/shop/categoria/collane/le-piume-albi-illustrati/")
add(P, "Peggy trova tutto", "Tonka Uzu", "Collana 0-3", "dai 2 anni", F0, "Libro cartonato", "", "No", "Catalogo", "https://camelozampa.com/doc/catalogo-camelozampa-2025.pdf")
add(P, "Amici (miniserie A bocca aperta)", "Helen Oxenbury", "A bocca aperta", "0-2 anni", F0, "Libro cartonato", "", "No", "Catalogo", "https://www.camelozampa.com/shop/amici/")

# ---------------- LA MERIDIANA ----------------
P = "La Meridiana"
for t, a, g in [
    ("A Said piaceva il mare", "n.d.", "Narrativa"), ("Il mio Afghanistan (CAA)", "n.d.", "Narrativa"),
    ("L'albero del riccio", "n.d.", "Albo illustrato"), ("La Puglia è una favola", "n.d.", "Albo illustrato"),
    ("Milagro in CAA", "n.d.", "Narrativa"), ("Canto di Natale", "Charles Dickens", "Classico illustrato"),
    ("Dracula", "Bram Stoker", "Classico illustrato"), ("Il diario di Anna Frank", "Anna Frank", "Narrativa"),
    ("Piccole donne", "Louisa May Alcott", "Classico illustrato"), ("Giacomo di cristallo e altre storie", "Gianni Rodari", "Classico illustrato"),
    ("Frida Kahlo pittrice coraggiosa", "Teresa Righetti", "Divulgazione/Saggio"), ("Paul Klee pittore musicista", "Teresa Righetti", "Divulgazione/Saggio"),
]:
    tema = ""
    if t == "A Said piaceva il mare": tema = "Migrazione / accoglienza (da verificare)"
    if t == "Il mio Afghanistan (CAA)": tema = "Migrazione / culture (da verificare)"
    if t == "Il diario di Anna Frank": tema = "Memoria della Shoah"
    if t.startswith("Frida"): tema = "Disabilità e coraggio (da verificare)"
    add(P, t, a, "Libri accessibili in CAA e in simboli (anche Parimenti. Proprio perché cresco)", "n.d. (collana Parimenti per giovani-adulti)", FA if t in ("Dracula", "Il diario di Anna Frank", "Piccole donne") else FN, "Libro in simboli (CAA)" if False else g, tema, "Sì", "Da verificare", M_URL)
add(P, "Cinque racconti brevi con protagonisti adolescenti (collana Parimenti)", "AA.VV.", "Parimenti. Proprio perché cresco", "adolescenti / giovani-adulti", FA, "Narrativa", "Inclusione: lettura per disabilità cognitive e linguistiche", "Sì", "Catalogo", "https://www.lameridiana.it/blog/post/cs-09112021/")
add(P, "Di sole e di tempesta", "G. Boari, E. Buccoliero", "Percorsi educativi", "per educatori/insegnanti", FN, "Divulgazione/Saggio", "Violenza assistita: sostegno ai bambini", "No", "Catalogo", "https://www.bookrepublic.it/ebooks/editore/edizioni-la-meridiana/")
add(P, "Accendiamo i diritti. Percorsi e attività per educare ai diritti dell'infanzia", "AA.VV.", "Percorsi educativi bambini", "per educatori/insegnanti", FN, "Divulgazione/Saggio", "Diritti dell'infanzia e dell'adolescenza", "No", "Catalogo", "https://www.libreriadelsanto.it/libri_di/editore/La_Meridiana/page1.html")

# ---------------- GALLUCCI ----------------
P = "Gallucci"
G_URL = "https://www.ibs.it/libri/editori/gallucci"
for t, col, eta, f, g, tema in [
    ("Collana Scorri e gioca (N. Choux)", "Scorri e gioca", "fino a 3 anni", F0, "Libro cartonato", ""),
    ("Collana Indistruttilibri", "Indistruttilibri", "prima infanzia", F0, "Libro cartonato", ""),
    ("Collana Abbecelibri", "Abbecelibri", "fino a 6 anni", F3, "Albo illustrato", ""),
    ("Collana Lilliput", "Lilliput", "+5", F3, "Albo illustrato", ""),
    ("Collana Prime letture (es. Fernando il coccodrillo)", "Prime letture", "+6", FP, "Narrativa", ""),
    ("Il mondo di Lucrezia 2", "UAO", "10+", F10, "Romanzo", ""),
    ("La casa nella prateria (serie)", "UAO", "UAO: 9+ / 10+ / 11+", F9, "Romanzo", ""),
    ("Collezione Libri e giochi delle emozioni", "Collezioni", "n.d.", FN, "Gioco/Activity book", "Emozioni"),
    ("Tu sei importante", "Illustrati", "n.d.", FN, "Albo illustrato", "Inclusione e rappresentazione"),
    ("Collana Young (es. Due angeli del fango)", "Young", "+14", FA, "Romanzo", ""),
    ("Collana Young adult (es. Happy. Mi troverai qui)", "Young adult", "+16", FA, "Romanzo", ""),
]:
    add(P, t, "(vedi catalogo)", col, eta, f, g, tema, "No", "Collana" if eta not in ("n.d.",) else "Da verificare", G_URL)

# ---------------- PELLEDOCA ----------------
P = "Pelledoca"
PL = "https://www.ibs.it/libri/editori/pelledoca-editore"
for t, a, col, eta, f, g, tema, ver, src in [
    ("Susan Duckling e il caso del regalo di compleanno", "Maddalena Schiavo; ill. Laura Zani", "Piccole Piume", "dai 7 anni", FP, "Giallo/Mistero", "", "Catalogo", "https://www.verbanonews.it/aree-geografiche/lombardia/2021/10/08/libri-per-cominciare-ad-affrontare-le-paure-in-vista-di-halloween/935987/"),
    ("Il mistero del tappeto peloso", "Maddalena Schiavo", "Piccole Piume", "prime letture", FP, "Giallo/Mistero", "", "Collana", PL),
    ("Il mistero della ricetta del Montblanc", "Maddalena Schiavo", "Piccole Piume", "prime letture", FP, "Giallo/Mistero", "", "Collana", PL),
    ("Nel castello di Azzurrina", "Carolina D'Angelo; ill. Cecilia Fallongo", "Piccole Piume (da verificare)", "n.d.", FP, "Giallo/Mistero", "", "Da verificare", "https://www.macrolibrarsi.it/edizioni/_pelledoca-editore.php"),
    ("Lupi, licantropi e luna piena", "Marco Greganti", "n.d.", "n.d.", FN, "Giallo/Mistero", "", "Da verificare", "https://www.librerie.coop/editori/pelledoca-editore/"),
    ("Bambina d'ombra", "Brian Freschi", "OcchiAperti", "n.d.", FN, "Graphic novel/Fumetto", "", "Da verificare", "https://www.librerie.coop/editori/pelledoca-editore/"),
    ("Uno scrigno a doppio segreto", "Anna Vivarelli, Guido Quarzo", "NeroInchiostro", "n.d. (protagonista dodicenne)", F10, "Avventura", "", "Da verificare", "https://www.librerie.coop/editori/pelledoca-editore/"),
    ("Non chiudere gli occhi", "Formaggi", "NeroInchiostro", "10+ (collana); protagonisti 12 anni", F10, "Giallo/Mistero", "", "Collana", "https://libreriamo.it/libri/nasce-pelledoca-la-nuova-casa-editrice/"),
    ("Il mistero del cadavere senza testa", "Luca Occhi", "NeroInchiostro", "11-14 (lista consigli di lettura)", FM, "Giallo/Mistero", "", "Catalogo", "https://www.comune.castelfrancoveneto.tv.it/wp-content/uploads/2025/05/Consigli-di-lettura-estate-11_14.pdf"),
    ("Dark web", "Sara Magnoli", "NeroInchiostro", "dagli 11 anni", FM, "Giallo/Mistero", "Vita digitale e vita reale", "Catalogo", "https://sapere.virgilio.it/rubriche/libri-gialli-bambini-12921"),
    ("Fuori dal branco", "Sara Magnoli", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("Alive", "Alessandro Pasquinucci", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("Villa Mannara", "Laura Orsolini", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("La Barabuia", "Gabriele Nanni", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("Una lunghissima notte", "Annalisa Strada", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("Lunamadre", "Teo Benedetti", "NeroInchiostro", "10+ (collana)", F10, "Giallo/Mistero", "", "Collana", PL),
    ("La fuga di Pat", "Roberto Piumini", "n.d.", "n.d.", FN, "Narrativa", "", "Da verificare", "https://www.hoepli.it/editori/pelledoca_editore.html"),
]:
    add(P, t, a, col, eta, f, g, tema, "No", ver, src)

# ---------------- MOMO ----------------
P = "Momo Edizioni"
MO = "https://momoedizioni.it/categoria/libri/libri-monelli/"
add(P, "Che cos'è il sesso?", "F. D'Onofrio, S. Montanaro; ill. L. Montalto", "Libri Monelli", "da 7 anni", FP, "Divulgazione/Saggio", "Educazione affettiva e sessuale", "No", "Catalogo", "https://www.libraccio.it/libro/9791280298102/francesca-d")
add(P, "Bollicina", "Ciampacavallo (copertina di Zerocalcare)", "Libri Monelli – CiampaLibri", "24 pp.; età n.d.", FN, "Libro in simboli (CAA)", "Disabilità: libro scritto da persone con disabilità; PECS CAA e QR audio", "Sì", "Da verificare", "https://momoedizioni.it/catalogo/bollicina/")
for t, a, g, tema in [
    ("Animalia. Il mio primo bestiario", "", "Albo illustrato", ""),
    ("Portami al mare. Un viaggio nel Mediterraneo per piccoli scienziati", "", "Divulgazione/Saggio", "Ambiente"),
    ("SOS Oceano. Impariamo a difendere il mare", "E. Fanelli, E. Azzurro", "Divulgazione/Saggio", "Ambiente"),
    ("Karukinka. La nostra terra", "Erica Silvestri", "Albo illustrato", "Ambiente"),
    ("Cambiare il mondo con i libri. L'incredibile storia di Giangiacomo Feltrinelli", "Mattia Tombolini; ill. Marta Baroni", "Narrativa", "Storia e impegno civile"),
    ("Essere tempesta. Vita e morte di Giacomo Matteotti", "", "Narrativa", "Antifascismo / storia"),
    ("Mia patria sono gli oppressi. Don Milani, la scuola, la guerra", "", "Narrativa", "Scuola, pace, diritti"),
    ("Libera dal ciclo", "Flavia Fazi", "Divulgazione/Saggio", "Ciclo mestruale / corpo"),
    ("Gli adulti fanno schifo", "Gina", "Narrativa", "Adolescenza e rapporto con gli adulti"),
    ("Genitori che puzzano", "", "Narrativa", "Famiglia"),
    ("Il carcere è un mondo di carta", "", "Narrativa", "Carcere e diritti"),
    ("Il lago che combatte", "Militant A", "Narrativa", "Ambiente / lotta sociale"),
    ("Ad occhi aperti", "", "Narrativa", ""),
    ("A cosa servono le cose?", "", "Albo illustrato", ""),
]:
    add(P, t, a or "n.d.", "Libri Monelli", "n.d.", FN, g, tema, "No", "Da verificare", MO)

# ---------------- TUNUÈ ----------------
P = "Tunuè"
TU = "https://www.tunue.com/product-category/young-adult/"
for t, a, col, eta, f, g, tema, src in [
    ("Timothy Top (serie)", "", "Kids (fumetti)", "6+", FP, "Graphic novel/Fumetto", "", "https://www.tunue.com/product-category/graphic-novel-per-ragazzi/"),
    ("Investigators (serie)", "", "Kids (fumetti)", "6-12", FP, "Graphic novel/Fumetto", "", "https://www.tunue.com/product-category/graphic-novel-per-ragazzi/"),
    ("Monster Allergy (serie)", "", "Kids (fumetti)", "6-12", FP, "Graphic novel/Fumetto", "", "https://www.tunue.com/product-category/graphic-novel-per-ragazzi/"),
    ("Brina (serie)", "", "Kids (fumetti)", "7+", FP, "Graphic novel/Fumetto", "", "https://www.tunue.com/product-category/graphic-novel-per-ragazzi/"),
    ("Strega a modo mio (serie)", "", "Kids (fumetti)", "8+", F8, "Graphic novel/Fumetto", "", "https://www.tunue.com/product-category/graphic-novel-per-ragazzi/"),
    ("Le sorelle Grémillet (serie)", "Sophie Bertrand, ecc.", "Kids / Young Adult", "10+", F10, "Graphic novel/Fumetto", "Famiglia / crescita", TU),
    ("American Born Chinese", "Gene Luen Yang", "Tipitondi / Young Adult", "n.d.", FM, "Graphic novel/Fumetto", "Bullismo, identità e razzismo", "https://www.tunue.com/product-tag/bullying/"),
    ("Noi due", "", "Young Adult", "12-21", FA, "Graphic novel/Fumetto", "Relazioni", TU),
    ("L'Attraversaspecchi, vol. 1 – Fidanzati dell'inverno", "", "Young Adult", "12-21", FA, "Graphic novel/Fumetto", "", TU),
    ("Angelica e il principe orso (in arrivo)", "", "Young Adult", "12-21", FA, "Graphic novel/Fumetto", "", TU),
]:
    add(P, t, a or "n.d.", col, eta, f, g, tema, "No", "Catalogo" if eta != "n.d." else "Da verificare", src)

# ---------------- CRONACHE RIBELLI ----------------
P = "Cronache Ribelli"
CR = "https://cronacheribelli.it/collections/libri-per-bambin-e-ragazz"
add(P, "C'era una (ri)volta – Quaranta storie da leggere e colorare (vol. I)", "AA.VV.", "Libri per bambin* e ragazz*", "6-10 anni", FP, "Narrativa", "Diritti, antirazzismo, omofobia, violenza di genere, povertà", "No", "Catalogo", CR)
add(P, "C'era una (ri)volta – Quaranta storie da leggere e colorare (vol. II)", "AA.VV.", "Libri per bambin* e ragazz*", "6-10 anni", FP, "Narrativa", "Diritti, antirazzismo, lotte sociali", "No", "Catalogo", CR)
add(P, "Laika e i cani venuti dallo spazio", "n.d.", "Libri per bambin* e ragazz*", "n.d.", FN, "Albo illustrato", "", "No", "Da verificare", CR)

# ---------------- DALIA ----------------
P = "Dalia Edizioni"
DR = "https://www.daliaedizioni.it/categoria-prodotto/dalia-ragazzi/"
for t, a, g, tema in [
    ("Argo e Calypso, cani da biblioteca", "Andrea Laprovitera", "Narrativa", "Lettura e biblioteche"),
    ("Il mistero di Pyrgi. Avventura tra gli Etruschi", "Chiarastella Campanelli", "Avventura", "Storia"),
    ("Alla scoperta della libertà", "Giulia Ceccarani", "Narrativa", "Libertà e regole"),
    ("Il mio superpotere è la gentilezza", "Giulia Ceccarani", "Narrativa", "Gentilezza e rispetto"),
    ("Slò, la magia della lentezza", "n.d.", "Narrativa", ""),
    ("Arturo e Nadir", "Roberta Argenti", "Narrativa", ""),
    ("Il vampiro Cicorione", "n.d.", "Narrativa", ""),
    ("Damian e la lupa", "n.d.", "Narrativa", ""),
    ("Febe e i gatti ruba storie", "Elisa Campaci", "Narrativa", ""),
    ("Loris Wood verso l'isola di plastica", "Iride Bartolucci", "Avventura", "Ambiente (inquinamento da plastica)"),
]:
    add(P, t, a, "Dalia Ragazzi", "collana 5+ (IBS: prevalentemente da 7)", FP, g, tema, "No", "Collana", DR)
add(P, "UVA. Una Vita Altrove", "Giovanni Capotombolo", "Dalia Narrativa", "adolescenti", FA, "Romanzo", "Adolescenza e futuro", "No", "Da verificare", "https://www.librerie.coop/editori/dalia-edizioni/")
add(P, "Collana Gli spericolati (storie comiche, romanzo breve illustrato)", "(in progetto)", "Gli spericolati", "8+", F8, "Narrativa", "Possibilità del mondo, rispetto", "No", "Collana", "https://www.daliaedizioni.it/proposte/")

# ---------------- CORSARE ----------------
P = "Edizioni Corsare"
SC = "https://www.edizionicorsare.it/categoria-prodotto/collane/storie-corsare/"
for t in ["Alieni", "La città d'acqua", "Il ladro di Picasso", "Leopard boy", "Roka il figlio della terra", "Due detective e mezzo",
          "Bradi va in città", "Virginia e il tesoro perduto di Michelangelo", "Frottole e segreti", "Un emoji per amico",
          "Viola, la gallina e il soldato", "Ryunio. La bambina-drago", "Miss Mina", "Due ragazzi nella Firenze dei Medici"]:
    g = "Giallo/Mistero" if t.startswith(("Due detective", "Il ladro")) else ("Avventura" if t.startswith(("Leopard", "Roka", "Virginia", "Alieni")) else "Romanzo")
    tema = "Amicizia e tecnologia" if t.startswith("Un emoji") else ("Guerra e memoria" if t.startswith("Viola") else "")
    add(P, t, "(vedi scheda)", "Storie corsare", "9-13 anni", F9, g, tema, "No", "Collana", SC)
add(P, "Pericolo smartphone", "(vedi scheda)", "Storie corsare", "9-13 anni (+ guida per adulti)", F9, "Divulgazione/Saggio", "Uso consapevole di web e social", "No", "Collana", SC)
add(P, "Wonder Football Club (serie)", "Marco Innocenti", "Wonder Football Club", "da 10 anni (IBS)", F10, "Romanzo", "Sport e squadra", "No", "Catalogo", "https://www.ibs.it/libri/collane/storie-corsare-p204353")
add(P, "La collanina (prime letture autonome)", "(vedi collana)", "La collanina", "da 7 anni", FP, "Narrativa", "", "No", "Collana", "https://www.edizionicorsare.it/collane/")
for t, tema in [("L'estate in cui Ethan salvò San Francisco", ""), ("Americhea", ""), ("Curare le pietre", ""),
                ("Diana e il segreto dei lupasky", ""), ("L'ultimo tornado", ""), ("Quattro giorni", "")]:
    add(P, t, "(vedi scheda)", "Narrativa YA / Border", "da 14 anni", FA, "Romanzo", tema, "No", "Collana", "https://www.edizionicorsare.it/categoria-prodotto/collane/border-young-adult/")

for t, a, pr in [("Farfalle", "Lisa Riccardi; ill. A. Mosca", 18.0), ("Brava Chiocciolina!", "Silvia Vecchini", 16.0),
                 ("Il bimboleone e altri bambini", "Gabriele Clima; ill. Giacomo Agnello Modica", 18.0)]:
    add("Edizioni Corsare", t, a, "Area ragazzi (albi illustrati)", "n.d.", FN, "Albo illustrato", "", "No", "Da verificare", "https://www.libraccio.it/editore/edizioni-corsare/libri.html")

# ---------------- MINIBOMBO ----------------
P = "Minibombo"
MB = "https://www.minibombo.it/"
for t, a, eta, f, g in [
    ("Orso, buco!", "Nicola Grossi", "0+ / 3+ (da verificare)", F0, "Albo illustrato"),
    ("Tutino con la terra", "Silvia Borando", "da 1 anno", F0, "Libro cartonato"),
    ("Tutino nella neve", "Silvia Borando", "da 1 anno", F0, "Libro cartonato"),
    ("Cerchio al parco", "Silvia Borando", "n.d.", F0, "Silent book"),
    ("Il libro bianco", "Silvia Borando", "dai 2 anni", F0, "Silent book"),
    ("Grande gatto piccolo gatto", "Silvia Borando", "dai 2 anni", F0, "Albo illustrato"),
    ("Affamato come un lupo", "Silvia Borando", "da 3 anni", F3, "Albo illustrato"),
    ("Ha visto la mia coda?", "Alberto Lot", "dai 3 anni", F3, "Albo illustrato"),
    ("Gabbiano più gabbiano meno", "(vedi scheda)", "da 3 anni", F3, "Albo illustrato"),
    ("Allarme rosso!", "Marco Scalcione", "da 3 anni", F3, "Albo illustrato"),
    ("Ho visto una talpa", "(vedi scheda)", "da 3 anni", F3, "Albo illustrato"),
    ("Storie brevi", "Silvia Borando", "dai 3 anni", F3, "Albo illustrato"),
    ("Il dubbio", "S. Borando, A. Lot", "n.d.", F3, "Albo illustrato"),
    ("Un segreto importante", "A. Mozzillo; ill. E. Marzano", "n.d.", F3, "Albo illustrato"),
    ("Il cespuglio", "S. Borando, E. Francescutto, D. Pomodoro", "n.d.", F3, "Albo illustrato"),
    ("Il piccolo fienile rosso", "Kim Crumrine", "n.d.", F3, "Albo illustrato"),
    ("Gatto Nero, Gatta Bianca", "Silvia Borando", "n.d.", F3, "Albo illustrato"),
]:
    ver = "Catalogo" if eta not in ("n.d.", "0+ / 3+ (da verificare)") else "Da verificare"
    add(P, t, a, "Catalogo Minibombo (0-6 anni)", eta, f, g, "", "No", ver, MB)
add(P, "La mia famiglia è uno zoo", "Teresa Sdralevich", "Catalogo Minibombo", "n.d.", F3, "Albo illustrato", "Famiglie diverse (da verificare)", "No", "Da verificare", "https://librerie.unicatt.it/editore-minibombo-2842.html")
for t in ("Orso, buco!", "Affamato come un lupo", "Il libro gatto"):
    add(P, t + " – versione in simboli (InBook)", "traduzione in simboli a cura di reti di biblioteche", "InBook (non edizione Minibombo)", "n.d.", F3, "Libro in simboli (CAA)", "Inclusione: libro in simboli", "Sì", "Da verificare",
        "https://www.055firenze.it/art/216489/libri-in-simboli-nasce-la-prima-Rete-Operativa-Toscana-delle-Biblioteche-Inbook-libri-in-simboli")

# ---------------- PEOPLE ----------------
P = "People"
add(P, "Non esistono piccoli antifascisti", "Johannes Bückler", "Kids / Immaginazione", "n.d.", FN, "Narrativa", "Antifascismo", "No", "Da verificare", "https://www.libreriauniversitaria.it/libri-editore_People-people.htm")
add(P, "Io non ho paura del lupo", "Tommaso D'Errico", "Kids", "n.d.", FN, "Albo illustrato", "", "No", "Da verificare", "https://www.ibs.it/libri/editori/people")
add(P, "Immagina. Crea. Vivi!", "n.d.", "Kids", "n.d.", FN, "Albo illustrato", "", "No", "Da verificare", "https://www.hoepli.it/editori/people.html")

# ---------------- NERO LATTE ----------------
add("Nero Latte", "La bottega di Bepo il sarto", "Alessandra D'Amico", "Albi illustrati", "n.d.", FN, "Albo illustrato", "", "No", "Da verificare", "https://leggeretutti.eu/nasce-una-nuova-casa-editrice-nero_latte-edizioni/")
add("Nero Latte", "Verso il mondo (poesia, per adulti)", "Pier Francesco Latte", "Poesia", "adulti", FN, "Poesia", "", "No", "Catalogo", "https://leggeretutti.eu/nasce-una-nuova-casa-editrice-nero_latte-edizioni/")
add("Nero Latte", "Catalogo: poesia e albi illustrati; collabora con circa 30 librerie indipendenti", "Patrizia Baglione (editrice)", "Poesia / Albi illustrati", "n.d.", FN, "—", "", "No", "Da verificare", "https://www.pennaespada.it/in-pochi-mesi-la-casa-editrice-in-contatto-con-ben-33-librerie-indipendenti")

# ---------------- LOGOS (catalogo aggiornato 2026 fornito dall'editore) ----------------
import re as _re
PRICES_L = {}
P = "Logos"
_wl = openpyxl.load_workbook('cataloghi/LISTA_TITOLI_LOGOSEDIZIONI_2026_1.xlsx', data_only=True).active
_SM = set("e di del della dei delle degli il la lo i gli le un una uno in da con per a al alla ai nel nella su che non ma o".split())
def _tc(t):
    out = []
    for i, w in enumerate(str(t).strip().lower().split()):
        if "’" in w or "'" in w:
            p0, p1 = _re.split(r"[’']", w, 1)
            w = p0 + "’" + (p1[:1].upper() + p1[1:] if p1 else "")
            if i == 0: w = w[:1].upper() + w[1:]
        elif i == 0 or w not in _SM:
            w = w[:1].upper() + w[1:]
        out.append(w)
    return " ".join(out)
_AG = {"0-3": ("0-3 anni", F0), "da 3 anni": ("da 3 anni", F3), "da 6 anni": ("da 6 anni", FP), "da 8 anni": ("da 8 anni", F8), "da 12 anni": ("da 12 anni", FM)}
_GN = {"ALBO CIOPI": "Albo illustrato", "ILLUSTRATI": "Albo illustrato", "BIBLIOTECA CIOPI": "Narrativa", "FUMETTO CIOPI": "Graphic novel/Fumetto", "FUMETTO": "Graphic novel/Fumetto"}
for r in _wl.iter_rows(min_row=2, values_only=True):
    if not r[1] or not r[7]: continue
    eta = str(r[7]).strip().lower()
    if eta not in _AG: continue
    t = _tc(r[1]); col = str(r[6] or "").strip(); au = str(r[4]).strip() if r[4] else ""
    if r[5] and r[5] != r[4]: au = (au + "; ill. " + str(r[5])) if au else "ill. " + str(r[5])
    ISBNS[("Logos", t)] = str(r[0])
    add(P, t, au or "n.d.", col.title(), _AG[eta][0], _AG[eta][1], _GN.get(col, "Albo illustrato"), "", "No", "Catalogo", r[8] or "https://www.libri.it")
    if isinstance(r[3], (int, float)): PRICES_L[(P, t)] = (float(r[3]), "Lista titoli Logos Edizioni 2026 (PVP)")

# ---------------- LAPIS ----------------
P = "Lapis"
LP = "https://www.edizionilapis.it/collane"
for t, a, col, eta, f, g, tema in [
    ("Avventure di Tato e Nanà", "(vedi collana)", "Tato e Nanà", "dai 18 mesi", F0, "Libro cartonato", ""),
    ("Libri gioco cartonati", "(vedi collana)", "Libri gioco", "da 1 anno", F0, "Libro cartonato", ""),
    ("I senza parole di Attilio", "A. Cassinelli", "I senza parole", "0-3 / 3-6", F0, "Silent book", ""),
    ("Le mini storie di Attilio", "A. Cassinelli", "Libri cartonati", "dai 3 anni", F3, "Libro cartonato", "Ecologia, natura, amicizia"),
    ("Storie di una principessina (Tilly)", "(vedi collana)", "Storie di una principessina", "circa 4 anni", F3, "Albo illustrato", ""),
    ("I libri di Gek Tessaro / Il mondo di Tony Ross / Il mondo di Salina Yoon", "(vedi collane)", "Albi 3-6", "3-6 anni", F3, "Albo illustrato", ""),
    ("Mabel (prime letture)", "(vedi collana)", "Mabel", "prime letture", FP, "Narrativa", ""),
    ("Dardo il cane di Leonardo", "(vedi collana)", "Dardo", "dai 7 anni", FP, "Narrativa", ""),
    ("Versi diversi / Sì, io sono (poesia)", "Friot, Tullet, Carminati", "Poesia", "da 7 anni", FP, "Poesia", ""),
    ("Omero – Iliade / Odissea per ragazzi", "Cinquetti, Guicciardini", "Classici", "dai 8 anni", F8, "Classico illustrato", ""),
    ("I misteri di Teo", "(due autrici)", "I misteri di Teo", "da 9 anni", F9, "Giallo/Mistero", ""),
    ("Collana I lapislazzuli", "Cognolato; Lorenzoni; ecc.", "I lapislazzuli", "dai 4 anni", F3, "Albo illustrato", "Temi sociali (storia, guerre, persone e luoghi)"),
]:
    ver = "Collana" if eta not in ("n.d.",) else "Da verificare"
    add(P, t, a, col, eta, f, g, tema, "No", ver, LP)
add(P, "Foto di gruppo (versione in simboli InBook)", "Gek Tessaro", "InBook (traduzione di rete bibliotecaria)", "n.d.", F3, "Libro in simboli (CAA)", "Inclusione: libro in simboli", "Sì", "Da verificare",
    "https://www.055firenze.it/art/216489/libri-in-simboli-nasce-la-prima-Rete-Operativa-Toscana-delle-Biblioteche-Inbook-libri-in-simboli")

# ---------------- L'IPPOCAMPO ----------------
P = "L'Ippocampo"
IP = "https://www.ippocampoedizioni.it/catalogo/tema/107"
IM = "https://www.ippocampoedizioni.it/catalogo/tema/135"
add(P, "Bebè 0-2 anni (sezione)", "(vedi catalogo)", "L'Ippocampo Ragazzi", "0-2", F0, "Libro cartonato", "", "No", "Catalogo", "https://www.ippocampoedizioni.it/")
add(P, "Libri coccole Emma Dodd (sezione)", "Emma Dodd", "L'Ippocampo Ragazzi", "0-2", F0, "Libro cartonato", "", "No", "Catalogo", "https://www.ippocampoedizioni.it/")
add(P, "Prescuola (sezione)", "(vedi catalogo)", "L'Ippocampo Ragazzi", "prescuola", F3, "Albo illustrato", "", "No", "Catalogo", "https://www.ippocampoedizioni.it/")
for t, a, g in [
    ("L'isola del tesoro (ed. MinaLima)", "R. L. Stevenson", "Avventura"), ("20.000 leghe sotto i mari", "Jules Verne", "Avventura"),
    ("Il giro del mondo in 80 giorni", "Jules Verne", "Avventura"), ("Il libro della giungla", "Rudyard Kipling", "Avventura"),
    ("Peter Pan (ed. MinaLima)", "J. M. Barrie", "Avventura"), ("Il meraviglioso mago di Oz", "L. Frank Baum", "Avventura"),
    ("Maldoror – I ragazzi della leggenda (trilogia)", "Philippe Lechermeier", "Romanzo"), ("Le cronache dell'acero e del ciliegio (4 libri)", "Camille Monceaux", "Avventura"),
    ("Le streghe di Venezia (3 volumi)", "Sébastien Perez, Marco Mazzoni", "Romanzo"), ("Sangue d'ontano – La volpe artica", "Léa Versand", "Romanzo"),
    ("Terra rossa", "Jean-François Chabas", "Romanzo"), ("Il giardino segreto (ed. MinaLima)", "F. Hodgson Burnett", "Classico illustrato"),
]:
    add(P, t, a, "Narrativa ragazzi", "n.d. (sezione narrativa ragazzi)", FM if g != "Classico illustrato" else FN, g, "", "No", "Collana", IP)
for t, a, g in [
    ("Atlante dei relitti", "Jo Arnold, Philip Harris", "Divulgazione/Saggio"), ("Kidstory. La storia dell'umanità narrata dai ragazzi", "T. Weiss Gabbay, S. Fuman", "Divulgazione/Saggio"),
    ("Le tribolazioni del giovane Cosimo", "Marie Desplechin", "Romanzo"), ("Anatomia / Oceano / Nascita / Mondo vegetale (libri animati)", "Hélène Druvert", "Divulgazione/Saggio"),
    ("Enciclopedia del Meraviglioso (I draghi, Le streghe, I fantasmi...)", "Benjamin Lacombe e altri", "Divulgazione/Saggio"),
    ("Esplorando l'universo / Esplorando gli elementi", "Isabel Thomas, Sara Gillingham", "Divulgazione/Saggio"),
]:
    add(P, t, a, "Scuole medie (tema del catalogo)", "sezione Scuole medie", FM, g, "", "No", "Collana", IM)

# ---------------- NEO ----------------
P = "Neo Edizioni"
add(P, "GAP – Grottesco Adolescenziale Periferico", "Placido Di Stefano", "Dry / narrativa", "adulti (tema adolescenza)", FA, "Romanzo", "Adolescenza ai margini", "No", "Da verificare", "https://www.neoedizioni.it/neo/")
add(P, "Catalogo: collane Iena, Dry, Potlach, Intimate, i Nei, Cromo", "—", "Catalogo", "adulti", FN, "—", "", "No", "Catalogo", "https://www.neoedizioni.it/neo/categoria-prodotto/catalogo/")

# ---------------- TETRA ----------------
add("Tetra Edizioni", "Catalogo di narrativa contemporanea per adulti (nessun titolo per ragazzi)", "—", "Narrativa", "adulti", FN, "Narrativa", "", "No", "Catalogo", "https://www.tetraedizioni.com/")

# ---------------- EFFEQU ----------------
P = "effequ"
EF = "https://www.effequ.it/"
for t, a, tema in [("ZAMPA", "Francesca Matteoni; ill. Prisca Milanese", "Natura e relazioni tra esseri viventi"),
                   ("XY", "Lù Casini; ill. Caterina Di Paolo", "Identità e corpo"),
                   ("WEB", "Antonio Pavolini; ill. Nicola Giorgio", "Uso consapevole di internet"),
                   ("VIETATO", "Simone Tempia; ill. Marianna Balducci", "Regole e limiti")]:
    add(P, t, a, "Scatoline", "da 5 anni", F3, "Divulgazione/Saggio", tema, "No", "Catalogo", EF)
add(P, "Collana Rondini (narrativa) – es. Minerva, Come si smette di avere una faccia", "vari", "Rondini", "adulti", FN, "Narrativa", "", "No", "Catalogo", EF)

# ---------------- BECCOGIALLO ----------------
P = "BeccoGiallo"
BK = "https://www.beccogiallo.it/categoria-prodotto/kids/"
add(P, "Il bullismo spiegato ai bambini", "Andrea Laprovitera; ill. Mirti", "CriticalKids", "n.d. (48 pp.)", FN, "Albo illustrato", "Bullismo", "No", "Da verificare", "https://www.beccogiallo.it/prodotto/il-bullismo-spiegato-ai-bambini/")
for t, a, tema in [
    ("La sessualità spiegata ai bambini e alle bambine di ieri, oggi e domani", "Chiara Gregori; ill. J. Criveller", "Educazione sessuale e affettiva"),
    ("La Costituzione spiegata ai bambini (Piero e Nilde)", "Francesca Parmigiani; ill. Dora Creminati", "Diritti e cittadinanza"),
    ("La Resistenza spiegata ai bambini", "Francesca Parmigiani", "Resistenza e libertà"),
    ("Le emozioni spiegate alle bambine e ai bambini", "(vedi scheda)", "Emozioni"),
    ("La Shoah spiegata alle bambine e ai bambini", "(vedi scheda)", "Memoria della Shoah"),
    ("La disabilità spiegata ai bambini e alle bambine", "(vedi scheda)", "Disabilità"),
    ("Il primo soccorso spiegato alle bambine e ai bambini", "(vedi scheda)", ""),
    ("Il cinema spiegato ai bambini e alle bambine", "(vedi scheda)", ""),
    ("Il mio corpo!", "(vedi scheda)", "Corpo"),
    ("Woman Power", "(vedi scheda)", "Donne e parità"),
    ("Figli della foresta. Due vite africane rapite dal colonialismo italiano", "Igiaba Scego; ill. Chiara Abastanotti", "Colonialismo e razzismo"),
]:
    add(P, t, a, "Kids / CriticalKids", "n.d.", FN, "Divulgazione/Saggio", tema, "No", "Da verificare", BK)
for t in ("Adele Crudele (serie)", "Ariol (serie)", "Archeo Tales – Il mito del Minotauro", "Paleo Stories – Estinzione", "Come l'acqua"):
    add(P, t, "(vedi scheda)", "Kids", "n.d.", FN, "Graphic novel/Fumetto", "", "No", "Da verificare", BK)


SCONTO_REF = "Condizioni!$B$2"
LFU = "https://www.lafeltrinelli.it/libri/editori/uovonero"
PRICES = {
 ("Uovonero", "Che rabbia!"): (12.00, LFU),
 ("Uovonero", "30 giorni per capire l'autismo"): (16.00, "https://www.satellitelibri.it/schedaeditore.php?maschera=editori&chiave=00047&pintOperazione=2&tab="),
 ("Uovonero", "I quaderni di #intantofaccioqualcosa 4 – Attività con la CAA"): (8.90, "https://www.uovonero.com/libri/i-quaderni"),
 ("Uovonero", "I quaderni di #intantofaccioqualcosa 6 – Attività con la CAA 2"): (8.90, "https://www.uovonero.com/libri/i-quaderni"),
 ("Settenove", "Leggere senza stereotipi. Percorsi educativi 0-6 anni (per adulti)"): (16.00, "https://www.settenove.it/articoli/leggere-senza-stereotipi/317"),
 ("effequ", "ZAMPA"): (6.00, "https://www.effequ.it/"), ("effequ", "XY"): (6.00, "https://www.effequ.it/"),
 ("effequ", "WEB"): (6.00, "https://www.effequ.it/"), ("effequ", "VIETATO"): (6.00, "https://www.effequ.it/"),
 ("Pelledoca", "Susan Duckling e il caso del regalo di compleanno"): (11.00, "https://www.verbanonews.it/aree-geografiche/lombardia/2021/10/08/libri-per-cominciare-ad-affrontare-le-paure-in-vista-di-halloween/935987/"),
 ("Pelledoca", "Nel castello di Azzurrina"): (12.00, "https://www.macrolibrarsi.it/edizioni/_pelledoca-editore.php"),
 ("Pelledoca", "Dark web"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "Lunamadre"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "Villa Mannara"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "La Barabuia"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "Alive"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "Una lunghissima notte"): (16.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Pelledoca", "Bambina d'ombra"): (18.00, "https://www.lafeltrinelli.it/libri/editori/pelledoca-editore"),
 ("Edizioni Corsare", "Farfalle"): (18.00, "https://librerie.unicatt.it/editore-edizioni-corsare-2670.html"),
 ("Edizioni Corsare", "Brava Chiocciolina!"): (16.00, "https://www.libraccio.it/editore/edizioni-corsare/libri.html"),
 ("Edizioni Corsare", "Il bimboleone e altri bambini"): (18.00, "https://www.libraccio.it/editore/edizioni-corsare/libri.html"),
 ("Cronache Ribelli", "C'era una (ri)volta – Quaranta storie da leggere e colorare (vol. I)"): (9.00, "https://www.bookletnews.it/index.php/catalogo/cronache-ribelli"),
 ("Gallucci", "La casa nella prateria (serie)"): (13.50, "https://www.lafeltrinelli.it/libri/editori/gallucci"),
}
MER = "https://www.lameridiana.it/categorie/libri-per-una-cultura-accessibile/libri-accessibili.html"
for t, p in [("A Said piaceva il mare", 12.34), ("Il mio Afghanistan (CAA)", 12.34), ("L'albero del riccio", 15.20),
             ("La Puglia è una favola", 17.10), ("Milagro in CAA", 15.20), ("Canto di Natale", 12.34), ("Dracula", 12.34),
             ("Il diario di Anna Frank", 12.34), ("Piccole donne", 15.20), ("Giacomo di cristallo e altre storie", 12.34),
             ("Frida Kahlo pittrice coraggiosa", 14.73), ("Paul Klee pittore musicista", 14.73)]:
    PRICES[("La Meridiana", t)] = (p, MER + " (prezzo 'da' indicato sul sito, formato non specificato: verificare)")


# ---------------- TERRE DI MEZZO (catalogo 1° semestre 2026) ----------------
TP = "Terre di Mezzo"
TDM = json.load(open('tdm.json', encoding='utf-8'))
GTDM = {"Albi illustrati": "Albo illustrato", "Per i più piccoli": "Libro cartonato", "Primi lettori": "Narrativa", "Giovani lettori": "Romanzo",
        "I Tascabili": "Narrativa", "Collana Slanci": "Romanzo", "Silent book": "Silent book", "Albi per tutte le età": "Albo illustrato"}
for x in TDM:
    f = {"3-6 anni": F3, "0-3 anni": F0, "Prime letture (6-8)": FP, "Scuola media (11-13)": FM, "Adolescenza (14+)": FA, "n.d. / tutte le età": FN}[x["fascia"]]
    eta = {"Per i più piccoli": "per i più piccoli", "Albi illustrati": "albi illustrati (età indicativa)", "Primi lettori": "primi lettori",
           "Giovani lettori": "giovani lettori", "I Tascabili": "tascabili", "Collana Slanci": "Young Adult", "Silent book": "per tutte le età", "Albi per tutte le età": "per tutte le età"}[x["sezione"]]
    add(TP, x["titolo"], x["autore"], "L'Acchiappastorie – " + x["sezione"], eta, f, GTDM[x["sezione"]], "", "No", "Collana",
        "Catalogo Terre di Mezzo 1° semestre 2026, p. %d (ISBN %s)" % (x["pag"], x["isbn"]))
    ISBNS[(TP, x["titolo"])] = x["isbn"]
    PRICES[(TP, x["titolo"])] = (x["prezzo"], "Catalogo Terre di Mezzo 1° semestre 2026")

PRICES.update(PRICES_L)
for _t, _i in {"Ettore. L'uomo straordinariamente forte": "9788890860553", "Io sono così": "9788898947010", "June e Lea": "9788890860539", "Papà aspetta un bimbo!": "9788890860515",
               "Selvaggia": "9788898947102", "Il piccolo giardiniere": "9788898947119", "Mi piace Spiderman... e allora?": "9788890860560", "Ada al contrario": "9788898947423",
               "La strana storia di Cappuccetto Blu": "9788898947492", "C'è qualcosa di più noioso che essere una principessa rosa?": "9788898947560", "La nudità che male fa?": "9788898947621", "Il bosco in casa": "9788898947355"}.items():
    ISBNS[("Settenove", _t)] = _i


# ---------------- MINIBOMBO (catalogo 2026) ----------------
MBJ = json.load(open('mb.json', encoding='utf-8'))
for _x in MBJ:
    _f = {"0-3 anni": F0, "3-6 anni": F3}[_x["fascia"]]
    add("Minibombo", _x["titolo"], _x["autore"], "Catalogo Minibombo 2026", _x["eta"], _f, "Libro cartonato" if _x["cartone"] else "Albo illustrato", "", "No", "Catalogo", "Catalogo Minibombo 2026 (ISBN %s)" % _x["isbn"])
    PRICES[("Minibombo", _x["titolo"])] = (_x["prezzo"], "Catalogo Minibombo 2026")
    ISBNS[("Minibombo", _x["titolo"])] = _x["isbn"]


# ---------------- PELLEDOCA (catalogo 2025) ----------------
PDJ = json.load(open('pd.json', encoding='utf-8'))
_FM = {"3-6 anni": F3, "Prime letture (6-8)": FP, "8+": F8, "9+": F9, "10+": F10, "Scuola media (11-13)": FM, "Adolescenza (14+)": FA}
for _x in PDJ:
    add("Pelledoca", _x["titolo"], _x["autore"], _x["collana"], _x["eta"], _FM[_x["fascia"]], _x["genere"], _x["tema"], "No", "Catalogo", "Catalogo Pelledoca 2025" + (" (ISBN %s)" % _x["isbn"] if _x["isbn"] else ""))
    PRICES[("Pelledoca", _x["titolo"])] = (_x["prezzo"], "Catalogo Pelledoca 2025")
    if _x["isbn"]: ISBNS[("Pelledoca", _x["titolo"])] = _x["isbn"]


# ---------------- DALIA (dal sito daliaedizioni.it) ----------------
import sys; sys.path.insert(0, ".")
from dalia_data import DAL, LISTA
for _t, _p, _a, _c in LISTA:
    _d = DAL.get(_t)
    _eta = _d[1] if _d and _d[1] else ""
    _n = None
    if _eta == "5+": _f, _e = F3, "5+"
    elif _eta == "7+": _f, _e = FP, "7+"
    elif _eta == "8+": _f, _e = F8, "8+"
    else: _f, _e = FN, "n.d."
    _g = "Narrativa" if _c != "Guide fantastiche" else "Divulgazione/Saggio"
    add("Dalia Edizioni", _t, _a or "(vedi scheda)", _c, _e, _f, _g, (_d[3] if _d else ""), "No", "Catalogo" if _eta else "Da verificare", "https://www.daliaedizioni.it/categoria-prodotto/illustrati-kids/")
    PRICES[("Dalia Edizioni", _t)] = (_p, "Sito Dalia Edizioni")
    if _d and _d[4]: ISBNS[("Dalia Edizioni", _t)] = _d[4]

# ================= COSTRUZIONE EXCEL =================
ORDER = ["Settenove", "Uovonero", "Terre di Mezzo", "Camelozampa", "La Meridiana", "Gallucci", "Pelledoca", "Momo Edizioni", "Tunuè",
         "Cronache Ribelli", "Dalia Edizioni", "Edizioni Corsare", "Minibombo", "People", "Nero Latte", "Logos",
         "Lapis", "L'Ippocampo", "Neo Edizioni", "Tetra Edizioni", "effequ", "BeccoGiallo"]
fo = {f: i for i, f in enumerate(FASCE)}
R[:] = [_r for _r in R if not (_r[0] == "Edizioni Corsare" and _r[1].startswith("La collanina"))
        and not (_r[0] == "Dalia Edizioni" and _r[10] != "https://www.daliaedizioni.it/categoria-prodotto/illustrati-kids/" and "UVA" not in _r[1])
        and not (_r[0] == "Pelledoca" and "Catalogo Pelledoca 2025" not in _r[10])
        and not (_r[0] == "Minibombo" and _r[3] in ("Catalogo Minibombo (0-6 anni)", "Catalogo Minibombo"))
        and not (_r[0] == "Settenove" and _r[1].startswith("Cosa faremo da grandi"))]
import sys; sys.path.insert(0, ".")
from uov_data import S as _S, S_AGE as _SA, CAM_CAT as _CC
for _r in R:
    if _r[0] == "Settenove" and _r[1] in _SA:
        _r[4], _r[5], _r[9] = _SA[_r[1]][0], _SA[_r[1]][1], "Catalogo"
    if _r[0] == "Settenove" and _r[1] in _S and _S[_r[1]][0] is not None:
        PRICES[("Settenove", _r[1])] = (_S[_r[1]][0], "Sito Settenove (prezzo di copertina)")
for _r in R:
    if _r[0] == "Camelozampa" and _r[1] in _CC:
        _p, _sy, _e, _f = _CC[_r[1]]
        PRICES[("Camelozampa", _r[1])] = (_p, "Catalogo Camelozampa 2026")
        if _e: _r[4], _r[5], _r[9] = _e, _f, "Catalogo"

# ---------------- MOMO (dal sito momoedizioni.it) ----------------
import sys; sys.path.insert(0, ".")
from momo_data import MOMO, MOMO_P
import re as _r3, unicodedata as _u3
def _m(x): return _r3.sub(r"[^a-z0-9]", "", _u3.normalize("NFKD", str(x).lower().replace("’", "'")).encode("ascii", "ignore").decode())
def _momo_key(t):
    for k in MOMO_P:
        if _m(t).startswith(_m(k)): return k
    return None
MOMO_P.setdefault("Karukinka", 15)
for _row in R:
    if _row[0] != "Momo Edizioni": continue
    _k = _momo_key(_row[1])
    if not _k: continue
    PRICES[("Momo Edizioni", _row[1])] = (float(MOMO_P[_k]), "Sito Momo Edizioni")
    _d = MOMO.get(_k)
    if _d:
        _eta, _syn, _tema, _isbn, _aut = _d
        _n = _r3.search(r"\d+", _eta or "")
        if _n:
            _a = int(_n.group())
            _row[4] = _eta; _row[9] = "Catalogo"
            _row[5] = F3 if _a <= 5 else (FP if _a <= 7 else (F10 if _a <= 10 else FM))
        if _tema: _row[7] = _tema
        if _aut: _row[2] = _aut
        if _isbn: ISBNS[("Momo Edizioni", _row[1])] = _isbn

# ---------------- CORSARE (catalogo primavera 2026) ----------------
from corsare_data import C as _CC2, ADD as _CADD
def _ckey(t):
    for k in _CC2:
        if _m(t).startswith(_m(k)): return k
    return None
for _row in R:
    if _row[0] != "Edizioni Corsare": continue
    _k = _ckey(_row[1])
    if not _k: continue
    _p, _eta, _isbn, _syn, _aut, _tema = _CC2[_k]
    PRICES[("Edizioni Corsare", _row[1])] = (_p, "Catalogo Edizioni Corsare Primavera 2026")
    _row[4] = _eta; _row[9] = "Catalogo"; ISBNS[("Edizioni Corsare", _row[1])] = _isbn
    _a = int(_r3.search(r"\d+", _eta).group())
    _row[5] = F3 if _a <= 5 else (FP if _a <= 7 else (F8 if _a == 8 else (F9 if _a == 9 else (F10 if _a == 10 else (FM if _a <= 13 else FA)))))
    _row[2] = _aut or _row[2]
    if _tema: _row[7] = _tema
_FF = {"0-3 anni": F0, "3-6 anni": F3, "9+": F9}
for _t, _p, _eta, _f, _g, _isbn, _syn, _aut, _tema, _caa in _CADD:
    _nr = [ "Edizioni Corsare", _t, _aut, "Area ragazzi (albi illustrati)", _eta, _FF[_f], _g, _tema, _caa, "Catalogo", "Catalogo Edizioni Corsare Primavera 2026"]
    R.append(_nr)
    PRICES[("Edizioni Corsare", _t)] = (_p, "Catalogo Edizioni Corsare Primavera 2026")
    ISBNS[("Edizioni Corsare", _t)] = _isbn

# ---------------- BECCOGIALLO (catalogo Fall 2024) ----------------
from bg_data import BG as _BG, BG_P as _BGP, BG_ADD as _BGA
for _row in R:
    if _row[0] != "BeccoGiallo": continue
    for _k, _pv in _BGP.items():
        if _m(_row[1]).startswith(_m(_k)): PRICES[("BeccoGiallo", _row[1])] = (_pv, "Sito BeccoGiallo (prezzo di listino)")
    for _k, _d in _BG.items():
        if _m(_row[1]).startswith(_m(_k)):
            _row[4] = _d[0]; _row[9] = "Catalogo"; _row[2] = _d[2]; _row[7] = _d[3]
            _a = int(_r3.search(r"\d+", _d[0]).group())
            _row[5] = F3 if _a <= 5 else FP

for _t, _p, _eta, _aut, _tema, _syn in _BGA:
    R.append(["BeccoGiallo", _t, _aut, "Kids / CriticalKids", _eta, F3, "Divulgazione/Saggio", _tema, "No", "Collana", "https://www.beccogiallo.it/categoria-prodotto/kids/criticalkids/"])
    PRICES[("BeccoGiallo", _t)] = (_p, "Sito BeccoGiallo (prezzo di listino)")
# ---------------- L'IPPOCAMPO (catalogo autunno 2026) ----------------
from ip_data import IP as _IP
def _ipkey(t):
    for k in _IP:
        if _m(t).startswith(_m(k)): return k
    return None
for _row in R:
    if _row[0] != "L'Ippocampo": continue
    _k = _ipkey(_row[1])
    if not _k: continue
    _p, _eta, _isbn, _syn = _IP[_k]
    if _p: PRICES[("L'Ippocampo", _row[1])] = (_p, "Catalogo L'Ippocampo Autunno 2026")
    ISBNS[("L'Ippocampo", _row[1])] = _isbn
    if _eta:
        _row[4] = _eta; _row[9] = "Catalogo"
        _n = _r3.findall(r"\d+", _eta); _a = int(_n[0])
        _row[5] = FP if _a <= 7 else (F8 if _a == 8 else (F9 if _a == 9 else (F10 if _a == 10 else (FM if _a <= 13 else FA))))

# ---------------- TUNUÈ (dal sito tunue.com) ----------------
from tunue_data import TU as _TU
for _row in R:
    if _row[0] != "Tunuè": continue
    if _m(_row[1]).startswith("timothytop"):
        _row[1] = "Terry Time e il coraggio della Resistenza"; _row[3] = "Kids (fumetti)"
    for _k, _d in _TU.items():
        if _m(_row[1]).startswith(_m(_k)):
            _p, _eta, _isbn, _syn, _aut = _d
            PRICES[("Tunuè", _row[1])] = (_p, "Sito Tunué (prezzo di un volume della serie)")
            ISBNS[("Tunuè", _row[1])] = _isbn
            _row[4] = _eta; _row[9] = "Catalogo"; _row[2] = _aut
            _a = int(_r3.search(r"\d+", _eta).group()); _row[5] = FP if _a <= 7 else (F8 if _a == 8 else (F9 if _a == 9 else (F10 if _a == 10 else (FM if _a <= 13 else FA))))

# ---------------- CAMELOZAMPA (catalogo 2025 caricato) ----------------
from cz_data import CZ_ISBN as _CZI
for _row in R:
    if _row[0] != "Camelozampa": continue
    for _k, _i in _CZI.items():
        if _m(_row[1]).startswith(_m(_k)): ISBNS[("Camelozampa", _row[1])] = _i

# ---------------- UOVONERO (catalogo per librai) ----------------
from uo_cat import UC as _UC
def _uckey(t):
    k = _m(t)
    if k in _UC: return k
    return next((kk for kk in _UC if len(kk) > 8 and (kk.startswith(k) or k.startswith(kk))), None)
for _row in R:
    if _row[0] != "Uovonero": continue
    _k = _uckey(_row[1])
    if not _k: continue
    _p, _isbn, _syn = _UC[_k]
    if _p: PRICES[("Uovonero", _row[1])] = (_p, "Catalogo Uovonero per librai")
    ISBNS[("Uovonero", _row[1])] = _isbn

R.sort(key=lambda r: (ORDER.index(r[0]), fo[r[5]], r[1]))


# ---------------- SELEZIONE PER IL FLIPBOOK ----------------
import re as _re2, unicodedata as _ud
def _nn(x): return _re2.sub(r"[^a-z0-9]", "", _ud.normalize("NFKD", str(x).lower().replace("’", "'")).encode("ascii", "ignore").decode())
_TD = json.load(open('tdm.json', encoding='utf-8'))
_TD_IDX = [5, 8, 20, 40, 41, 43, 46, 47, 50, 51, 61, 63, 72, 80, 88, 104, 125, 127, 140, 150, 161, 165, 192, 217, 224, 230, 282, 290, 296, 277, 275, 328, 339, 349, 360, 362, 379, 389, 390, 398, 399, 403, 404, 405]
SELECT = {
 "Terre di Mezzo": {_nn(_TD[i]["titolo"]) for i in _TD_IDX},
 "Logos": {_nn(t) for t in ["Edmondo e la cacca misteriosa", "Pinocchio", "La Sirenetta", "Bella bambina dai capelli turchini", "Mira", "Cosa fanno le bambine?", "Cosimo", "Barbablù"]} | {"lafamigliavolpitassi1lincontro"},
 "Minibombo": {_nn(t) for t in ["Orso, buco!", "Tutino con la terra", "Tutino nella neve", "Cerchio al parco", "Il libro bianco - Ed. speciale", "Grande Gatto Piccolo Gatto", "Affamato come un lupo", "Ha visto la mia coda?", "Gabbiano più gabbiano meno", "Allarme rosso!", "Ho visto una talpa", "Storie brevi", "Il dubbio", "Un segreto importante", "Il cespuglio", "Il piccolo fienile rosso", "Gatto Nero, Gatta Bianca", "La mia famiglia è uno zoo"]},
 "Pelledoca": {_nn(t) for t in ["Susan Duckling e il caso del regalo di compleanno", "Les 3 Moustaches. Il Mistero del Tappeto Peloso", "Les 3 Moustaches. Il Mistero della Ricetta del Montblanc", "Lupi, licantropi e luna piena", "Bambina d'ombra", "Uno scrigno a doppio segreto", "Il mistero del cadavere senza testa", "Dark web", "Fuori dal branco", "Alive", "Villa Mannara", "La Barabuia", "Una lunghissima notte", "Lunamadre", "La fuga di Pat", "Il caso Elisa P.", "Uscita di sicurezza"]},
 "Dalia Edizioni": {_nn(t) for t in ["Argo e Calypso, cani da biblioteca", "Il mistero di Pyrgi. Avventura tra gli Etruschi", "Alla scoperta della libertà", "Il mio superpotere è la gentilezza", "Slò, la magia della lentezza", "Arturo e Nadir", "Il vampiro Cicorione", "Damian e la Lupa", "Febe e i gatti ruba storie", "Loris Wood verso l’isola di plastica", "UVA. Una Vita Altrove"]},
}
def _insel(r):
    if r[0] not in SELECT: return "Sì"
    if r[3].startswith("InBook"): return "Sì"
    return "Sì" if _nn(r[1]) in SELECT[r[0]] else "No"
for _r in R: _r.append(_insel(_r))

def etichette(r):
    tags = []
    if TAGF[r[5]]: tags.append(TAGF[r[5]])
    g = r[6]
    if g in ("Romanzo",): tags += ["romanzo", "narrativa"]
    elif g in ("Narrativa", "Giallo/Mistero"): tags.append("narrativa")
    elif g == "Avventura": tags += ["avventura", "narrativa"]
    if "ullismo" in r[7]: tags.append("bullismo")
    if r[8] == "Sì": tags.append("CAA")
    return ", ".join(tags)

HEAD = ["Casa editrice", "Titolo", "Autori / a cura di", "Collana", "Età (come indicata in catalogo)", "Fascia (categoria)",
        "Genere", "Tema sociale", "In CAA / simboli", "Affidabilità dell'età", "Etichette categorie richieste", "Fonte (URL)",
        "Prezzo copertina (€)", "Fonte prezzo", "Prezzo netto a libreria (€)", "Nel flipbook"]
WIDTHS = [18, 52, 34, 28, 26, 22, 22, 40, 12, 16, 34, 55, 14, 40, 16, 12]

FN_ = "Arial"
hfont = Font(name=FN_, bold=True, color="FFFFFF", size=10)
hfill = PatternFill("solid", start_color="1F3864")
bfont = Font(name=FN_, size=10)
thin = Side(style="thin", color="D9D9D9")
border = Border(top=thin, bottom=thin, left=thin, right=thin)
caa_fill = PatternFill("solid", start_color="E2EFDA")
ver_fill = PatternFill("solid", start_color="FFF2CC")

def write_table(ws, rows, start=1):
    for c, h in enumerate(HEAD, 1):
        cell = ws.cell(row=start, column=c, value=h)
        cell.font, cell.fill, cell.border = hfont, hfill, border
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows, start + 1):
        pr = PRICES.get((r[0], r[1]), (None, ""))
        vals = r[:10] + [etichette(r), r[10], pr[0], pr[1], None]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=c, value=v)
            cell.font, cell.border = bfont, border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=13).number_format = "#,##0.00"
        f = ws.cell(row=i, column=15, value=f'=IF(ISNUMBER(M{i}),ROUND(M{i}*(1-{SCONTO_REF}),2),"")')
        f.font, f.border, f.number_format = bfont, border, "#,##0.00"
        fl = ws.cell(row=i, column=16, value=r[11]); fl.font, fl.border = bfont, border
        if r[8] == "Sì":
            ws.cell(row=i, column=9).fill = caa_fill
        if r[9] == "Da verificare":
            ws.cell(row=i, column=10).fill = ver_fill
        if pr[0] is None:
            ws.cell(row=i, column=13).fill = ver_fill
    for c, w in enumerate(WIDTHS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = ws.cell(row=start + 1, column=3)
    ws.auto_filter.ref = f"A{start}:{get_column_letter(len(HEAD))}{start + len(rows)}"
    ws.row_dimensions[start].height = 32

wb = Workbook()

# --- Leggimi
ws = wb.active; ws.title = "Leggimi"
lines = [
    ("Catalogo case editrici indipendenti – libri per bambini e ragazzi", True),
    ("Dati raccolti il 11 ottobre 2026 dai siti degli editori e dagli store online (IBS, Feltrinelli, Unilibro, Librerie.coop). Per ogni riga trovi la fonte.", False),
    ("", False),
    ("Come leggerlo", True),
    ("• Nuovi fogli: 'Condizioni' (sconto libreria modificabile), 'Contatti editori' (email, telefono, distributori, aggancio per la mail) ed 'Email modello'.", False),
    ("• Prezzi: solo dove verificabili (fonte in colonna); il prezzo netto si calcola con lo sconto del foglio 'Condizioni'.", False),
    ("• 'Tutti i titoli': elenco completo con filtri (casa editrice, fascia, genere, tema sociale, CAA).", False),
    ("• Un foglio per ogni casa editrice (22), ordinati per fascia d'età.", False),
    ("• 'CAA': solo i libri in simboli / Comunicazione Aumentativa Alternativa. 'Temi sociali': solo i titoli con un tema sociale indicato. 'Riepilogo': conteggi con formule.", False),
    ("• Colonna 'Affidabilità dell'età': Catalogo = età scritta dall'editore o dal negozio per quel titolo; Collana = età della collana; Da verificare = età non trovata o dedotta.", False),
    ("• Celle gialle = da verificare sulla scheda del libro. Celle verdi = libro in CAA.", False),
    ("", False),
    ("Limiti importanti", True),
    ("• Prima, seconda e terza media sono unite nella fascia 'Scuola media (11-13)': i cataloghi usano le età, non le classi.", False),
    ("• Per gli editori con cataloghi grandi (Logos, Lapis, Minibombo, Gallucci, Uovonero) l'elenco è rappresentativo, non esaustivo.", False),
    ("• Alcuni temi sociali sono marcati 'tema noto, da verificare': non li ho trovati nella scheda dell'editore.", False),
    ("• Minibombo e Lapis non hanno edizioni CAA proprie nei cataloghi consultati: le versioni in simboli di loro titoli sono traduzioni di reti di biblioteche (InBook).", False),
    ("• Nero Latte pubblica poesia e albi illustrati (ho trovato solo un albo, senza età). Tetra e Neo Edizioni: cataloghi per adulti. Pelledoca: solo gialli e mistero, nessuna CAA. Gallucci: nessuna CAA trovata.", False),
    ("", False),
    ("Editori di libri in CAA non nella tua lista (da considerare)", True),
    ("• Officina Babùk (Babalibri + Uovonero): solo libri accessibili con testo in simboli. Fonte: lachiccaufficiostampa.it/eventi/officina-babuk-i-libri-in-simboli/", False),
    ("• Homeless Book: collana CAA diretta da Maria Caterina Minardi. Fonte: homelessbook.it/collana/caa--comunicazione-aumentativa-alternativa/32", False),
]
for i, (t, b) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t)
    c.font = Font(name=FN_, size=14 if i == 1 else 10, bold=b)
    c.alignment = Alignment(wrap_text=True, vertical="top")
ws.column_dimensions["A"].width = 140

# --- Tutti i titoli
wsA = wb.create_sheet("Tutti i titoli")
write_table(wsA, R)
n = len(R)

# --- CAA
wsC = wb.create_sheet("CAA")
caa_rows = [r for r in R if r[8] == "Sì"]
write_table(wsC, caa_rows)

# --- Temi sociali
wsT = wb.create_sheet("Temi sociali")
tem_rows = sorted([r for r in R if r[7].strip()], key=lambda r: (r[7], ORDER.index(r[0])))
write_table(wsT, tem_rows)

# --- fogli per editore
for p in ORDER:
    s = wb.create_sheet(p[:31])
    write_table(s, [r for r in R if r[0] == p])

# --- Riepilogo (formule)
wsR = wb.create_sheet("Riepilogo", 1)
wsR["A1"] = "Riepilogo: catalogo completo raccolto (tutte le righe). Nel flipbook entra solo la selezione: colonna 'Nel flipbook' = Sì"
wsR["A1"].font = Font(name=FN_, bold=True, size=12)
wsR.cell(row=3, column=1, value="Fascia").font = hfont
wsR.cell(row=3, column=1).fill = hfill
for j, p in enumerate(ORDER, 2):
    c = wsR.cell(row=3, column=j, value=p); c.font, c.fill = hfont, hfill
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
c = wsR.cell(row=3, column=len(ORDER) + 2, value="Totale"); c.font, c.fill = hfont, hfill
last = n + 1
rng_pub = f"'Tutti i titoli'!$A$2:$A${last}"
rng_f = f"'Tutti i titoli'!$F$2:$F${last}"
rng_caa = f"'Tutti i titoli'!$I$2:$I${last}"
rng_t = f"'Tutti i titoli'!$H$2:$H${last}"
for i, f in enumerate(FASCE, 4):
    wsR.cell(row=i, column=1, value=f).font = bfont
    for j in range(2, len(ORDER) + 2):
        col = get_column_letter(j)
        c = wsR.cell(row=i, column=j, value=f'=COUNTIFS({rng_pub},{col}$3,{rng_f},$A{i})')
        c.font = bfont; c.alignment = Alignment(horizontal="center")
    tc = get_column_letter(len(ORDER) + 2)
    c = wsR.cell(row=i, column=len(ORDER) + 2, value=f"=SUM(B{i}:{get_column_letter(len(ORDER)+1)}{i})"); c.font = Font(name=FN_, bold=True, size=10)
r0 = 4 + len(FASCE)
extra = [("Totale titoli", None), ("di cui in CAA / simboli", rng_caa), ("di cui con tema sociale", rng_t)]
for k, (lab, rg) in enumerate(extra):
    row = r0 + k
    wsR.cell(row=row, column=1, value=lab).font = Font(name=FN_, bold=True, size=10)
    for j in range(2, len(ORDER) + 2):
        col = get_column_letter(j)
        if k == 0:
            fml = f"=SUM({col}4:{col}{r0-1})"
        elif k == 1:
            fml = f'=COUNTIFS({rng_pub},{col}$3,{rng_caa},"Sì")'
        else:
            fml = f'=COUNTIFS({rng_pub},{col}$3,{rng_t},"<>")'
        c = wsR.cell(row=row, column=j, value=fml); c.font = Font(name=FN_, bold=True, size=10); c.alignment = Alignment(horizontal="center")
    c = wsR.cell(row=row, column=len(ORDER) + 2, value=f"=SUM(B{row}:{get_column_letter(len(ORDER)+1)}{row})"); c.font = Font(name=FN_, bold=True, size=10)
wsR.cell(row=r0 + 4, column=1, value="Nota: le righe segnaposto (es. 'Nessun titolo per ragazzi trovato', 'Fasce d'età del catalogo') sono contate come 'n.d.' e non sono libri.").font = Font(name=FN_, italic=True, size=9)
wsR.column_dimensions["A"].width = 30
for j in range(2, len(ORDER) + 3):
    wsR.column_dimensions[get_column_letter(j)].width = 13
wsR.row_dimensions[3].height = 42
wsR.freeze_panes = "B4"


# ================= FOGLI AGGIUNTIVI: Condizioni, Contatti, Email =================
wsK = wb.create_sheet("Condizioni", 2)
wsK["A1"] = "Parametri per il calcolo del prezzo netto a libreria"
wsK["A1"].font = Font(name=FN_, bold=True, size=12)
wsK["A2"] = "Sconto libreria ipotizzato (modifica questa cella)"
wsK["A2"].font = bfont
wsK["B2"] = 0.30
wsK["B2"].number_format = "0%"
wsK["B2"].font = Font(name=FN_, size=10, color="0000FF")
wsK["B2"].fill = PatternFill("solid", start_color="FFF2CC")
notes = [
    "Il 30% è un valore tipico per le librerie indipendenti, solo indicativo: lo sconto reale lo decide ogni editore o distributore.",
    "Sostituiscilo con lo sconto concordato: la colonna 'Prezzo netto a libreria (€)' di tutti i fogli si aggiorna da sola.",
    "Prezzi di copertina: inseriti solo dove ho trovato un prezzo verificabile (fonte in colonna 'Fonte prezzo'). Le celle gialle sono da compilare con il listino dell'editore.",
    "Per il conto vendita (deposito) i tempi di rendiconto, la resa e lo sconto vanno concordati per iscritto con ciascun editore.",
]
for i, t in enumerate(notes, 4):
    c = wsK.cell(row=i, column=1, value=t); c.font = bfont; c.alignment = Alignment(wrap_text=True, vertical="top")
wsK.column_dimensions["A"].width = 110
wsK.column_dimensions["B"].width = 12

CONT_HEAD = ["Casa editrice", "Sede", "Email generale", "Email commerciale / ordini / librerie", "Telefono", "Sito",
             "Distribuzione / promozione per le librerie", "Note su condizioni di vendita", "Aggancio per la mail", "Fonte (URL)"]
MELI = "Messaggerie Libri (meli.dirgen@meli.it, numero verde librerie 800 804 900)"
EMME = "Emme Promozione (info@emmepromozione.it)"
CONT = [
 ("Settenove", "Via don Minzoni 44/e, 61043 Cagli (PU)", "ordini@settenove.it", "ordini@settenove.it", "n.d.", "https://www.settenove.it",
  f"Promozione: {EMME}, tel. 02 45771. Distribuzione: {MELI}", "Nessuna condizione di conto vendita trovata: da chiedere.",
  "i vostri libri contro stereotipi e violenza di genere, e il progetto Leggere senza stereotipi", "https://www.settenove.it/spazio-librerie"),
 ("Uovonero", "Via Marazzi 12, 26013 Crema (CR)", "libri@uovonero.com", "libri@uovonero.com", "0373 500 622", "https://www.uovonero.com",
  f"Promozione: {EMME}. Distribuzione: {MELI}", "Da chiedere.",
  "i libri in simboli (CAA) e i titoli su autismo, DSA e disabilità", "https://www.uovonero.com/"),
 ("Terre di Mezzo", "Via Calatafimi 10, 20122 Milano", "segreteria@terre.it", "Referente librerie: Stefania Falbo (email sulla pagina Contatti del sito); scuole: Sabina Eleonori", "02 83242426", "https://www.terre.it",
  "Distribuzione: Messaggerie Libri (800 804 900, info@meli.it). Promozione: Emme Promozione (02 458741, segreteria@emmepromozione.it)", "Da chiedere. Ha un referente dedicato alle scuole.",
  "la collana L'Acchiappastorie per bambini e ragazzi, tra albi illustrati e narrativa", "https://www.terre.it/contatti/"),
 ("Camelozampa", "Via F.lli Cervi 39, 35043 Monselice (PD)", "info@camelozampa.com", "ordini@camelozampa.com (scuole e biblioteche, sconti per la PA); promozione@camelozampa.com (eventi)", "0429 536477", "https://www.camelozampa.com",
  f"Distribuzione: {MELI}. Promozione: {EMME}", "Per le librerie passare da distributore o promotore; vendita diretta con sconti solo per enti pubblici.",
  "i vostri romanzi e albi su salute mentale, famiglie, identità e diritti", "https://www.camelozampa.com/contatti/"),
 ("La Meridiana", "Via Sergio Fontana 10/C, 70056 Molfetta (BA)", "info@lameridiana.it", "info@lameridiana.it", "080 3971945", "https://www.lameridiana.it",
  "Vedi la pagina 'Per le librerie' sul sito", "Da chiedere.",
  "la collezione 'Per una cultura accessibile' con libri in CAA e simboli", "https://www.lameridiana.it/"),
 ("Gallucci", "Via Liberiana 17, 00185 Roma", "info@galluccieditore.com", "biz@galluccieditore.com (ordini, tel. 06 84130332)", "06 8413033", "https://www.galluccieditore.com",
  "Promozione: PDE (direzione@pde.it, 02 7529151). Distribuzione: Messaggerie Libri (customer.service@meli.it)", "Da chiedere.",
  "il vostro catalogo per fasce d'età, da 0 a Young Adult", "https://www.galluccieditore.com/it/contatti"),
 ("Pelledoca", "Redazione: Via Voghera 25, 20144 Milano; sede legale: Via Corridoni 11, 20122 Milano", "info@pelledocaeditore.it", "info@pelledocaeditore.it", "n.d.", "https://pelledocaeditore.it",
  "Distribuzione: A.L.I. Agenzia Libraria Internazionale (alilibri.it). Ufficio stampa: ufficiostampa@pelledocaeditore.it", "Da chiedere.", "i vostri gialli e thriller per ragazzi, che avvicinano alla lettura anche i lettori più riluttanti", "https://pelledocaeditore.it/chi-siamo/"),
 ("Momo Edizioni", "Viale Londra 47, 00142 Roma (indirizzo da privacy policy)", "direzione@momoedizioni.it", "direzione@momoedizioni.it", "n.d.", "https://momoedizioni.it",
  "n.d.: sul sito invitano a chiedere i libri al libraio", "Da chiedere.", "LibriMonelli e Bollicina, il libro in CAA scritto da persone con disabilità", "https://momoedizioni.it/policy/"),
 ("Tunuè", "Via Giovanni Cena 4, 04100 Latina", "info@tunue.com", "commerciale@tunue.com", "0773 661760; 351 9192281", "https://www.tunue.com",
  "n.d.", "Da chiedere.", "i vostri fumetti per ragazzi, da Kids a Young Adult, e il titolo contro il bullismo American Born Chinese", "https://www.tunue.com/contatti/"),
 ("Cronache Ribelli", "n.d.", "info@cronacheribelli.it", "ordini@cronacheribelli.it (WhatsApp ordini: 348 175 0434)", "n.d.", "https://cronacheribelli.it",
  "Vendita diretta", "Spediscono solo a librerie del circuito indipendente, SOLO in conto assoluto e con minimo 5 libri: il conto vendita non è previsto.",
  "C'era una (ri)volta e i vostri libri di storia dal basso per bambini", "https://cronacheribelli.it/pages/distribuzione"),
 ("Dalia Edizioni", "Via XX Settembre 39, 05100 Terni", "info@daliaedizioni.it", "amministrazione@daliaedizioni.it (ordini e fatturazione)", "nessun telefono: solo email", "https://www.daliaedizioni.it",
  "Mediali Diffusione (ordini@medialibridiffusione.it, 06 6627304), Fastbook, distributori regionali", "Contatto solo via email. Gestiscono direttamente alcune librerie.",
  "Dalia Ragazzi e le vostre storie illustrate; siamo tra realtà indipendenti vicine al territorio", "https://www.daliaedizioni.it/contatti/"),
 ("Edizioni Corsare", "Via Ambrosi 15, 06125 Perugia", "edizionicorsare@gmail.com", "edizionicorsare@gmail.com (WhatsApp 347 2264702)", "075 5847055", "https://www.edizionicorsare.it",
  "Distribuzione: Messaggerie Libri. Promozione: PDE", "Da chiedere.", "Storie corsare, La collanina e la narrativa YA", "https://www.umbrialibri.com/editori/edizioni-corsare/"),
 ("Minibombo", "Viale Piave 4h, 42121 Reggio Emilia", "info@minibombo.it", "info@minibombo.it", "0522 749356", "https://www.minibombo.it",
  "Promozione: PDE Promozione (Via de’ Giudei 4, Bologna, 051 352704). Distribuzione: Messaggerie Libri (02 457741)", "Da chiedere.", "i vostri libri 0-6 anni e le versioni in simboli di alcuni vostri titoli", "https://www.minibombo.it/"),
 ("People", "Via Einaudi 3, 21052 Busto Arsizio (VA)", "info@peoplepub.it", "amministrazione@peoplepub.it (rapporti con le librerie); ordini@peoplepub.it", "0331 1629312", "https://www.peoplepub.it",
  "n.d.", "Catalogo prevalentemente per adulti; la linea Kids è piccola.", "la vostra linea Kids", "https://www.peoplepub.it/contatti"),
 ("Nero Latte", "Via Pastena 1, Veroli (FR)", "nerolatteedizioni@gmail.com", "nerolatteedizioni@gmail.com", "n.d.", "https://nerolatteedizioni.com",
  "Rapporto diretto con librerie indipendenti (circa 30)", "Molto probabilmente disponibile al dialogo diretto; da chiedere.", "i vostri albi illustrati e la poesia", "https://www.facebook.com/p/nerolatteedizioni-61570596786466/"),
 ("Logos", "Strada Curtatona 5/2, 41126 Modena", "commerciale@libri.it", "commerciale@libri.it", "059 412648", "https://www.libri.it",
  "Sul sito sezione 'Distribuzione per librerie e fumetterie'", "CASA EDITRICE IN CHIUSURA: lista titoli 2026 con prezzi (PVP) ricevuta dall'editore; quantità e condizioni di acquisto dello stock da concordare con commerciale@libri.it.", "il vostro catalogo di albi, fumetti e narrativa illustrata ancora disponibile", "https://www.salonelibro.it/elenco-espositori/logosedizioni/SL62296214"),
 ("Lapis", "Via Francesco Ferrara 50, 00191 Roma", "lapis@edizionilapis.it", "commerciale.lapis@edizionilapis.it", "06 3295935", "https://www.edizionilapis.it",
  "Rete vendita: PDE (051 352 704). Distribuzione: Messaggerie Libri (800 804 900)", "Da chiedere.", "il vostro catalogo da 1 a 13 anni, dai cartonati alla narrativa", "https://www.edizionilapis.it/contatti"),
 ("L'Ippocampo", "Via Petrarca 18, 20123 Milano (altre fonti indicano Corso Magenta 69)", "info@ippocampoedizioni.it", "commerciale@ippocampoedizioni.it; ordini@ippocampoedizioni.it", "02 48513735", "https://www.ippocampoedizioni.it",
  "Rapporto diretto con le librerie: gestiscono loro promozione e fatturazione", "Condizioni di vendita da chiedere al commerciale.", "L'Ippocampo Ragazzi e la narrativa per ragazzi", "https://www.ippocampoedizioni.it/contatti/per-librai-e-rivenditori"),
 ("Neo Edizioni", "Via Giardini del Rio 15, 67031 Castel di Sangro (AQ)", "info@neoedizioni.it", "info@neoedizioni.it", "n.d.", "https://www.neoedizioni.it/neo/",
  "Pagina 'Distribuzione' e 'Librerie di fiducia' sul sito", "Catalogo per adulti.", "i titoli di narrativa adatti ai ragazzi più grandi", "https://www.neoedizioni.it/neo/"),
 ("Tetra Edizioni", "Via Genova 88, 01100 Viterbo", "n.d. (vedi pagina Contatti)", "n.d.", "n.d.", "https://www.tetraedizioni.com",
  "n.d.", "Catalogo per adulti (racconti): probabilmente non pertinente.", "—", "https://www.tetraedizioni.com/"),
 ("effequ", "n.d.", "effequ@effequ.it", "commerciale@effequ.it; librerie@effequ.it (incontri in libreria)", "06 90232210", "https://www.effequ.it",
  "Promossa da LYR e distribuita da ALI, in tutte le librerie d'Italia", "Per esigenze specifiche scrivere al commerciale.", "la collana Scatoline per bambini dai 5 anni", "https://www.effequ.it/contatti/"),
 ("BeccoGiallo", "Via Rolando da Piazzola 9, 35139 Padova", "info@beccogiallo.it", "commerciale.beccogiallo@gmail.com (acquisti diretti di librerie, associazioni, enti)", "346 5455244", "https://www.beccogiallo.it",
  "Distribuzione: Messaggerie Libri", "Acquisto diretto possibile per librerie, associazioni ed enti.", "i vostri libri 'spiegati ai bambini' (bullismo, Costituzione, disabilità)", "https://www.beccogiallo.it/contatti/"),
]
wsC2 = wb.create_sheet("Contatti editori", 3)
for c, h in enumerate(CONT_HEAD, 1):
    cell = wsC2.cell(row=1, column=c, value=h)
    cell.font, cell.fill, cell.border = hfont, hfill, border
    cell.alignment = Alignment(wrap_text=True, vertical="center")
for i, row in enumerate(CONT, 2):
    for c, v in enumerate(row, 1):
        cell = wsC2.cell(row=i, column=c, value=v)
        cell.font, cell.border = bfont, border
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        if v in ("n.d.", "Da chiedere."):
            cell.fill = ver_fill
for c, w in enumerate([18, 38, 28, 42, 20, 30, 44, 44, 44, 50], 1):
    wsC2.column_dimensions[get_column_letter(c)].width = w
wsC2.freeze_panes = "B2"
wsC2.auto_filter.ref = f"A1:J{len(CONT)+1}"

EMAIL = """OGGETTO: Proposta di collaborazione: la microlibreria sociale di Bottegart e il vostro catalogo

Gentile [Nome / Ufficio commerciale],

mi chiamo [Nome Cognome] e scrivo a nome di [Nome microlibreria], la microlibreria dell'associazione Bottegart, a [Comune].

Bottegart lavora con continuità sul nostro territorio [dal ANNO, con attività: laboratori, incontri, ...]. La microlibreria è nata con una finalità sociale: portare libri belli, inclusivi e accessibili a bambine, bambini e ragazzi, comprese le famiglie che faticano ad arrivarci. Con il supporto dei servizi sociali sosteniamo lo studio: contribuiamo all'acquisto di materiale scolastico per tutti i bambini in difficoltà socio-economiche del nostro territorio. [Se confermato: una quota del ricavato delle vendite è destinata a questo progetto.]

Scriviamo proprio a voi perché il vostro catalogo è tra quelli che vogliamo sugli scaffali: [AGGANCIO PERSONALIZZATO, vedi foglio "Contatti editori"]. Per noi questi libri sono strumenti concreti di inclusione, e li accompagniamo con letture e incontri per le famiglie e le scuole.

Sappiamo che una libreria piccola come la nostra può sembrare un rischio. Per questo proponiamo un percorso graduale e trasparente:

1. Primo ordine in conto assoluto, selezionato (circa [N] titoli), pagato entro [30] giorni dalla fattura.
2. Una seconda fornitura in conto vendita di [N] titoli per ampliare l'assortimento, a condizioni che tutelano voi:
   - rendiconto scritto ogni [60/90] giorni e saldo delle copie vendute entro [15/30] giorni dal rendiconto;
   - copie conservate con cura, esposte in uno spazio dedicato e presentate negli eventi;
   - restituzione delle copie invendute in buono stato, con spese di spedizione a nostro carico;
   - un unico referente e documentazione fiscale in regola ([C.F./P.IVA]);
   - una breve scrittura firmata con quantità, sconto e tempi concordati.

Se il conto vendita non rientra nelle vostre politiche, vi chiediamo comunque di dirci a quali condizioni potremmo lavorare insieme.

Per valutare la proposta, potete inviarci il catalogo e il listino aggiornato, le condizioni di sconto e di resa, e il nome della persona con cui parlare?

Saremmo felici di costruire con voi un rapporto stabile e di portare i vostri libri a chi, sul nostro territorio, ne ha più bisogno.

Grazie per l'attenzione. Cordiali saluti,

[Nome Cognome]
[Ruolo] - [Nome microlibreria] / Associazione Bottegart
[Indirizzo] - [Telefono] - [Email] - [Sito / social]
"""
wsE = wb.create_sheet("Email modello", 4)
wsE["A1"] = "Modello di email per proporre il conto vendita (compila le parti tra parentesi quadre)"
wsE["A1"].font = Font(name=FN_, bold=True, size=12)
c = wsE["A3"]; c.value = EMAIL; c.font = bfont; c.alignment = Alignment(wrap_text=True, vertical="top")
wsE.column_dimensions["A"].width = 130
wsE.row_dimensions[3].height = 520
wsE["A5"] = ""

out = "output/catalogo_case_editrici_indipendenti.xlsx"
wb.save(out)
json.dump({a + "|" + b: v for (a, b), v in ISBNS.items()}, open("isbn.json", "w", encoding="utf-8"), ensure_ascii=False)
print("righe:", n, "CAA:", len(caa_rows), "temi:", len(tem_rows))
