# -*- coding: utf-8 -*-
import json, re, openpyxl
from uov_data import U, S, CAM_CAT

wb = openpyxl.load_workbook('output/catalogo_case_editrici_indipendenti.xlsx', data_only=True)
ws = wb['Tutti i titoli']
FASCE = ["0-3 anni", "3-6 anni", "Prime letture (6-8)", "8+", "9+", "10+", "Scuola media (11-13)", "Adolescenza (14+)", "n.d. / tutte le età"]
LABEL = {"n.d. / tutte le età": "Tutte le età (età da verificare)"}
BAD_START = ("Collana", "Collezione", "Catalogo", "Fasce d", "Nessun", "Leggere senza")
BAD_EXACT = {"Libri gioco cartonati", "I libri di Gek Tessaro / Il mondo di Tony Ross / Il mondo di Salina Yoon"}
GEN = {"—", ""}
NOAUTH = {"(vedi scheda)", "(vedi catalogo)", "(vedi collana)", "(vedi collane)", "n.d.", "(in progetto)", "—", "", None, "vari", "AA.VV."}

SYN = {
 "L'isola del tesoro (ed. MinaLima)": "Il giovane Jim Hawkins trova la mappa di un tesoro e salpa con pirati e marinai: una caccia al tesoro tra tradimenti e coraggio.",
 "20.000 leghe sotto i mari": "Un professore e i suoi compagni vengono catturati dal misterioso capitano Nemo e viaggiano a bordo del sottomarino Nautilus.",
 "Il giro del mondo in 80 giorni": "Phileas Fogg scommette di riuscire a fare il giro del mondo in ottanta giorni e parte con il fedele Passepartout.",
 "Il libro della giungla": "Il cucciolo d’uomo Mowgli cresce tra i lupi e impara le leggi della giungla con Baloo e Bagheera.",
 "Peter Pan (ed. MinaLima)": "Peter Pan porta Wendy e i suoi fratelli all’Isola che non c’è, tra Bimbi Sperduti, fate e il capitano Uncino.",
 "Il meraviglioso mago di Oz": "Dorothy viene portata da un ciclone nel regno di Oz e segue la strada di mattoni gialli per tornare a casa.",
 "Il giardino segreto (ed. MinaLima)": "Mary, rimasta orfana, scopre un giardino abbandonato che cambia lei e chi le sta intorno.",
 "Piccole donne": "La vita delle sorelle March tra sogni, famiglia e crescita, in versione accessibile con simboli.",
 "Dracula": "Il conte Dracula arriva dalla Transilvania a Londra e un gruppo di amici cerca di fermarlo, in versione accessibile con simboli.",
 "Canto di Natale": "L’avaro Scrooge, visitato da tre spiriti la notte di Natale, cambia il suo cuore.",
 "Il diario di Anna Frank": "Il diario scritto da Anna negli anni in cui la sua famiglia si nasconde ad Amsterdam.",
 "Un pesce sull'albero": "Ally nasconde da anni le sue difficoltà di lettura, finché un nuovo insegnante le fa scoprire che il suo cervello funziona in modo diverso.",
 "Una specie di scintilla": "Addie, undicenne autistica, lotta perché nella sua cittadina scozzese sia ricordata la condanna di alcune donne per stregoneria.",
 "Il mistero del London Eye": "Ted e la sorella Kat cercano il cugino scomparso dopo un giro sulla ruota panoramica: Ted risolve il caso con il suo modo speciale di pensare.",
 "Omero – Iliade / Odissea per ragazzi": "I poemi di Omero raccontati ai ragazzi: la guerra di Troia e il lungo ritorno di Ulisse.",
 "La casa nella prateria (serie)": "Le avventure della famiglia Ingalls, pionieri nell’America dell’Ottocento, raccontate da Laura Ingalls Wilder.",
 "Il brutto anatroccolo": "Un anatroccolo rifiutato da tutti per il suo aspetto scopre di essere un cigno. Fiaba con testo e simboli.",
 "Il lupo e i sette capretti": "Sette capretti rimasti soli in casa devono farsi furbi per non essere ingannati dal lupo. Fiaba con testo e simboli.",
 "Raperonzolo": "La ragazza dai lunghi capelli, chiusa in una torre, trova il modo di essere libera. Fiaba con testo e simboli.",
 "Il gatto con gli stivali": "Un gatto astuto fa la fortuna del più giovane figlio di un mugnaio. Fiaba con testo e simboli.",
 "I musicanti di Brema": "Quattro animali ormai anziani scoprono che insieme possono ricominciare. Fiaba con testo e simboli.",
 "Biancaneve": "Biancaneve trova rifugio dai sette nani per sfuggire all’invidia della regina. Fiaba con testo e simboli.",
 "La principessa sul pisello": "Una principessa vera si riconosce da una prova sotto venti materassi. Fiaba con testo e simboli.",
 "I vestiti nuovi dell'imperatore": "Un imperatore vanitoso, un abito che non esiste e un bambino che dice la verità. Fiaba con testo e simboli.",
 "La grande fabbrica delle parole": "In un paese dove si parla pochissimo perché le parole si devono comprare, Philéas vuole dichiarare il suo amore a Cybelle e deve procurarsi le parole giuste. Uno dei grandi successi di Terre di Mezzo.",
 "Hank Zipzer e le cascate del Niagara (vol. 1)": "Hank Zipzer è un ragazzo con la dislessia che affronta la scuola con fantasia e umorismo.",
}

def keep(r):
    pub, t, a, col, eta, f, g, tema, caa, ver, tag, src, price = r[:13]
    if len(r) > 15 and r[15] != 'Sì': return False
    if t is None or t.startswith(BAD_START) or t in BAD_EXACT or "(sezione)" in t: return False
    if pub in ("Neo Edizioni", "Tetra Edizioni"): return False
    if g in GEN or col == "Rondini": return False
    e = (eta or "").lower()
    if e.startswith("adulti") or "educatori" in e or "insegnanti" in e: return False
    return True

def fix(s):
    return s.replace("'", "’") if isinstance(s, str) else s

def synopsis(r):
    pub, t, a, col, eta, f, g, tema, caa, ver, tag, src, price = r[:13]
    if t in SYN: return SYN[t]
    parts = []
    if caa == "Sì" and g != "Libro in simboli (CAA)":
        parts.append("Libro accessibile con simboli CAA.")
    base = g if g else "Libro"
    if g == "Libro in simboli (CAA)":
        base = "Libro in simboli CAA, con testo accompagnato da simboli e accessibile anche a chi non legge in modo convenzionale"
    if eta and eta not in ("n.d.", "n.d. (collana)", "—"):
        base += f" · età: {eta}"
    parts.insert(0, base + ".")
    if tema:
        tm = re.sub(r"\s*\((tema noto[^)]*|da verificare)\)", "", tema)
        parts.append(f"Tema: {tm}.")
    if col and col not in ("Catalogo", "—"):
        parts.append(f"Collana: {col}.")
    return " ".join(parts)

TDMSYN = {x['titolo']: x['sinossi'] for x in json.load(open('tdm.json', encoding='utf-8'))}
ISB = json.load(open('isbn.json', encoding='utf-8'))
MBSYN = {x['titolo']: x['sinossi'] for x in json.load(open('mb.json', encoding='utf-8'))}
PDSYN = {x['titolo']: x['sinossi'] for x in json.load(open('pd.json', encoding='utf-8'))}
from dalia_data import DAL as DALS
from momo_data import MOMO as MOMOS
import re as _rm, unicodedata as _um
_mm = lambda x: _rm.sub(r'[^a-z0-9]', '', _um.normalize('NFKD', str(x).lower().replace('’', "'")).encode('ascii', 'ignore').decode())
from corsare_data import C as CORS, ADD as CADD
CSYN = {_mm(k): v[3] for k, v in CORS.items()}
for _a in CADD: CSYN[_mm(_a[0])] = _a[6]
from bg_data import BG as BGS, BG_ADD as BGA
from ip_data import IP as IPS
from tunue_data import TU as TUS
from cz_data import CZ_SYN
from uo_cat import UC as UCAT
books = []
for r in ws.iter_rows(min_row=2, values_only=True):
    if not keep(r): continue
    pub, t, a, col, eta, f, g, tema, caa, ver, tag, src, price = r[:13]
    au = "" if a in NOAUTH else fix(a)
    syn = synopsis(r); pr = price if isinstance(price, (int, float)) else None; real = t in SYN
    if pub == "Uovonero" and t in U:
        pr, syn = U[t]; real = True
    if pub == "Settenove" and t in S:
        pr, sx = S[t]
        if sx: syn, real = sx, True
    if pub == "L’Ippocampo" or pub == "L'Ippocampo":
        for _k, _d in IPS.items():
            if _mm(t).startswith(_mm(_k)) and _d[3]: syn, real = _d[3], True
    if pub == 'Uovonero':
        _kk = _mm(t)
        _kk = _kk if _kk in UCAT else next((x for x in UCAT if len(x) > 8 and (x.startswith(_kk) or _kk.startswith(x))), None)
        if _kk and UCAT[_kk][2]: syn, real = UCAT[_kk][2], True
    if pub == 'Tunuè':
        for _k, _d in TUS.items():
            if _mm(t).startswith(_mm(_k)): syn, real = _d[3], True
    if pub == 'BeccoGiallo':
        for _a in BGA:
            if _mm(t) == _mm(_a[0]) and _a[5]: syn, real = _a[5], True
        for _k, _d in BGS.items():
            if _mm(t).startswith(_mm(_k)): syn, real = _d[1], True
    if pub == 'Edizioni Corsare':
        for _k, _v in CSYN.items():
            if _mm(t).startswith(_k): syn, real = _v, True
    if pub == 'Momo Edizioni':
        for _k, _d in MOMOS.items():
            if _mm(t).startswith(_mm(_k)) and _d[1]:
                syn, real = _d[1], True
    if pub == 'Dalia Edizioni' and t in DALS:
        syn, real = DALS[t][2], True
    if pub == 'Pelledoca' and t in PDSYN and PDSYN[t]:
        syn, real = PDSYN[t], True
    if pub == 'Minibombo' and t in MBSYN and MBSYN[t]:
        syn, real = MBSYN[t], True
    if pub == 'Camelozampa':
        for _k, _v in CZ_SYN.items():
            if _mm(t).startswith(_mm(_k)): syn, real = _v, True
    if pub == 'Camelozampa' and t in CAM_CAT and CAM_CAT[t][1]:
        syn, real = CAM_CAT[t][1], True
    if pub == 'Terre di Mezzo' and t in TDMSYN:
        syn, real = TDMSYN[t], True
    th = re.sub(r"\s*\((tema noto[^)]*|da verificare)\)", "", tema or "")
    acc = []
    if caa == "Sì": acc.append("Testo con simboli CAA: accessibile anche a chi non legge in modo convenzionale")
    if col and "alta leggibilità" in col: acc.append("Carattere ad alta leggibilità, utile ai lettori con dislessia")
    books.append(dict(id=len(books), t=fix(t), a=au, p=fix(pub), f=f, s=fix(syn), real=real,
                      pr=pr, caa=(caa == "Sì"), e=fix(eta if eta not in ("n.d.", None) else ""),
                      isbn=ISB.get(pub + '|' + t, ''), g=fix(g or ""), c=fix(col if col not in ("Catalogo", "—", None) else ""), th=fix(th), acc=fix("; ".join(acc))))

# dati modificabili
FL = [LABEL.get(f, f) for f in FASCE]
lines = []
for b in books:
    row = dict(titolo=b["t"], autore=b["a"], editore=b["p"], fascia=LABEL.get(b["f"], b["f"]), eta=b["e"],
               prezzo=b["pr"], caa=b["caa"], sinossi=b["s"], verificata=b["real"], tipo=b["g"], collana=b["c"],
               temi=b["th"], accessibilita=b["acc"], isbn=b["isbn"], copertina="")
    lines.append("  " + json.dumps(row, ensure_ascii=False).replace("</", "<\\/") + ",")
COV = json.load(open('covers.json'))
rd = lambda n: open(f'assets/{n}.b64').read()
HTML = open('flip_template.html', encoding='utf-8').read()
for k, v in {"__FASCE__": json.dumps(FL, ensure_ascii=False), "__LIBRI__": "\n".join(lines), "__P1__": rd('poster1'),
             "__P2__": rd('poster2'), "__BART__": rd('bartegart'), "__SOSTARE__": rd('sostare'), "__COPERTINE__": json.dumps({**{b["isbn"]: COV[b["isbn"]] for b in books if b["isbn"] in COV}, **{k: v for k, v in COV.items() if k.startswith("t:")}})}.items():
    HTML = HTML.replace(k, v)
open('output/flipbook_bottegart.html', 'w', encoding='utf-8').write(HTML)
print(len(books), "libri;", len(HTML) // 1024, "KB")
