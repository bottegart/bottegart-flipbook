# -*- coding: utf-8 -*-
"""Ordini delle scuole (#ioleggoperché): unisce le liste desiderata, le divide per scuola e per casa editrice.
Input (cartella ordini/, esclusa da GitHub perché contiene i dati delle scuole):
  righe_catalogo.json  richieste fatte con il catalogo online (titoli già abbinati)
  nido.json            richieste del nido con ISBN
  items.json           richieste libere (liste Word), una per riga
  ricerca.json         risultati della ricerca dei titoli su IBS / La Feltrinelli
Output: ordini/ordini_scuole.xlsx"""
import json, re, collections, importlib.util
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

D = 'ordini/'
spec = importlib.util.spec_from_file_location("ce", "contatti_editori.py")
src = open("contatti_editori.py", encoding="utf-8").read()
CONT = {}
exec(src[src.index("C = {"):src.index("CATALOGO =")], {}, CONT)
CONT = CONT["C"]

EDITORI = {"Becco Giallo": "BeccoGiallo", "minibombo": "Minibombo", "Edizioni La Meridiana": "La Meridiana", "Pelledoca Editore": "Pelledoca",
           "Tunué": "Tunuè", "L'Ippocampo": "L’Ippocampo", "L'Ippocampo Ragazzi": "L’Ippocampo", "Logos": "Logos", "Edizioni Corsare": "Edizioni Corsare",
           "Momo": "Momo Edizioni", "CR Edizioni": "Cronache Ribelli", "Dalia": "Dalia Edizioni"}
def editore(e):
    e = (e or "").replace("\\'", "’").replace("'", "’").strip()
    return EDITORI.get(e.replace("’", "'"), e)
def prezzo_copertina(c):
    try:
        return round(float(c["prezzo"]) + float(c.get("sconto") or 0), 2)
    except (TypeError, ValueError):
        return None
def titolo_pulito(t):
    return re.split(r"\. Ediz\.|\. Con |\. Nuova ediz", t)[0].replace("\\'", "’").replace("'", "’")

righe = json.load(open(D + "righe_catalogo.json", encoding="utf-8"))
righe += json.load(open(D + "nido.json", encoding="utf-8"))
I = json.load(open(D + "items.json", encoding="utf-8"))
R = json.load(open(D + "ricerca.json", encoding="utf-8"))
_sc = json.load(open(D + "scelte.json", encoding="utf-8")) if __import__("os").path.exists(D + "scelte.json") else {}
SCELTE, CLASSICI = _sc.get("scelte", {}), set(_sc.get("edizione_da_confermare", []))
R2 = json.load(open(D + "ricerca2.json", encoding="utf-8")) if __import__("os").path.exists(D + "ricerca2.json") else {}
TUTTI_ISBN = {c["isbn"]: c for v in R.values() for c in v["cands"]}
TUTTI_ISBN.update({c["isbn"]: c for v in R2.values() for c in v})
for k, it in enumerate(I):
    r = R.get(str(k), {"cands": []})
    tutti = r["cands"] + R2.get(str(k), [])
    scelta = SCELTE.get(str(k))
    if scelta == "nessuno":
        best, stato = None, "Non trovato"
    elif scelta == "eduki":
        best, stato = None, "Materiale didattico eduki (non è un libro)"
    elif isinstance(scelta, dict):
        best = dict(scelta, sconto=0, disp=scelta.get("disp", "libro del catalogo Bottegart"))
        stato = scelta.get("stato", "Trovato (scelto a mano)")
    elif scelta:
        best = next((c for c in tutti if c["isbn"] == scelta), None) or TUTTI_ISBN.get(scelta)
        stato = "Trovato (scelto a mano)" if best else "Non trovato"
    else:
        best = r["cands"][0] if r["cands"] else None
        stato = "Trovato" if best and best["score"] >= 1.0 else ("Da verificare" if best and best["score"] >= 0.6 else "Non trovato")
        if stato == "Non trovato": best = None
    if best and k in CLASSICI: stato = "Trovato (edizione da confermare)"
    righe.append(dict(scuola="Scuole di Montecastrilli e Avigliano Umbro (liste Word)", plesso=it["plesso"], classe=it["classe"], richiesta=it["testo"],
                      titolo=titolo_pulito(best["titolo"]) if best else "", autore=(best or {}).get("autore") or "", editore=editore((best or {}).get("editore")),
                      isbn=(best or {}).get("isbn", ""), copie=1, prezzo=prezzo_copertina(best) if best else None, stato=stato, fonte=it["fonte"],
                      disp=(best or {}).get("disp") or "",
                      alternative=("" if stato.startswith(("Trovato", "Fuori", "Materiale")) else "; ".join(f'{titolo_pulito(c["titolo"])} – {editore(c["editore"])} ({c["isbn"]})' for c in tutti[:3]))))

# --- regole di Bottegart sugli ordini ---------------------------------------------------------
# 1. niente autopubblicati, materiali eduki, libri fuori catalogo o non disponibili
# 2. niente classici, a meno che non li pubblichi una casa editrice del nostro catalogo:
#    in quel caso si ordina la nostra edizione
CATALOGO_LIBRI = {b["titolo"]: b for b in json.load(open("output/libri_flipbook.json", encoding="utf-8"))}
CLASSICI_RE = re.compile(r"Pinocchio|^Cuore$|diario di Anna Frank|Oliver Twist|viaggi di Gulliver|Alice nel paese|mago di Oz|giro del mondo in 80|"
    r"Dalla Terra alla Luna|Viaggio al centro della Terra|Rosso Malpelo|Canto di Natale|Piccolo Principe|Se questo è un uomo|sistema periodico|"
    r"fattoria degli animali|amico ritrovato|barone rampante|cavaliere inesistente|visconte dimezzato|sentiero dei nidi|foresta-radice|"
    r"giorno della civetta|gabbiano Jonathan|cronache di Narnia|storia infinita|Classicini|Favole al telefono|^Matilde$|fabbrica di cioccolato|"
    r"^Il GGG$|magica medicina|Lotta Combinaguai|invasione degli orsi|bambino sottovuoto|Quando Hitler rubò|tigre all.ora del tè|bruco Maisazio|"
    r"^Pezzettino|Un anno di sorprese|bruco misuratutto|^Il sentiero segreto|faccia di Leonardo|Piccole donne raccontato|lupo e i sette capretti", re.I)
NOSTRA_EDIZIONE = [(r"Pinocchio", "Logos", "Pinocchio", ""), (r"mago di Oz", "L’Ippocampo", "Il meraviglioso mago di Oz", ""),
    (r"giro del mondo in 80", "L’Ippocampo", "Il giro del mondo in 80 giorni", ""),
    (r"Anna Frank", "La Meridiana", "Il diario di Anna Frank", "versione in simboli CAA"), (r"Canto di Natale", "La Meridiana", "Canto di Natale", "versione in simboli CAA"),
    (r"Piccole donne", "La Meridiana", "Piccole donne", "versione in simboli CAA"), (r"lupo e i sette capretti", "Uovonero", "Il lupo e i sette capretti", "versione in simboli CAA")]
for r in righe:
    st, d = r["stato"], (r.get("disp") or "")
    if st.startswith(("Materiale", "Fuori")):
        r["stato"] = "Escluso: " + ("materiale eduki" if st.startswith("Materiale") else "fuori catalogo"); continue
    if r["editore"].startswith("Autopubblicato"):
        r["stato"] = "Escluso: autopubblicato"; continue
    if re.search(r"non dispon", d, re.I):
        r["stato"] = "Escluso: non disponibile"; continue
    if r["titolo"] and CLASSICI_RE.search(r["titolo"]) and r["editore"] not in CONT:
        sub = next(((e, t, n) for pat, e, t, n in NOSTRA_EDIZIONE if re.search(pat, r["titolo"], re.I)), None)
        b = CATALOGO_LIBRI.get(sub[1]) if sub else None
        if b and b["editore"] == sub[0]:
            r.update(titolo=b["titolo"] + (f" ({sub[2]})" if sub[2] else ""), autore=b["autore"], editore=b["editore"], isbn=b["isbn"], prezzo=b["prezzo"],
                     disp="libro del catalogo Bottegart", stato="Sostituito con la nostra edizione")
        else:
            r["stato"] = "Escluso: classico di altro editore"
    elif r["stato"] == "Trovato (edizione da confermare)":
        r["stato"] = "Trovato"

FN_ = "Arial"
hfont = Font(name=FN_, bold=True, color="FFFFFF", size=10); hfill = PatternFill("solid", start_color="1F3864")
bfont = Font(name=FN_, size=10); bold = Font(name=FN_, size=10, bold=True)
thin = Side(style="thin", color="D9D9D9"); border = Border(top=thin, bottom=thin, left=thin, right=thin)
subfill = PatternFill("solid", start_color="DDEBF7"); warn = PatternFill("solid", start_color="FFF2CC"); bad = PatternFill("solid", start_color="FCE4D6")

def tabella(ws, head, rows, widths, money=(), sub=None):
    for c, h in enumerate(head, 1):
        cell = ws.cell(row=1, column=c, value=h); cell.font, cell.fill, cell.border = hfont, hfill, border
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for i, r in enumerate(rows, 2):
        for c, v in enumerate(r, 1):
            cell = ws.cell(row=i, column=c, value=v); cell.font, cell.border = bfont, border
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if c in money: cell.number_format = "#,##0.00 €"
        if sub and sub(r):
            for c in range(1, len(head) + 1): ws.cell(row=i, column=c).fill = subfill; ws.cell(row=i, column=c).font = bold
    for c, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = "A2"; ws.row_dimensions[1].height = 30

# gruppi editoriali principali (un solo canale d’ordine per gli editori dello stesso gruppo)
GRUPPI = {"Mondadori": "Gruppo Mondadori", "Piemme": "Gruppo Mondadori", "Einaudi": "Gruppo Mondadori", "Rizzoli": "Gruppo Mondadori",
          "Fabbri": "Gruppo Mondadori", "Sperling & Kupfer": "Gruppo Mondadori", "Salani": "Gruppo GeMS", "Guanda": "Gruppo GeMS",
          "Giunti Editore": "Gruppo Giunti", "Bompiani": "Gruppo Giunti", "EL": "Gruppo EL", "Emme Edizioni": "Gruppo EL", "Einaudi Ragazzi": "Gruppo EL",
          "Feltrinelli": "Gruppo Feltrinelli", "Autopubblicato (Amazon)": "Acquisto su Amazon", "Autopubblicato": "Acquisto su Amazon"}
def chi(r): return r["plesso"] + (f' {r["classe"]}' if r["classe"] else "")
ordinabili = [r for r in righe if r["stato"].startswith(("Trovato", "Dal ", "Sostituito"))]
nel_catalogo = set(CONT)

wb = Workbook()
# 1. riepilogo per casa editrice
ws = wb.active; ws.title = "Riepilogo editori"
per_ed = collections.defaultdict(list)
for r in ordinabili: per_ed[r["editore"]].append(r)
rows = []
for e, rr in sorted(per_ed.items(), key=lambda x: (x[0] not in nel_catalogo, -sum(r["copie"] for r in x[1]), x[0])):
    tit = {r["isbn"] or r["titolo"] for r in rr}
    tot = sum((r["prezzo"] or 0) * r["copie"] for r in rr)
    tel, mail, ref, note, _ = CONT.get(e, ("", "", "", "", ""))
    rows.append([e, GRUPPI.get(e, "—"), "Sì" if e in nel_catalogo else "No", len(tit), sum(r["copie"] for r in rr), round(tot, 2), len({r["plesso"] for r in rr}), tel, mail, ref, note])
tabella(ws, ["Casa editrice", "Gruppo editoriale", "Nel catalogo Bottegart", "Titoli", "Copie", "Totale a prezzo di copertina", "Scuole/plessi", "Telefono", "Email", "Referente", "Note"],
        rows, [26, 18, 12, 8, 8, 14, 10, 22, 36, 26, 40], money=(6,))
n = len(rows) + 2
ws.cell(row=n, column=1, value="TOTALE").font = bold
for c, v in ((4, sum(r[3] for r in rows)), (5, sum(r[4] for r in rows)), (6, round(sum(r[5] for r in rows), 2))):
    x = ws.cell(row=n, column=c, value=v); x.font = bold
    if c == 6: x.number_format = "#,##0.00 €"
# riepilogo per canale d’ordine: editori del catalogo uno per uno, gli altri per gruppo editoriale
wsg = wb.create_sheet("Riepilogo per gruppo", 1)
can = collections.OrderedDict()
for r in rows:
    k = r[0] if r[2] == "Sì" else (r[1] if r[1] != "—" else "Altri editori indipendenti")
    c = can.setdefault(k, ["", 0, 0, 0.0, []]); c[0] = "Catalogo Bottegart" if r[2] == "Sì" else ("Autopubblicati, solo online" if r[1] == "Acquisto su Amazon" else ("Gruppo editoriale" if r[1] != "—" else "Fuori catalogo"))
    c[1] += r[3]; c[2] += r[4]; c[3] += r[5]; c[4].append(r[0])
grows = [[k, v[0], len(v[4]), v[1], v[2], round(v[3], 2), ", ".join(v[4]) if len(v[4]) > 1 or v[0] != "Catalogo Bottegart" else ""] for k, v in can.items()]
tabella(wsg, ["Canale d’ordine", "Tipo", "Editori", "Titoli", "Copie", "Totale a prezzo di copertina", "Editori compresi"], grows, [28, 18, 8, 8, 8, 14, 90], money=(6,))

# 2. ordine per casa editrice (titoli sommati tra le scuole)
ws = wb.create_sheet("Ordine per editore")
rows = []
for e, rr in sorted(per_ed.items(), key=lambda x: (x[0] not in nel_catalogo, x[0].lower())):
    g = collections.OrderedDict()
    for r in sorted(rr, key=lambda r: r["titolo"].lower()):
        g.setdefault(r["isbn"] or r["titolo"], []).append(r)
    rows.append([e, f"{len(g)} titoli", "", "", None, sum(r["copie"] for r in rr), round(sum((r["prezzo"] or 0) * r["copie"] for r in rr), 2), "subtotale casa editrice", ""])
    for k, lst in g.items():
        r0 = lst[0]; cop = sum(r["copie"] for r in lst)
        dest = collections.Counter(chi(r) for r in lst)
        rows.append([e, r0["titolo"], r0["autore"], r0["isbn"], r0["prezzo"], cop, round((r0["prezzo"] or 0) * cop, 2), "; ".join(f"{s} ({q})" for s, q in dest.items()), r0.get("disp", "")])
tabella(ws, ["Casa editrice", "Titolo", "Autore", "ISBN", "Prezzo", "Copie", "Totale", "Per chi (copie)", "Disponibilità (IBS/Feltrinelli)"], rows,
        [22, 46, 26, 16, 10, 7, 11, 60, 22], money=(5, 7), sub=lambda r: r[7] == "subtotale casa editrice")

# 3. per scuola
ws = wb.create_sheet("Per scuola")
rows = []
for s, rr in collections.OrderedDict((s, [r for r in righe if r["plesso"] == s]) for s in dict.fromkeys(r["plesso"] for r in righe)).items():
    ok = [r for r in rr if r in ordinabili]
    rows.append([s, "", f"{len(rr)} richieste", "", "", "", None, sum(r["copie"] for r in ok), round(sum((r["prezzo"] or 0) * r["copie"] for r in ok), 2), "subtotale (solo libri trovati)"])
    for r in rr:
        rows.append([s, r["classe"], r["richiesta"], r["titolo"], r["editore"], r["isbn"], r["prezzo"], r["copie"], round((r["prezzo"] or 0) * r["copie"], 2) if r["prezzo"] else None, r["stato"]])
tabella(ws, ["Scuola / plesso", "Classe", "Richiesta originale", "Libro", "Casa editrice", "ISBN", "Prezzo", "Copie", "Totale", "Stato"], rows,
        [26, 12, 48, 40, 20, 16, 10, 7, 11, 22], money=(7, 9), sub=lambda r: r[9] == "subtotale (solo libri trovati)")
for i, r in enumerate(rows, 2):
    if r[9].startswith("Sostituito"): ws.cell(row=i, column=10).fill = warn
    if r[9].startswith(("Escluso", "Non trovato")): ws.cell(row=i, column=10).fill = bad

# 4. da verificare
ws = wb.create_sheet("Esclusi")
dv = [r for r in righe if r not in ordinabili]
tabella(ws, ["Scuola / plesso", "Classe", "Richiesta originale", "Motivo", "Libro trovato", "Casa editrice", "ISBN", "Altre possibilità"],
        [[r["plesso"], r["classe"], r["richiesta"], r["stato"], r["titolo"], r["editore"], r["isbn"], r.get("alternative", "")] for r in dv],
        [24, 12, 48, 14, 36, 18, 16, 70])

wb.save(D + "ordini_scuole.xlsx")
c = collections.Counter(r["stato"] for r in righe)
print(len(righe), "richieste:", dict(c), "· editori:", len(per_ed))
