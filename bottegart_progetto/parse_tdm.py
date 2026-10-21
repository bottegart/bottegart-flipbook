import re, json
t = open('tdm.txt', encoding='utf-8').read().replace('\u200b', '')
pages = t.split('\f')
SEC = [(4, 8, "0-3 anni", "Per i più piccoli"), (9, 38, "3-6 anni", "Albi illustrati"), (39, 41, "n.d. / tutte le età", "Silent book"), (42, 47, "n.d. / tutte le età", "Albi per tutte le età"), (48, 63, "Prime letture (6-8)", "Primi lettori"),
       (64, 70, "Scuola media (11-13)", "Giovani lettori"), (71, 75, "Scuola media (11-13)", "I Tascabili"), (76, 77, "Adolescenza (14+)", "Collana Slanci")]
RC = re.compile(r'\s{3,}(\d+\s*PAG.*|[\d,]+\s*EURO.*|ISBN\s*\d+.*|USCITA.*|ANCHE .*|CARTONATO.*|HARDCOVER.*|BROSSURA.*|CA?R?T?O?N?A?T?O?.*|[\d,]+)$')
def left(l):
    l = l.rstrip()
    m = re.split(r'\s{3,}', l.strip(), maxsplit=1)
    return m[0].strip() if m else ''
out = []
for pno in range(4, 78):
    sec = next((s for s in SEC if s[0] <= pno <= s[1]), None)
    if not sec: continue
    L = pages[pno - 1].split('\n')
    for i, l in enumerate(L):
        m = re.search(r'ISBN\s*(\d{13})', l)
        if not m: continue
        # title: scan back
        j = i
        while j >= 0:
            s = left(L[j]) if L[j].strip() else ''
            if s and s.isupper() and len(s) > 3 and not re.match(r'^(NOVITÀ|ISBN|\d)', s) and 'PAG' not in s:
                break
            j -= 1
        if j < 0: continue
        title = left(L[j]); 
        a = left(L[j-1]) if j > 0 and L[j-1].strip() else ''
        # sinossi: dalla riga dopo il titolo fino alla prima riga vuota dopo ISBN
        k = j + 1; body = []
        while k < len(L):
            if k > i and not L[k].strip(): break
            if L[k].strip(): body.append(left(L[k]))
            k += 1
        txt = ' '.join(body)
        txt = re.sub(r'^NOVITÀ\s*', '', txt)
        blk = '\n'.join(L[j:k])
        pr = re.search(r'([\d]+,\d{2})\s*EURO', blk)
        pg = re.search(r'(\d+)\s*PAG', blk)
        out.append(dict(pag=pno, sezione=sec[3], fascia=sec[2], autore=a, titolo=title.title() if title.isupper() else title,
                        titolo_raw=title, prezzo=float(pr.group(1).replace(',', '.')) if pr else None,
                        pagine=int(pg.group(1)) if pg else None, isbn=m.group(1), sinossi=re.sub(r'\s+', ' ', txt).strip()))
STOP = set("il lo la i gli le un una uno e di del della che per con tra fra in da su non ma ho ha sei era mio mia".split())
def smart(t, syn):
    toks = re.findall(r"[\wÀ-ÿ’']+", syn)
    cap = {w for w in toks if w[:1].isupper()}
    res = []
    for i, w in enumerate(t.lower().split(' ')):
        base = re.sub(r"[^\wÀ-ÿ]", "", w)
        if i == 0 or (len(base) > 2 and base not in STOP and base.capitalize() in cap):
            w = re.sub(r"^([^\wÀ-ÿ]*)(\w)", lambda m: m.group(1) + m.group(2).upper(), w)
        res.append(w)
    return " ".join(res)
seen = set(); ded = []
for o in out:
    if o["isbn"] in seen: continue
    seen.add(o["isbn"]); o["titolo"] = smart(o["titolo_raw"], o["sinossi"]); ded.append(o)
out = ded
json.dump(out, open('tdm.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print(len(out))
for o in out[:3] + out[40:43] + out[-3:]:
    print(o['pag'], o['sezione'], '|', o['autore'], '|', o['titolo_raw'], '|', o['prezzo'], o['pagine'], '\n   ', o['sinossi'][:230])
