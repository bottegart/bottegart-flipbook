import fitz, re, json, base64, unicodedata
d = fitz.open('cataloghi/catalogo-camelozampa-2025.pdf')
clean = lambda s: re.sub(r'\s+', ' ', s.replace('\u00ad', '')).strip()
nrm = lambda s: re.sub(r'[^a-z0-9]', '', unicodedata.normalize('NFKD', s.lower()).encode('ascii', 'ignore').decode())
T = ["Dove sono finiti i bambini", "Isidoro e la luna", "70 trucchi per non fare il bagno", "Furia Mascherata", "Felice abbastanza", "La mia vita imprevista", "La figlia della biblioteca", "Oggi non si torna a casa", "Quell'estate con Yente", "Il gruppo", "Lemming e fiocchi di neve", "Vulcano", "Figlia del destino", "Senza titolo", "Dove nasco ogni giorno", "Il giardino di Rita", "Cielo di ciliegie", "La sorpresa della volpe", "Avrò cura di te", "Felicità è una parola semplice", "Il cortile selvaggio", "Peggy trova tutto", "Amici"]
out = {}
for pno in range(len(d)):
    pg = d[pno]; blocks = [b for b in pg.get_text('blocks') if b[4].strip()]
    imgs = [im for im in pg.get_image_info(xrefs=True) if im['xref'] > 0 and (im['bbox'][2]-im['bbox'][0]) > 100 and (im['bbox'][3]-im['bbox'][1]) > 130]
    for t in T:
        if t in out: continue
        tb = next((b for b in blocks if nrm(clean(b[4])) == nrm(t) and (b[3] - b[1]) < 40), None)
        if not tb: continue
        x0 = tb[0]
        col = [b for b in blocks if b is not tb and x0 - 6 <= b[0] <= x0 + 8 and b[1] > tb[3] - 2 and b[1] < tb[1] + 330]
        col.sort(key=lambda b: b[1])
        texts = [clean(b[4]) for b in col]
        syn = max(texts, key=len) if texts else ''
        det = next((x for x in texts if re.search(r'\bpp\.|anni', x) and 'scritto' in x or 'illustrat' in x.lower() and 'pp' in x), '')
        tag = texts[0] if texts else ''
        cand = [im for im in imgs if x0 - 5 <= im['bbox'][0] <= x0 + 330 and abs(im['bbox'][1] - tb[1]) < 70]
        cand.sort(key=lambda im: im['bbox'][0])
        img = None
        if cand:
            bb = fitz.Rect(cand[0]['bbox']); pix = pg.get_pixmap(clip=bb, dpi=110)
            img = 'data:image/jpeg;base64,' + base64.b64encode(pix.tobytes('jpeg', jpg_quality=68)).decode()
        out[t] = dict(p=pno + 1, tag=tag, syn=syn, det=det, cover=img)
json.dump(out, open('cz_out.json', 'w'), ensure_ascii=False)
for t in T:
    o = out.get(t)
    print(t, '->', 'NO' if not o else (o['p'], 'cover' if o['cover'] else 'nocover', '|', o['det'][:70], '|', o['syn'][:60]))
