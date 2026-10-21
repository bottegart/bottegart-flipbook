import fitz, re, json, base64, sys
sys.path.insert(0, ".")
covers = json.load(open('covers.json'))
n0 = len(covers)
norm = lambda s: re.sub(r'[^0-9]', '', s)
def colors(pix):
    s = pix.samples; step = max(1, len(s) // 600)
    return len({s[i:i + pix.n] for i in range(0, len(s) - pix.n, pix.n * 7)})
def run(pdf, isbns, tag):
    d = fitz.open(pdf); ok = 0
    for pno in range(len(d)):
        pg = d[pno]
        blocks = pg.get_text('blocks')
        imgs = [im for im in pg.get_image_info(xrefs=True) if im['xref'] > 0]
        cand = []
        for im in imgs:
            b = fitz.Rect(im['bbox']); w, h = b.width, b.height
            if w < 45 or h < 45 or w > pg.rect.width * 0.7 or h > pg.rect.height * 0.9: continue
            if not (0.45 < w / h < 1.8): continue
            cand.append(b)
        used = set()
        for isb in isbns:
            if isb in covers: continue
            blk = next((b for b in blocks if isb in norm(b[4])), None)
            if not blk: continue
            r = fitz.Rect(blk[:4]); best = None
            for k, b in enumerate(cand):
                if k in used: continue
                dx = max(b.x0 - r.x1, r.x0 - b.x1, 0); dy = max(b.y0 - r.y1, r.y0 - b.y1, 0)
                dist = (dx * dx + dy * dy) ** 0.5
                if best is None or dist < best[0]: best = (dist, k, b)
            if best and best[0] < 120:
                pix = pg.get_pixmap(clip=best[2], dpi=100)
                if colors(pix) < 25: continue
                covers[isb] = 'data:image/jpeg;base64,' + base64.b64encode(pix.tobytes('jpeg', jpg_quality=68)).decode()
                used.add(best[1]); ok += 1
    print(tag, ok, '/', len(isbns))
mb = json.load(open('mb.json')); run('cataloghi/minibombo_catalogo_2026.pdf', [x['isbn'] for x in mb], 'Minibombo')
pd = json.load(open('pd.json')); run('cataloghi/Catalogo-Pelledoca-2025.pdf', [x['isbn'] for x in pd if x['isbn']], 'Pelledoca')
from corsare_data import C, ADD
run('cataloghi/CATALOGO-PRIMAVERA-2026-ridotto.pdf', [v[2] for v in C.values()] + [a[5] for a in ADD], 'Corsare')
from ip_data import IP
run('cataloghi/catalogo-autunno-2026_a5_layout.pdf', [v[2] for v in IP.values()], 'Ippocampo')
json.dump(covers, open('covers.json', 'w'))
print(n0, '->', len(covers))
