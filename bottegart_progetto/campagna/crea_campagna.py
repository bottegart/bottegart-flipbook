# -*- coding: utf-8 -*-
"""Materiali della campagna #ioleggoperché 2026 della Microlibreria So’Stare.
Crea in campagna/output/: post e infografica per Instagram/Facebook, storia, banner, locandina A4 (PDF e PNG).
Uso (da bottegart_progetto):  python campagna/crea_campagna.py
Testi, date e orari sono nel blocco INFO qui sotto: cambiali lì e rilancia."""
import base64, os, asyncio
from playwright.async_api import async_playwright

INFO = dict(
    titolo="Dona un libro<br>alla tua classe",
    sottotitolo="Rendi viva la biblioteca della scuola!",
    giorni=[("sab", "7"), ("dom", "8"), ("sab", "14"), ("dom", "15")],
    mese="novembre",
    orari="ore 10–13 e 16–19",
    luogo="Microlibreria So’Stare",
    indirizzo="Configni Centro, 8",
    chi="Associazione Bottegart",
)
D = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(D, "output"); os.makedirs(OUT, exist_ok=True)
def b64(path, mime):
    return f"data:{mime};base64," + base64.b64encode(open(path, "rb").read()).decode()
IMG = {k: b64(os.path.join(D, "img", f), m) for k, f, m in [
    ("logo", "logo_ioleggoperche_2026_orizzontale.jpg", "image/jpeg"), ("comunita", "ill_comunita.jpg", "image/jpeg"),
    ("lettrici", "ill_lettrici.jpg", "image/jpeg"), ("bart", "logo-bottegart.png", "image/png"), ("sostare", "logo-sostare.png", "image/png")]}
F = os.path.join(D, "..", "assets", "fonts")
FONTS = "\n".join(f"@font-face{{font-family:'{n}';font-weight:{w};src:url({b64(os.path.join(F, f), 'font/woff2')}) format('woff2')}}"
                  for n, f, w in [("Gochi Hand", "GochiHand-normal.woff2", "400"), ("Bricolage Grotesque", "BricolageGrotesque-normal.woff2", "200 800")])

PIN = '<svg viewBox="0 0 24 24" width="1em" height="1em" aria-hidden="true"><path fill="currentColor" d="M12 2a7 7 0 0 0-7 7c0 5.2 7 13 7 13s7-7.8 7-13a7 7 0 0 0-7-7zm0 9.5A2.5 2.5 0 1 1 12 6.5a2.5 2.5 0 0 1 0 5z"/></svg>'
CLOCK = '<svg viewBox="0 0 24 24" width="1em" height="1em" aria-hidden="true"><circle cx="12" cy="12" r="9.5" fill="none" stroke="currentColor" stroke-width="2.4"/><path d="M12 6.5V12l3.8 2.4" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"/></svg>'

CSS = FONTS + """
*{box-sizing:border-box;margin:0;padding:0}
:root{--rosso:#f2141b;--nero:#141414;--verde:#d9e98a;--giallo:#ffd02a;--carta:#ffffff;--blu:#1f8fd6}
body{background:var(--carta);color:var(--nero);font-family:'Bricolage Grotesque',Arial,sans-serif;-webkit-font-smoothing:antialiased}
.f{position:relative;overflow:hidden;background:var(--carta);display:flex;flex-direction:column;align-items:center;text-align:center}
.mano{font-family:'Gochi Hand','Comic Sans MS',cursive;font-weight:400;line-height:.95}
h1{font-family:'Gochi Hand',cursive;font-weight:400;line-height:.92;letter-spacing:-.01em}
h1 em{font-style:normal;color:var(--rosso)}
.sub{font-family:'Gochi Hand',cursive;color:var(--rosso)}
.ill{display:block;object-fit:contain}
.giorni{display:flex;gap:var(--g);justify-content:center}
.g{background:var(--rosso);color:#fff;border-radius:var(--r);display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:0 .25em 0 #9c0a0f}
.g b{font-weight:800;line-height:.9;letter-spacing:-.03em}
.g span{text-transform:uppercase;font-weight:700;letter-spacing:.12em}
.mese{font-family:'Gochi Hand',cursive}
.riga{display:flex;align-items:center;justify-content:center;gap:.35em;font-weight:700}
.riga svg{color:var(--rosso);flex:none}
.loghi{display:flex;align-items:center;justify-content:center;gap:var(--lg)}
.loghi img{border-radius:50%;background:#fff;box-shadow:0 0 0 2px #eee}
.loghi p{text-align:left;font-weight:600;line-height:1.2}
.loghi p small{display:block;font-weight:500;opacity:.7}
.banda{position:absolute;left:0;right:0;height:var(--bh);background:repeating-linear-gradient(90deg,var(--rosso) 0 50%,var(--nero) 50% 100%);background-size:var(--bw) 100%}
"""

def blocco_giorni(scala):
    return '<div class="giorni">' + "".join(f'<div class="g"><span>{a}</span><b>{n}</b></div>' for a, n in INFO["giorni"]) + "</div>"

def loghi():
    return (f'<div class="loghi"><img src="{IMG["bart"]}" alt="Associazione Bottegart"><img src="{IMG["sostare"]}" alt="So’Stare">'
            f'<p>{INFO["luogo"]}<small>{INFO["chi"]}</small></p></div>')

def pagina(w, h, css, corpo):
    return f'<!doctype html><html lang="it"><head><meta charset="utf-8"><style>{CSS}{css}</style></head><body><div class="f" style="width:{w}px;height:{h}px">{corpo}</div></body></html>'

I = INFO
# 1. post Instagram / Facebook 1080x1350
POST = pagina(1080, 1350, """
.f{padding:56px 60px 48px;--g:18px;--r:26px;--lg:22px;--bh:16px;--bw:60px}
.banda{top:0}.logo{width:560px;margin-top:10px}
h1{font-size:104px;margin-top:14px}.sub{font-size:48px;margin-top:4px}
.ill{width:880px;height:372px;margin-top:8px}
.g{width:132px;height:132px}.g b{font-size:66px}.g span{font-size:20px}
.mese{font-size:40px;margin-top:4px}
.riga{font-size:32px}.dove{margin-top:8px}
.loghi{margin-top:auto}.loghi img{width:78px;height:78px}.loghi p{font-size:24px}
""", f"""<div class="banda"></div><img class="logo" src="{IMG['logo']}" alt="#ioleggoperché 2026">
<h1>Dona <em>un libro</em><br>alla tua classe</h1><p class="sub">{I['sottotitolo']}</p>
<img class="ill" src="{IMG['comunita']}" alt="">
{blocco_giorni(1)}<p class="mese">{I['mese']}</p>
<p class="riga">{CLOCK} {I['orari']}</p><p class="riga dove">{PIN} {I['luogo']} · {I['indirizzo']}</p>
{loghi()}""")

# 2. storia 1080x1920
STORIA = pagina(1080, 1920, """
.f{padding:120px 70px 110px;--g:20px;--r:30px;--lg:24px;--bh:22px;--bw:80px}
.banda{top:0}.b2{top:auto;bottom:0}.logo{width:700px}
h1{font-size:146px;margin-top:36px}.sub{font-size:54px;margin-top:12px}
.ill{width:900px;height:540px;margin-top:16px}
.g{width:190px;height:190px}.g b{font-size:96px}.g span{font-size:28px}
.mese{font-size:56px;margin-top:8px}
.riga{font-size:44px;margin-top:6px}.dove{margin-top:16px}
.loghi{margin-top:auto}.loghi img{width:118px;height:118px}.loghi p{font-size:34px}
""", f"""<div class="banda"></div><div class="banda b2"></div><img class="logo" src="{IMG['logo']}" alt="#ioleggoperché 2026">
<h1>Dona <em>un libro</em><br>alla tua classe</h1><p class="sub">{I['sottotitolo']}</p>
<img class="ill" src="{IMG['lettrici']}" alt="">
{blocco_giorni(1)}<p class="mese">{I['mese']}</p>
<p class="riga">{CLOCK} {I['orari']}</p><p class="riga dove">{PIN} {I['luogo']}</p><p class="riga">{I['indirizzo']}</p>
{loghi()}""")

# 3. infografica "Come funziona" 1080x1350
PASSI = [("Vieni a trovarci", f"sabato e domenica 7-8 e 14-15 {I['mese']}, {I['orari']}, alla {I['luogo']}."),
         ("Scegli un libro", "dalla lista dei desideri della classe di tuo figlio o di tua figlia, oppure lasciati consigliare: abbiamo libri per tutte le età, anche in CAA."),
         ("Scrivi una dedica", "lo acquisti e lo lasci a noi, con il tuo messaggio per bambine, bambini e insegnanti."),
         ("Lo portiamo a scuola", "a fine campagna consegniamo tutti i libri donati: la biblioteca della classe cresce!")]
ICONE = ['<svg viewBox="0 0 64 64"><path d="M10 28 32 12l22 16v24H10z" fill="#fff" stroke="#141414" stroke-width="4" stroke-linejoin="round"/><rect x="26" y="36" width="12" height="16" fill="#f2141b"/></svg>',
         '<svg viewBox="0 0 64 64"><path d="M8 14h20c3 0 4 2 4 4v34c0-2-2-4-4-4H8z" fill="#fff" stroke="#141414" stroke-width="4" stroke-linejoin="round"/><path d="M56 14H36c-3 0-4 2-4 4v34c0-2 2-4 4-4h20z" fill="#ffd02a" stroke="#141414" stroke-width="4" stroke-linejoin="round"/></svg>',
         '<svg viewBox="0 0 64 64"><path d="M32 54S8 40 8 24a12 12 0 0 1 24-4 12 12 0 0 1 24 4c0 16-24 30-24 30z" fill="#f2141b" stroke="#141414" stroke-width="4" stroke-linejoin="round"/></svg>',
         '<svg viewBox="0 0 64 64"><rect x="8" y="12" width="12" height="40" fill="#f2141b" stroke="#141414" stroke-width="4"/><rect x="22" y="18" width="10" height="34" fill="#ffd02a" stroke="#141414" stroke-width="4"/><rect x="34" y="10" width="10" height="42" fill="#1f8fd6" stroke="#141414" stroke-width="4"/><path d="M46 20l9-2 6 33-9 2z" fill="#d9e98a" stroke="#141414" stroke-width="4" stroke-linejoin="round"/></svg>']
INFO_G = pagina(1080, 1350, """
.f{padding:60px 70px 50px;--lg:20px;--bh:16px;--bw:60px;text-align:left}
.banda{top:0}.logo{width:440px;align-self:center;margin-top:10px}
h1{font-size:96px;text-align:center;margin-top:16px}
.passi{list-style:none;display:flex;flex-direction:column;gap:22px;margin-top:30px;width:100%}
.passi li{display:grid;grid-template-columns:96px 1fr;gap:26px;align-items:center;background:#fff;border:3px solid var(--nero);border-radius:28px;padding:22px 28px;box-shadow:0 7px 0 var(--nero)}
.ico{position:relative;width:96px;height:96px}.ico svg{width:96px;height:96px}
.ico i{position:absolute;top:-14px;left:-14px;width:46px;height:46px;border-radius:50%;background:var(--rosso);color:#fff;font-style:normal;font-weight:800;font-size:28px;display:flex;align-items:center;justify-content:center;border:3px solid var(--nero)}
.passi b{font-family:'Gochi Hand',cursive;font-weight:400;font-size:50px;line-height:1;display:block}
.passi p{font-size:27px;line-height:1.28;margin-top:4px}
.nota{margin-top:26px;font-size:26px;text-align:center;align-self:center;max-width:880px}
.nota strong{color:var(--rosso)}
.loghi{margin-top:auto;align-self:center}.loghi img{width:80px;height:80px}.loghi p{font-size:24px}
""", f"""<div class="banda"></div><img class="logo" src="{IMG['logo']}" alt="#ioleggoperché 2026">
<h1>Come <em>funziona</em></h1>
<ol class="passi">""" + "".join(f'<li><div class="ico">{ICONE[k]}<i>{k+1}</i></div><div><b>{t}</b><p>{p}</p></div></li>' for k, (t, p) in enumerate(PASSI)) + f"""</ol>
<p class="nota"><strong>Lo sapevi?</strong> Da quest’anno partecipano anche gli asili nido!</p>
{loghi()}""")

# 4. banner 1200x630 (evento Facebook, anteprima link WhatsApp)
BANNER = pagina(1200, 630, """
.f{flex-direction:row;align-items:stretch;text-align:left;--g:12px;--r:18px;--lg:14px;--bh:12px;--bw:48px}
.banda{top:0}.sx{width:610px;padding:44px 20px 30px 52px;display:flex;flex-direction:column}
.logo{width:380px}h1{font-size:76px;margin-top:14px}.sub{font-size:34px;margin-top:4px}
.giorni{justify-content:flex-start;margin-top:auto}.g{width:96px;height:96px}.g b{font-size:48px}.g span{font-size:15px}
.mese{font-size:30px;margin-top:2px}.riga{font-size:22px;justify-content:flex-start}
.dx{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:30px 26px 22px 0}
.ill{width:560px;height:430px}.loghi img{width:56px;height:56px}.loghi p{font-size:18px}
""", f"""<div class="banda"></div><div class="sx"><img class="logo" src="{IMG['logo']}" alt="#ioleggoperché 2026">
<h1>Dona <em>un libro</em><br>alla tua classe</h1><p class="sub">{I['sottotitolo']}</p>
{blocco_giorni(1)}<p class="mese" style="text-align:left">{I['mese']} · {I['orari']}</p><p class="riga">{PIN} {I['luogo']} · {I['indirizzo']}</p></div>
<div class="dx"><img class="ill" src="{IMG['comunita']}" alt="">{loghi()}</div>""")

# 5. locandina A4 (210x297 mm = 794x1123 px a 96 dpi)
A4 = pagina(794, 1123, """
@page{size:A4;margin:0}
.f{padding:52px 50px 40px;--g:14px;--r:20px;--lg:16px;--bh:12px;--bw:44px}
.banda{top:0}.b2{top:auto;bottom:0}.logo{width:430px}
h1{font-size:82px;margin-top:12px}.sub{font-size:36px;margin-top:4px}
.ill{width:660px;height:318px;margin-top:8px}
.g{width:104px;height:104px}.g b{font-size:52px}.g span{font-size:15px}
.mese{font-size:32px;margin-top:2px}
.riga{font-size:25px}.dove{margin-top:4px}
.loghi{margin-top:auto}.loghi img{width:70px;height:70px}.loghi p{font-size:20px}
""", f"""<div class="banda"></div><div class="banda b2"></div><img class="logo" src="{IMG['logo']}" alt="#ioleggoperché 2026">
<h1>Dona <em>un libro</em><br>alla tua classe</h1><p class="sub">{I['sottotitolo']}</p>
<img class="ill" src="{IMG['comunita']}" alt="">
{blocco_giorni(1)}<p class="mese">{I['mese']}</p>
<p class="riga">{CLOCK} {I['orari']}</p><p class="riga dove">{PIN} {I['luogo']} · {I['indirizzo']}</p>
{loghi()}""")

async def main():
    async with async_playwright() as p:
        b = await p.chromium.launch(executable_path=os.environ.get("CHROMIUM", "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"))
        for nome, html, w, h, scala in [("1_post_instagram_facebook", POST, 1080, 1350, 1), ("2_storia_instagram_whatsapp", STORIA, 1080, 1920, 1),
                                        ("3_infografica_come_funziona", INFO_G, 1080, 1350, 1), ("4_banner_evento_facebook", BANNER, 1200, 630, 1),
                                        ("5_locandina_A4", A4, 794, 1123, 3)]:
            pg = await b.new_page(viewport={"width": w, "height": h}, device_scale_factor=scala)
            await pg.set_content(html); await pg.wait_for_timeout(500)
            await pg.screenshot(path=os.path.join(OUT, nome + ".png"), clip={"x": 0, "y": 0, "width": w, "height": h})
            if nome.startswith("5_"):
                await pg.pdf(path=os.path.join(OUT, nome + ".pdf"), width="210mm", height="297mm", print_background=True)
            open(os.path.join(OUT, nome + ".html"), "w", encoding="utf-8").write(html) if False else None
            await pg.close()
        await b.close()
asyncio.run(main())
print("Creati in", OUT, sorted(os.listdir(OUT)))
