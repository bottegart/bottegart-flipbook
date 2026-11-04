"""Crea output/gestionale_bottegart.html: il gestionale della libreria in un solo file HTML.

Il gestionale funziona nel browser (anche senza internet) e salva i dati nel browser stesso.
Questo script incorpora il "catalogo di riferimento" (i titoli raccolti per il flipbook, con ISBN):
quando si scansiona un ISBN già presente nel catalogo, titolo, autore, editore e prezzo si compilano da soli.

Uso (dalla cartella bottegart_progetto):  python make_gestionale.py
"""
import json
import os

import openpyxl

QUI = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(QUI, "output", "catalogo_case_editrici_indipendenti.xlsx")
TEMPLATE = os.path.join(QUI, "gestionale_template.html")
OUT = os.path.join(QUI, "output", "gestionale_bottegart.html")


def isbn13_valido(s):
    if len(s) != 13 or not s.isdigit():
        return False
    tot = sum(int(c) * (1 if i % 2 == 0 else 3) for i, c in enumerate(s[:12]))
    return (10 - tot % 10) % 10 == int(s[12])


def num(v):
    return v if isinstance(v, (int, float)) else None


def main():
    isbn_map = json.load(open(os.path.join(QUI, "isbn.json"), encoding="utf-8"))
    catalogo = {}

    # 1) Tutti i titoli raccolti (Excel)
    wb = openpyxl.load_workbook(XLSX, read_only=True)
    righe = list(wb["Tutti i titoli"].iter_rows(values_only=True))
    intest = righe[0]
    col = {nome: i for i, nome in enumerate(intest)}
    for r in righe[1:]:
        editore, titolo = r[col["Casa editrice"]], r[col["Titolo"]]
        isbn = str(isbn_map.get(f"{editore}|{titolo}") or "").replace("-", "")
        if not isbn13_valido(isbn):
            continue
        catalogo[isbn] = {
            "titolo": titolo or "",
            "autore": r[col["Autori / a cura di"]] or "",
            "editore": editore or "",
            "collana": r[col["Collana"]] or "",
            "fascia": r[col["Fascia (categoria)"]] or "",
            "prezzo": num(r[col["Prezzo copertina (€)"]]),
        }

    # 2) Selezione del flipbook: dati più puliti e corretti, hanno la precedenza
    flip = json.load(open(os.path.join(QUI, "output", "libri_flipbook.json"), encoding="utf-8"))
    for b in flip:
        isbn = str(b.get("isbn") or "").replace("-", "")
        if not isbn13_valido(isbn):
            continue
        vecchio = catalogo.get(isbn, {})
        catalogo[isbn] = {
            "titolo": b.get("titolo") or vecchio.get("titolo", ""),
            "autore": b.get("autore") or vecchio.get("autore", ""),
            "editore": b.get("editore") or vecchio.get("editore", ""),
            "collana": b.get("collana") or vecchio.get("collana", ""),
            "fascia": b.get("fascia") or vecchio.get("fascia", ""),
            "prezzo": num(b.get("prezzo")) if num(b.get("prezzo")) is not None else vecchio.get("prezzo"),
        }

    dati = json.dumps(catalogo, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
    dati = dati.replace("</", "<\\/")
    html = open(TEMPLATE, encoding="utf-8").read()
    segnaposto = "/*CATALOGO_RIFERIMENTO*/{}"
    assert segnaposto in html, "segnaposto mancante nel template"
    html = html.replace(segnaposto, dati)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Creato {OUT} con {len(catalogo)} titoli nel catalogo di riferimento")


if __name__ == "__main__":
    main()
