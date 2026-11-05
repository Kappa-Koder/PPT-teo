#!/usr/bin/env python3
"""Costruisce index.html (pagina completa, funziona offline) da src/giove.html.

- Inserisce i font (cartella src/fonts) come data URI, così la presentazione
  ha lo stesso aspetto anche sul PC della scuola senza internet.
- Avvolge il contenuto nello scheletro <!doctype html> ... </html>.

Uso:  python3 build.py            -> scrive index.html
      python3 build.py --fragment OUT  -> scrive solo il contenuto (senza scheletro)
"""
import base64
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent
SRC = ROOT / "src" / "giove.html"
FONTS = ROOT / "src" / "fonts"

FACES = [
    ("Bricolage Grotesque", "700 800", "normal", "bricolage.woff2"),
    ("Atkinson Hyperlegible", "400", "normal", "atkinson-400.woff2"),
    ("Atkinson Hyperlegible", "700", "normal", "atkinson-700.woff2"),
    ("Atkinson Hyperlegible Mono", "400 700", "normal", "atkinson-mono.woff2"),
]


def font_css() -> str:
    rules = []
    for family, weight, style, file in FACES:
        data = base64.b64encode((FONTS / file).read_bytes()).decode("ascii")
        rules.append(
            "@font-face{font-family:'%s';font-style:%s;font-weight:%s;font-display:swap;"
            "src:url(data:font/woff2;base64,%s) format('woff2')}" % (family, style, weight, data)
        )
    return "\n".join(rules)


def fragment() -> str:
    html = SRC.read_text(encoding="utf-8")
    if "/*FONTS*/" not in html:
        sys.exit("segnaposto /*FONTS*/ mancante in src/giove.html")
    return html.replace("/*FONTS*/", font_css(), 1)


def main() -> None:
    body = fragment()
    if len(sys.argv) == 3 and sys.argv[1] == "--fragment":
        pathlib.Path(sys.argv[2]).write_text(body, encoding="utf-8")
        print("frammento scritto in", sys.argv[2])
        return
    page = (
        "<!doctype html>\n<html lang=\"it\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
        "</head>\n<body>\n" + body + "\n</body>\n</html>\n"
    )
    # Il <title> e gli <style> stanno nel frammento: i browser li accettano anche nel body,
    # ma per pulizia spostiamo tutto ciò che precede il primo <div id="deck"> nell'head.
    head_part, sep, rest = body.partition('<div id="deck"')
    if sep:
        page = (
            "<!doctype html>\n<html lang=\"it\">\n<head>\n<meta charset=\"utf-8\">\n"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, viewport-fit=cover\">\n"
            + head_part
            + "</head>\n<body>\n"
            + sep
            + rest
            + "\n</body>\n</html>\n"
        )
    (ROOT / "index.html").write_text(page, encoding="utf-8")
    print("index.html scritto (%d KB)" % (len(page.encode("utf-8")) // 1024))


if __name__ == "__main__":
    main()
