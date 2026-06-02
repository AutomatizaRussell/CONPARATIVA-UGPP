import pdfplumber, sys

ruta = sys.argv[1]
with pdfplumber.open(ruta) as pdf:
    pag = pdf.pages[0]
    print("=== extract_text (primeras 20 lineas) ===")
    txt = pag.extract_text() or ""
    for i, linea in enumerate(txt.split("\n")[:20]):
        print(repr(linea))
    print()
    print("=== extract_words (primeras 15 palabras) ===")
    words = pag.extract_words()
    for w in words[:15]:
        print(w["text"], "  x0=", round(w["x0"]), "top=", round(w["top"]))
