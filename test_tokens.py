import pdfplumber

with pdfplumber.open("DICIEMBRE.pdf") as pdf:
    for pn, page in enumerate(pdf.pages[:3]):
        words = page.extract_words(x_tolerance=3, y_tolerance=3)
        filas = {}
        for w in words:
            y = round(w["top"] / 4)
            filas.setdefault(y, []).append(w["text"])
        print(f"=== PAGINA {pn+1} ===")
        for y in sorted(filas):
            fila = filas[y]
            first = fila[0]
            if first.isdigit() and 1 <= int(first) <= 99:
                print(f"  ROW {first}: {fila}")
