import os
import json
import pandas as pd


CARPETA_DATOS = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Datos")


def detectar_columna_documento(df, nombre_archivo):
    """
    Retorna el nombre de la columna que contiene el número de documento.
    - facturas   -> columna A (índice 0), siempre llamada 'Cédula'
    - estructurado -> columna C (índice 2); tras el split queda como '*- Número'
    """
    if "_facturas" in nombre_archivo.lower():
        return df.columns[0]

    # estructurado: buscar la columna que termina en '- Número' (resultado del split CC/CE)
    for col in df.columns:
        if str(col).strip().endswith("- Número"):
            return col

    # fallback: columna C posicional
    if len(df.columns) >= 3:
        return df.columns[2]

    return df.columns[0]


def agrupar_por_documento(df, col_doc):
    """
    Agrupa las filas del DataFrame por número de documento.
    Devuelve un dict  { "1002727": { "documento": "...", "registros": [...] } }
    """
    grupos = {}
    for _, fila in df.iterrows():
        doc = str(fila[col_doc]).strip()
        if not doc or doc.lower() in ("nan", "", "none", "no detectado"):
            continue

        fila_dict = {k: ("" if str(v).lower() in ("nan", "none") else str(v).strip())
                     for k, v in fila.items()}

        if doc not in grupos:
            grupos[doc] = {"documento": doc, "registros": []}

        grupos[doc]["registros"].append(fila_dict)

    return grupos


def procesar_excel(ruta_excel):
    """
    Lee todas las hojas de un Excel y devuelve un dict agrupado por documento.
    Si tiene varias hojas, cada hoja mantiene su propio sub-dict dentro de 'hojas'.
    """
    nombre_archivo = os.path.basename(ruta_excel)
    xl = pd.ExcelFile(ruta_excel)
    hojas_resultado = {}

    for hoja in xl.sheet_names:
        df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)
        df = df.fillna("")

        if df.empty:
            continue

        col_doc = detectar_columna_documento(df, nombre_archivo)
        grupos = agrupar_por_documento(df, col_doc)

        if grupos:
            hojas_resultado[hoja] = grupos

    # Si solo hay una hoja, aplanar (no agregar nivel extra innecesario)
    if len(hojas_resultado) == 1:
        return list(hojas_resultado.values())[0]

    return hojas_resultado


def convertir_todos():
    if not os.path.exists(CARPETA_DATOS):
        print(f"La carpeta Datos no existe en: {CARPETA_DATOS}")
        return

    archivos_excel = [f for f in os.listdir(CARPETA_DATOS) if f.lower().endswith(".xlsx")]

    if not archivos_excel:
        print("No hay archivos Excel en la carpeta Datos.")
        return

    for nombre_excel in archivos_excel:
        ruta_excel = os.path.join(CARPETA_DATOS, nombre_excel)
        print(f"Procesando: {nombre_excel}")

        try:
            datos = procesar_excel(ruta_excel)
            nombre_json = os.path.splitext(nombre_excel)[0] + ".json"
            ruta_json = os.path.join(CARPETA_DATOS, nombre_json)

            with open(ruta_json, "w", encoding="utf-8") as f:
                json.dump(datos, f, ensure_ascii=False, indent=2)

            print(f"  -> {nombre_json}  ({len(datos)} usuarios únicos)")

        except Exception as e:
            print(f"  ERROR procesando {nombre_excel}: {e}")

    print("\nConversión completada.")


if __name__ == "__main__":
    convertir_todos()
