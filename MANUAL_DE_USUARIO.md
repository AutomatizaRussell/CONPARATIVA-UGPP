# Manual de Usuario
## Convertidor de Planillas PDF a Excel

---

## ¿Para qué sirve este programa?

Este programa toma los archivos PDF de planillas de seguridad social y facturas de nómina, y los convierte automáticamente en archivos Excel organizados y con formato profesional. Usted no necesita copiar ni pegar nada a mano.

---

## ¿Cómo abrir el programa?

Haga doble clic en el acceso directo del programa. Verá que aparece una ventana negra de fondo (consola) — es normal, no la cierre. El programa funciona a través de ventanas emergentes que aparecen sobre esa consola.

---

## Paso 1 — Convertir la Planilla UGPP

**Use esta opción cuando tenga el PDF de la planilla de seguridad social de la UGPP.**

Al ejecutar esta opción, aparece una ventana para buscar el archivo:

```
┌─────────────────────────────────────────────┐
│  Selecciona el archivo PDF para segmentar   │
│                                             │
│  Buscar en: [  Mis Documentos          ▼ ] │
│                                             │
│  [ ] ENERO_planilla.pdf                     │
│  [ ] FEBRERO_planilla.pdf                   │
│  [●] MARZO_planilla.pdf                     │
│                                             │
│  Nombre: MARZO_planilla.pdf                 │
│  Tipo:   Archivos PDF (*.pdf)               │
│                                             │
│              [ Abrir ]  [ Cancelar ]        │
└─────────────────────────────────────────────┘
```

**Qué hacer:**
1. Navegue hasta donde guardó el PDF de la planilla.
2. Haga clic sobre el archivo para seleccionarlo.
3. Haga clic en **Abrir**.

El programa trabaja automáticamente. Cuando termina, aparece este mensaje:

```
┌─────────────────────────────────────────┐
│  ✓  ¡Éxito!                            │
│                                         │
│  Se procesaron 3 tablas.                │
│                                         │
│  El archivo se guardó con diseño de     │
│  tabla profesional, fuentes             │
│  estandarizadas, efecto cebra y         │
│  autoajuste de celdas.                  │
│                                         │
│              [  Aceptar  ]              │
└─────────────────────────────────────────┘
```

Haga clic en **Aceptar**. El archivo Excel quedó guardado automáticamente en la carpeta **Datos** dentro del programa. Se llama igual que su PDF pero con `_estructurado.xlsx` al final.

> **Ejemplo:** Si su PDF se llama `MARZO.pdf`, el Excel se llama `MARZO_estructurado.xlsx`.

---

## Paso 2 — Convertir las Facturas de Nómina

**Use esta opción cuando tenga el PDF con las facturas individuales de cada empleado.**

Al ejecutar, aparece la misma ventana de selección de archivo:

```
┌─────────────────────────────────────────────┐
│       Selecciona el PDF de la Factura       │
│                                             │
│  Nombre: MARZO_facturas.pdf                 │
│  Tipo:   Archivos PDF (*.pdf)               │
│                                             │
│              [ Abrir ]  [ Cancelar ]        │
└─────────────────────────────────────────────┘
```

1. Seleccione el PDF de facturas del mismo mes.
2. Haga clic en **Abrir**.

Cuando termina, aparece el mensaje de éxito y **el Excel se abre automáticamente** en su pantalla para que pueda revisarlo de inmediato.

El archivo guardado se llama igual que su PDF pero con `_facturas.xlsx` al final.

> **Ejemplo:** `MARZO_facturas.pdf` → `MARZO_facturas.xlsx`

---

## Paso 3 — Convertir los Excel a formato interno (JSON)

**Este paso prepara los datos para el consolidado final. No requiere que usted haga nada más que esperar.**

Al ejecutar esta opción, el programa trabaja en silencio y muestra en la consola algo como esto:

```
Procesando: MARZO_estructurado.xlsx
  -> MARZO_estructurado.json  (45 usuarios únicos)
Procesando: MARZO_facturas.xlsx
  -> MARZO_facturas.json  (45 usuarios únicos)
Conversión completada.
```

Cuando aparezca `Conversión completada.` el paso terminó. No se abre ninguna ventana adicional.

---

## Paso 4 — Generar el Excel Consolidado Final

**Este es el resultado principal: un único Excel con todos los datos de aportes y pagos del mes por persona.**

Al ejecutar, el programa trabaja solo y muestra en consola:

```
Procesando: MARZO
  -> MARZO_CONSOLIDADO.xlsx  (45 personas)
Listo.
```

El archivo `MARZO_CONSOLIDADO.xlsx` queda guardado en la carpeta **Datos**. Ábralo desde allí con Excel.

### ¿Qué contiene el consolidado?

El Excel tiene una fila por cada persona con las siguientes columnas:

| Columna | Qué contiene |
|---------|-------------|
| A — TIPO | Tipo de documento (CC, CE, NIT…) |
| B — NÚMERO | Número de identificación |
| C — NOMBRE | Nombre completo del empleado |
| D — SALARIO MES | Salario básico del mes |
| E — BONO NO SALARIO | Bonos no constitutivos de salario |
| G — AUX. ROD. | Auxilio de rodamiento |
| H — PENSIÓN | Aporte pensión del trabajador |
| I — SALUD | Aporte salud del trabajador |
| J — FSP | Fondo de Solidaridad Pensional |
| K — PENSIÓN empresa | Aporte pensión de la empresa |
| L — FSP empresa | FSP a cargo de la empresa |
| M — SALUD empresa | Aporte salud de la empresa |
| N — ARL | Aporte a riesgos laborales |
| O — CCF | Caja de Compensación Familiar |
| P — SENA | Aporte al SENA |
| Q — ICBF | Aporte al ICBF |
| R — OBSERVACIONES | Espacio libre para notas manuales |

---

## Paso opcional — Resumen de Aportes en Línea

**Use esta opción si tiene el PDF de resumen mensual de Aportes en Línea (diferente a la planilla UGPP).**

Funciona igual que el Paso 1: seleccione el PDF, haga clic en **Abrir** y espere el mensaje de éxito. El archivo generado se llama `NOMBRE_PLANILLA.xlsx`.

---

## ¿Dónde quedan los archivos?

Todos los archivos generados se guardan automáticamente en la carpeta **Datos**, que está dentro de la carpeta del programa:

```
Datos/
├── MARZO_estructurado.xlsx     ← Resultado Paso 1
├── MARZO_facturas.xlsx         ← Resultado Paso 2
├── MARZO_estructurado.json     ← Resultado Paso 3 (uso interno)
├── MARZO_facturas.json         ← Resultado Paso 3 (uso interno)
└── MARZO_CONSOLIDADO.xlsx      ← Resultado Paso 4 (el que usa usted)
```

No mueva ni borre los archivos `.json` hasta que haya generado el consolidado del mes.

---

## Problemas frecuentes

### Aparece este mensaje de advertencia:

```
┌───────────────────────────────────────────┐
│  ⚠  Aviso                                │
│                                           │
│  No se encontraron los patrones de        │
│  inicio/fin en las tablas.                │
│                                           │
│              [  Aceptar  ]               │
└───────────────────────────────────────────┘
```

**Qué significa:** El PDF seleccionado no tiene el formato de planilla UGPP esperado, o puede ser un comprobante individual en lugar del archivo completo del mes.

**Qué hacer:** Verifique que seleccionó el PDF correcto. Los PDF escaneados (fotografías de documentos) no pueden procesarse.

---

### Aparece este mensaje de error:

```
┌───────────────────────────────────────────┐
│  ✕  Error Crítico                        │
│                                           │
│  Ocurrió un error inesperado:             │
│  [descripción técnica del error]          │
│                                           │
│              [  Aceptar  ]               │
└───────────────────────────────────────────┘
```

**Qué hacer:**
1. Cierre todos los archivos Excel que tenga abiertos de la carpeta Datos.
2. Intente nuevamente.
3. Si el error persiste, comuníquese con el administrador del sistema e informe el texto del error que aparece en el mensaje.

---

### El consolidado tiene columnas vacías (K a Q sin datos)

**Qué significa:** El número de cédula de esa persona aparece escrito de forma diferente en la planilla y en las facturas (por ejemplo, con puntos o sin puntos).

**Qué hacer:** Informe al administrador del sistema qué persona tiene las columnas vacías para que pueda corregir el problema.

---

### La consola muestra: `Aviso: 'MARZO_estructurado.json' sin facturas — se omite.`

**Qué significa:** El programa encontró el archivo de planilla del mes MARZO pero no encontró el archivo de facturas con el mismo nombre.

**Qué hacer:** Asegúrese de haber ejecutado también el **Paso 2** (conversión de facturas) antes de ejecutar el **Paso 4** (consolidado), y que ambos PDF correspondan al mismo mes.

---

## Resumen del proceso en orden

| Orden | Qué ejecutar | Qué necesita tener listo |
|-------|-------------|--------------------------|
| 1° | Convertidor Planilla UGPP | PDF de la planilla del mes |
| 2° | Convertidor Facturas | PDF de las facturas del mes |
| 3° | Excel a JSON | (automático, no requiere archivos extra) |
| 4° | Consolidador Final | (automático, usa los JSON del paso 3) |

Al terminar el paso 4, el archivo `_CONSOLIDADO.xlsx` está listo para usar.
