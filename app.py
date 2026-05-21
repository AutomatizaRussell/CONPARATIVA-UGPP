import sys
import threading
from pathlib import Path
from unittest.mock import patch, MagicMock
import tkinter as tk
from tkinter import filedialog, messagebox

BASE = Path(__file__).parent
sys.path.insert(0, str(BASE / "convertidor"))

PASOS = [
    "1. Procesar PDF Facturas",
    "2. Procesar PDF Estructura",
    "3. Convertir Excel a JSON",
    "4. Generar Excel Consolidado",
]


def _mock_root():
    m = MagicMock()
    m.withdraw = MagicMock()
    return m


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Comparativa UGPP")
        self.geometry("530x340")
        self.resizable(False, False)
        self.configure(bg="#F0F2F5")

        self.pdf_fac  = tk.StringVar()
        self.pdf_conv = tk.StringVar()
        self._lbl_pasos = []

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────

    def _build_ui(self):
        tk.Label(self, text="Comparativa UGPP", font=("Arial", 15, "bold"),
                 bg="#F0F2F5", fg="#2A4D69").pack(pady=(18, 10))

        self._fila_pdf(self.pdf_fac,  "PDF Facturas:",   "Selecciona el PDF de Facturas")
        self._fila_pdf(self.pdf_conv, "PDF Estructura:", "Selecciona el PDF de Estructura")

        self.btn = tk.Button(
            self, text="PROCESAR", font=("Arial", 11, "bold"),
            bg="#2A4D69", fg="white", relief="flat",
            padx=24, pady=7, cursor="hand2",
            command=self._iniciar,
        )
        self.btn.pack(pady=12)

        frm = tk.LabelFrame(self, text=" Progreso ", bg="#F0F2F5",
                            font=("Arial", 9), padx=10, pady=6)
        frm.pack(fill="x", padx=18, pady=(0, 14))

        for nombre in PASOS:
            var = tk.StringVar(value=f"  {nombre}")
            lbl = tk.Label(frm, textvariable=var, anchor="w",
                           font=("Arial", 10), bg="#F0F2F5", fg="gray")
            lbl.pack(fill="x", pady=1)
            self._lbl_pasos.append((var, lbl))

    def _fila_pdf(self, var, etiqueta, titulo):
        frm = tk.Frame(self, bg="#F0F2F5")
        frm.pack(fill="x", padx=18, pady=3)
        tk.Label(frm, text=etiqueta, width=15, anchor="w",
                 bg="#F0F2F5", font=("Arial", 10)).pack(side="left")
        tk.Entry(frm, textvariable=var, width=36,
                 font=("Arial", 9)).pack(side="left", padx=(0, 6))
        tk.Button(frm, text="...", width=3, cursor="hand2",
                  command=lambda v=var, t=titulo: self._elegir(v, t)).pack(side="left")

    def _elegir(self, var, titulo):
        path = filedialog.askopenfilename(
            title=titulo, filetypes=[("Archivos PDF", "*.pdf")]
        )
        if path:
            var.set(path)

    # ── estado de pasos ───────────────────────────────────────────────────

    def _set_paso(self, idx, estado):
        iconos  = {"espera": "  ", "proceso": ">>", "ok": "OK", "error": "!!"}
        colores = {"espera": "gray", "proceso": "#2A4D69", "ok": "green", "error": "red"}
        var, lbl = self._lbl_pasos[idx]
        var.set(f"[{iconos[estado]}] {PASOS[idx]}")
        lbl.configure(fg=colores[estado])

    # ── flujo principal ───────────────────────────────────────────────────

    def _iniciar(self):
        if not self.pdf_fac.get() or not self.pdf_conv.get():
            messagebox.showwarning("Atencion", "Selecciona ambos PDFs antes de continuar.")
            return
        self.btn.config(state="disabled")
        for i in range(len(PASOS)):
            self._set_paso(i, "espera")
        threading.Thread(target=self._flujo, daemon=True).start()

    def _flujo(self):
        pdf_fac  = self.pdf_fac.get()
        pdf_conv = self.pdf_conv.get()

        pasos_fn = [
            lambda: self._paso_factura(pdf_fac),
            lambda: self._paso_convertidor(pdf_conv),
            self._paso_json,
            self._paso_excel,
        ]

        for i, fn in enumerate(pasos_fn):
            self.after(0, self._set_paso, i, "proceso")
            try:
                fn()
                self.after(0, self._set_paso, i, "ok")
            except Exception as e:
                self.after(0, self._set_paso, i, "error")
                self.after(0, messagebox.showerror, "Error",
                           f"Fallo en paso {i + 1}:\n{e}")
                self.after(0, self.btn.config, {"state": "normal"})
                return

        self.after(0, messagebox.showinfo, "Listo",
                   "Proceso completado.\nEl Excel consolidado esta en la carpeta Datos.")
        self.after(0, self.btn.config, {"state": "normal"})

    # ── pasos individuales ────────────────────────────────────────────────

    def _paso_factura(self, pdf_path):
        import factura  # type: ignore
        with patch.object(factura, "Tk", return_value=_mock_root()), \
             patch("tkinter.filedialog.askopenfilename", return_value=pdf_path), \
             patch("tkinter.messagebox.showinfo"), \
             patch("tkinter.messagebox.showwarning"), \
             patch("tkinter.messagebox.showerror"), \
             patch("os.startfile"):
            factura.seleccionar_y_convertir_facturas()

    def _paso_convertidor(self, pdf_path):
        import convertidor
        with patch.object(convertidor, "Tk", return_value=_mock_root()), \
             patch("tkinter.filedialog.askopenfilename", return_value=pdf_path), \
             patch("tkinter.messagebox.showinfo"), \
             patch("tkinter.messagebox.showwarning"), \
             patch("tkinter.messagebox.showerror"):
            convertidor.seleccionar_y_convertir()

    def _paso_json(self):
        import excel_a_json  # type: ignore
        excel_a_json.convertir_todos()

    def _paso_excel(self):
        import poblar_excel  # type: ignore
        poblar_excel.main()


if __name__ == "__main__":
    App().mainloop()
