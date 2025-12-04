# analisis_financiero_unificado.py
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import math

# ----------------- UTILIDADES -----------------
def safe_float(s):
    try:
        return float(s)
    except:
        return 0.0

def safe_div(a, b):
    try:
        return a / b if b != 0 else float('nan')
    except:
        return float('nan')

def pct(v):
    return v * 100 if v == v else float('nan')

def fmt_num(v):
    return "N/A" if v != v else f"{v:,.2f}"

def fmt_pct(v):
    return "N/A" if v != v else f"{pct(v):.2f}%"

# ----------------- SCROLLABLE FRAME -----------------
class ScrollableFrame(ttk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.scrollable_frame, anchor="nw")
        # mousewheel support
        self.scrollable_frame.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.scrollable_frame.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

# ----------------- RATIOS (configuración) -----------------
CATEGORIES = {
    "Liquidez": ["Razón corriente", "Prueba ácida", "Capital de trabajo"],
    "Actividad": ["Rotación inventarios", "Período promedio de inventarios", "Rotación CxC", "Período promedio de cobro", "Rotación activos"],
    "Endeudamiento": ["Pasivo/Activo", "Deuda/Patrimonio", "Cobertura intereses"],
    "Rentabilidad": ["Margen neto", "ROA", "ROE"]
}

FIELDS_FOR_RATIO = {
    # Liquidez
    "Razón corriente": ["Activo corriente", "Pasivo corriente"],
    "Prueba ácida": ["Activo corriente", "Inventario promedio", "Pasivo corriente"],
    "Capital de trabajo": ["Activo corriente", "Pasivo corriente"],
    # Actividad
    "Rotación inventarios": ["Costo de ventas", "Inventario promedio"],
    "Período promedio de inventarios": ["Costo de ventas", "Inventario promedio"],
    "Rotación CxC": ["Ventas netas", "Cuentas por cobrar promedio"],
    "Período promedio de cobro": ["Ventas netas", "Cuentas por cobrar promedio"],
    "Rotación activos": ["Ventas netas", "Activo total"],
    # Endeudamiento
    "Pasivo/Activo": ["Pasivo total", "Activo total"],
    "Deuda/Patrimonio": ["Pasivo total", "Patrimonio"],
    "Cobertura intereses": ["UAII", "Gastos por intereses"],
    # Rentabilidad
    "Margen neto": ["Utilidad neta", "Ventas netas"],
    "ROA": ["Utilidad neta", "Activo total"],
    "ROE": ["Utilidad neta", "Patrimonio"]
}

HELP_TEXT = {
    "Inventario promedio": "=(Inventario inicial + Inventario final)/2. Si no ingresas inicial/final, escribe aquí el promedio.",
    "Cuentas por cobrar promedio": "Saldo promedio de clientes en el periodo.",
    "UAII": "Utilidad antes de intereses e impuestos (EBIT).",
    "Período promedio de cobro": "360 / Rotación CxC (días).",
    "Período promedio de inventarios": "360 / Rotación de inventarios (días)."
}

# ----------------- CÁLCULOS RATIOS -----------------
def compute_ratio(name, values):
    # values: dict field->float
    if name == "Razón corriente":
        v = safe_div(values.get("Activo corriente",0), values.get("Pasivo corriente",0))
        exp = f"{fmt_num(values.get('Activo corriente',0))} / {fmt_num(values.get('Pasivo corriente',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Prueba ácida":
        v = safe_div(values.get("Activo corriente",0) - values.get("Inventario promedio",0), values.get("Pasivo corriente",0))
        exp = f"({fmt_num(values.get('Activo corriente',0))} - {fmt_num(values.get('Inventario promedio',0))}) / {fmt_num(values.get('Pasivo corriente',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Capital de trabajo":
        v = values.get("Activo corriente",0) - values.get("Pasivo corriente",0)
        exp = f"{fmt_num(values.get('Activo corriente',0))} - {fmt_num(values.get('Pasivo corriente',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Rotación inventarios":
        v = safe_div(values.get("Costo de ventas",0), values.get("Inventario promedio",0))
        exp = f"{fmt_num(values.get('Costo de ventas',0))} / {fmt_num(values.get('Inventario promedio',0))} = {fmt_num(v)} vueltas/año"
        return v, exp
    if name == "Período promedio de inventarios":
        rot = safe_div(values.get("Costo de ventas",0), values.get("Inventario promedio",0))
        v = safe_div(360, rot) if rot and rot != 0 else float('nan')
        exp = f"360 / {fmt_num(rot)} = {fmt_num(v)} días"
        return v, exp
    if name == "Rotación CxC":
        v = safe_div(values.get("Ventas netas",0), values.get("Cuentas por cobrar promedio",0))
        exp = f"{fmt_num(values.get('Ventas netas',0))} / {fmt_num(values.get('Cuentas por cobrar promedio',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Período promedio de cobro":
        rot = safe_div(values.get("Ventas netas",0), values.get("Cuentas por cobrar promedio",0))
        v = safe_div(360, rot) if rot and rot != 0 else float('nan')
        exp = f"360 / {fmt_num(rot)} = {fmt_num(v)} días"
        return v, exp
    if name == "Rotación activos":
        v = safe_div(values.get("Ventas netas",0), values.get("Activo total",0))
        exp = f"{fmt_num(values.get('Ventas netas',0))} / {fmt_num(values.get('Activo total',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Pasivo/Activo":
        v = safe_div(values.get("Pasivo total",0), values.get("Activo total",0))
        exp = f"{fmt_num(values.get('Pasivo total',0))} / {fmt_num(values.get('Activo total',0))} = {fmt_pct(v)}"
        return v, exp
    if name == "Deuda/Patrimonio":
        v = safe_div(values.get("Pasivo total",0), values.get("Patrimonio",0))
        exp = f"{fmt_num(values.get('Pasivo total',0))} / {fmt_num(values.get('Patrimonio',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Cobertura intereses":
        v = safe_div(values.get("UAII",0), values.get("Gastos por intereses",0))
        exp = f"{fmt_num(values.get('UAII',0))} / {fmt_num(values.get('Gastos por intereses',0))} = {fmt_num(v)}"
        return v, exp
    if name == "Margen neto":
        v = safe_div(values.get("Utilidad neta",0), values.get("Ventas netas",0))
        exp = f"{fmt_num(values.get('Utilidad neta',0))} / {fmt_num(values.get('Ventas netas',0))} = {fmt_pct(v)}"
        return v, exp
    if name == "ROA":
        v = safe_div(values.get("Utilidad neta",0), values.get("Activo total",0))
        exp = f"{fmt_num(values.get('Utilidad neta',0))} / {fmt_num(values.get('Activo total',0))} = {fmt_pct(v)}"
        return v, exp
    if name == "ROE":
        v = safe_div(values.get("Utilidad neta",0), values.get("Patrimonio",0))
        exp = f"{fmt_num(values.get('Utilidad neta',0))} / {fmt_num(values.get('Patrimonio',0))} = {fmt_pct(v)}"
        return v, exp
    return float('nan'), "Ratio no implementado"

def interpret_ratio(name, value):
    # Interpretations (heuristic, based on your documents)
    if value != value: return "No disponible"
    if name == "Razón corriente":
        if value < 1: return "Insuficiente: riesgo de iliquidez a corto plazo."
        if value <= 2: return "Aceptable."
        return "Alta liquidez; posible uso ineficiente de recursos."
    if name == "Prueba ácida":
        return "Buena liquidez sin depender del inventario." if value >= 1 else "Riesgo: depende del inventario."
    if name == "Capital de trabajo":
        return "Positivo" if value >= 0 else "Negativo: riesgo operativo."
    if name == "Rotación inventarios":
        if value < 3: return "Rotación baja: exceso de inventario."
        if value <= 6: return "Rotación adecuada."
        return "Rotación alta."
    if name == "Período promedio de inventarios":
        return "Gestión adecuada." if value <= 120 else "Inventarios lentos."
    if name == "Rotación CxC":
        return "Cobranza eficiente." if value >= 4 else "Cobranza lenta."
    if name == "Período promedio de cobro":
        return "Período saludable." if value <= 90 else "Cobro lento; revisar políticas de crédito."
    if name == "Rotación activos":
        return "Activos bien aprovechados." if value >= 1 else "Activos infrautilizados."
    if name == "Pasivo/Activo":
        if value > 0.6: return "Alto endeudamiento."
        if value >= 0.4: return "Endeudamiento moderado."
        return "Bajo endeudamiento."
    if name == "Deuda/Patrimonio":
        return "Riesgo alto si >1." if value > 1 else "Estructura equilibrada."
    if name == "Cobertura intereses":
        return "Adecuada" if value >= 2 else "Cobertura insuficiente."
    if name == "Margen neto":
        if value < 0.05: return "Margen bajo."
        if value <= 0.15: return "Margen adecuado."
        return "Margen alto."
    if name == "ROA":
        return "ROA aceptable." if value >= 0.05 else "ROA bajo."
    if name == "ROE":
        if value < 0.10: return "ROE bajo."
        if value <= 0.20: return "ROE adecuado."
        return "ROE alto."
    return "Sin interpretación definida."

# ----------------- DUPONT / ANALYSIS FUNCTIONS -----------------
def calcular_ratios_y_dupont_from_inputs(current, prior, bg_current, bg_prior, er_current, er_prior):
    # current/prior: dict with keys as in top inputs
    d = current
    p = prior
    ventas = d.get('Ventas netas', 0)
    utilidad_neta = d.get('Utilidad neta', 0)
    activo_total = d.get('Activo total', 0)
    patrimonio = d.get('Patrimonio', 0)
    pasivo_total = d.get('Pasivo total', 0)
    costo_ventas = d.get('Costo de ventas', 0)
    inventario_prom = d.get('Inventario promedio', d.get('Inventarios', 0))
    cxc_prom = d.get('Cuentas por cobrar promedio', d.get('Cuentas por cobrar', 0))

    margen_neto = safe_div(utilidad_neta, ventas)
    rot_activos = safe_div(ventas, activo_total)
    equity_multiplier = safe_div(activo_total, patrimonio)
    roe_dupont = margen_neto * rot_activos * equity_multiplier

    rot_inventarios = safe_div(costo_ventas, inventario_prom)
    rot_cxc = safe_div(ventas, cxc_prom)

    razon_endeudamiento = safe_div(pasivo_total, activo_total)
    deuda_patrimonio = safe_div(pasivo_total, patrimonio)

    roa = safe_div(utilidad_neta, activo_total)
    roe = safe_div(utilidad_neta, patrimonio)

    contrib = {}
    try:
        if margen_neto > 0 and rot_activos > 0 and equity_multiplier > 0:
            log_m = math.log(margen_neto)
            log_r = math.log(rot_activos)
            log_a = math.log(equity_multiplier)
            total = log_m + log_r + log_a
            contrib['margen_pct'] = 100 * log_m / total
            contrib['rot_pct'] = 100 * log_r / total
            contrib['apal_pct'] = 100 * log_a / total
        else:
            contrib['margen_pct'] = contrib['rot_pct'] = contrib['apal_pct'] = float('nan')
    except Exception:
        contrib['margen_pct'] = contrib['rot_pct'] = contrib['apal_pct'] = float('nan')

    results = {
        'margen_neto': margen_neto,
        'rot_activos': rot_activos,
        'equity_multiplier': equity_multiplier,
        'roe_dupont': roe_dupont,
        'rot_inventarios': rot_inventarios,
        'rot_cxc': rot_cxc,
        'razon_endeudamiento': razon_endeudamiento,
        'deuda_patrimonio': deuda_patrimonio,
        'roa': roa,
        'roe': roe,
        'contribuciones': contrib
    }
    return results

def analisis_vertical(stmt, kind):
    av = {}
    if kind == "BG":
        total = stmt.get('Activo total', 0)
        for k,v in stmt.items():
            if k != 'Activo total':
                av[k] = safe_div(v, total)
    elif kind == "ER":
        total = stmt.get('Ventas netas', 0)
        for k,v in stmt.items():
            if k != 'Ventas netas':
                av[k] = safe_div(v, total)
    return av

def analisis_horizontal(cur, prior):
    horiz = {}
    for k in set(cur.keys()).union(prior.keys()):
        curv = cur.get(k, 0)
        prv = prior.get(k, 0)
        horiz[k] = safe_div(curv - prv, prv)
    return horiz

# ----------------- INTERFAZ / APLICACIÓN -----------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Análisis Financiero Unificado")
        root.geometry("1100x800")

        # Notebook (pestañas)
        notebook = ttk.Notebook(root)
        notebook.pack(fill='both', expand=True)

        # Tab 1: Ratios
        tab_ratios_container = ScrollableFrame(notebook)
        tab_ratios = tab_ratios_container.scrollable_frame
        notebook.add(tab_ratios_container, text="Ratios")

        # Tab 2: Análisis (DuPont + vertical + horizontal)
        tab_analysis_container = ScrollableFrame(notebook)
        tab_analysis = tab_analysis_container.scrollable_frame
        notebook.add(tab_analysis_container, text="DuPont + Vertical + Horizontal")

        # ---------- TAB RATIOS ----------
        ttk.Label(tab_ratios, text="Calculadora de Ratios", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=4, pady=8, sticky='w')
        ttk.Label(tab_ratios, text="Categoría:").grid(row=1, column=0, sticky='w', padx=6)
        self.combo_cat = ttk.Combobox(tab_ratios, values=list(CATEGORIES.keys()), state="readonly", width=30)
        self.combo_cat.grid(row=1, column=1, sticky='w', padx=6)
        self.combo_cat.bind("<<ComboboxSelected>>", self.on_category)

        ttk.Label(tab_ratios, text="Ratio:").grid(row=2, column=0, sticky='w', padx=6)
        self.combo_ratio = ttk.Combobox(tab_ratios, values=[], state="readonly", width=45)
        self.combo_ratio.grid(row=2, column=1, sticky='w', padx=6)
        self.combo_ratio.bind("<<ComboboxSelected>>", self.on_ratio)

        ttk.Label(tab_ratios, text="Datos (solo campos necesarios se habilitan):", font=("Segoe UI", 10, "bold")).grid(row=3, column=0, columnspan=4, sticky='w', padx=6, pady=6)

        # Create all possible input fields for ratios (but disabled until required)
        all_fields = set(sum(FIELDS_FOR_RATIO.values(), []))
        self.ratio_entries = {}
        self.ratio_help = {}
        r = 4
        col = 0
        for i, field in enumerate(sorted(all_fields)):
            lbl = ttk.Label(tab_ratios, text=field + ":")
            ent = ttk.Entry(tab_ratios, width=22, state='disabled')
            help_lbl = ttk.Label(tab_ratios, text=HELP_TEXT.get(field,""), foreground='gray')
            lbl.grid(row=r, column=col, sticky='w', padx=6, pady=2)
            ent.grid(row=r, column=col+1, sticky='w', padx=6, pady=2)
            help_lbl.grid(row=r+1, column=col, columnspan=2, sticky='w', padx=6)
            self.ratio_entries[field] = ent
            self.ratio_help[field] = help_lbl
            if col == 0:
                col = 2
            else:
                col = 0
                r += 2
        self.ratios_last_row = r + 2

        # Buttons and report
        ttk.Button(tab_ratios, text="Calcular ratio seleccionado", command=self.calculate_ratio).grid(row=self.ratios_last_row, column=0, padx=8, pady=12, sticky='w')
        ttk.Button(tab_ratios, text="Limpiar campos", command=self.reset_ratio_fields).grid(row=self.ratios_last_row, column=1, padx=8, pady=12, sticky='w')
        ttk.Label(tab_ratios, text="Reporte:").grid(row=self.ratios_last_row+1, column=0, sticky='w', padx=6)
        self.ratio_report = scrolledtext.ScrolledText(tab_ratios, width=120, height=18, wrap='word')
        self.ratio_report.grid(row=self.ratios_last_row+2, column=0, columnspan=4, padx=6, pady=6)
        self.ratio_report.configure(state='disabled')

        # ---------- TAB ANALYSIS ----------
        ttk.Label(tab_analysis, text="Análisis completo: DuPont + Vertical + Horizontal", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=6, pady=8, sticky='w')

        # Inputs: two-year data (current and prior) BG and ER
        ttk.Label(tab_analysis, text="Ingrese los estados financieros: Año actual (A) y Año previo (P)", font=("Segoe UI", 10, "bold")).grid(row=1, column=0, columnspan=6, sticky='w', padx=6)

        # Basic keys for top summary (for DuPont)
        keys = ['Ventas netas', 'Utilidad neta', 'Activo total', 'Patrimonio', 'Pasivo total', 'Costo de ventas']
        self.entries_current = {}
        self.entries_prior = {}

        ttk.Label(tab_analysis, text="Cuenta").grid(row=2, column=0, padx=4, pady=4)
        ttk.Label(tab_analysis, text="Año actual (A)").grid(row=2, column=1, padx=4, pady=4)
        ttk.Label(tab_analysis, text="Año previo (P)").grid(row=2, column=2, padx=4, pady=4)

        r = 3
        for k in keys:
            ttk.Label(tab_analysis, text=k).grid(row=r, column=0, sticky='w', padx=4, pady=2)
            ea = ttk.Entry(tab_analysis, width=18)
            ea.grid(row=r, column=1, padx=4, pady=2)
            ep = ttk.Entry(tab_analysis, width=18)
            ep.grid(row=r, column=2, padx=4, pady=2)
            self.entries_current[k] = ea
            self.entries_prior[k] = ep
            r += 1

        # Inventory initial/final for A and P (to calculate inventory promedio automatically)
        ttk.Label(tab_analysis, text="Inventario inicial A:").grid(row=r, column=0, sticky='w', padx=4, pady=2)
        self.inv_init_A = ttk.Entry(tab_analysis, width=18); self.inv_init_A.grid(row=r, column=1, padx=4, pady=2)
        ttk.Label(tab_analysis, text="Inventario final A:").grid(row=r, column=2, sticky='w', padx=4, pady=2)
        self.inv_final_A = ttk.Entry(tab_analysis, width=18); self.inv_final_A.grid(row=r, column=3, padx=4, pady=2)
        r += 1
        ttk.Label(tab_analysis, text="Inventario inicial P:").grid(row=r, column=0, sticky='w', padx=4, pady=2)
        self.inv_init_P = ttk.Entry(tab_analysis, width=18); self.inv_init_P.grid(row=r, column=1, padx=4, pady=2)
        ttk.Label(tab_analysis, text="Inventario final P:").grid(row=r, column=2, sticky='w', padx=4, pady=2)
        self.inv_final_P = ttk.Entry(tab_analysis, width=18); self.inv_final_P.grid(row=r, column=3, padx=4, pady=2)
        r += 1

        # Balance General detailed (A/P)
        ttk.Label(tab_analysis, text="Balance General (A / P)", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=6, sticky='w', padx=6, pady=6)
        r += 1
        bg_accounts = ['Activo total','Caja y bancos','Cuentas por cobrar','Inventarios','Activos fijos netos','Pasivo corriente','Pasivo no corriente','Pasivo total','Patrimonio']
        self.bg_cur = {}
        self.bg_pr = {}
        ttk.Label(tab_analysis, text="Cuenta").grid(row=r, column=0, sticky='w')
        ttk.Label(tab_analysis, text="Actual (A)").grid(row=r, column=1)
        ttk.Label(tab_analysis, text="Previo (P)").grid(row=r, column=2)
        r+=1
        for acc in bg_accounts:
            ttk.Label(tab_analysis, text=acc+":").grid(row=r, column=0, sticky='w', padx=4, pady=2)
            ecur = ttk.Entry(tab_analysis, width=18); ecur.grid(row=r, column=1, padx=4, pady=2)
            epri = ttk.Entry(tab_analysis, width=18); epri.grid(row=r, column=2, padx=4, pady=2)
            self.bg_cur[acc] = ecur
            self.bg_pr[acc] = epri
            r += 1

        # Estado de Resultados (A/P)
        ttk.Label(tab_analysis, text="Estado de Resultados (A / P)", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=6, sticky='w', padx=6, pady=6)
        r += 1
        er_accounts = ['Ventas netas','Costo de ventas','Utilidad bruta','Gastos operativos','Utilidad operativa','Gastos financieros','Utilidad antes de impuestos','Utilidad neta']
        self.er_cur = {}
        self.er_pr = {}
        ttk.Label(tab_analysis, text="Cuenta").grid(row=r, column=0, sticky='w')
        ttk.Label(tab_analysis, text="Actual (A)").grid(row=r, column=1)
        ttk.Label(tab_analysis, text="Previo (P)").grid(row=r, column=2)
        r+=1
        for acc in er_accounts:
            ttk.Label(tab_analysis, text=acc+":").grid(row=r, column=0, sticky='w', padx=4, pady=2)
            ecur = ttk.Entry(tab_analysis, width=18); ecur.grid(row=r, column=1, padx=4, pady=2)
            epri = ttk.Entry(tab_analysis, width=18); epri.grid(row=r, column=2, padx=4, pady=2)
            self.er_cur[acc] = ecur
            self.er_pr[acc] = epri
            r += 1

        # Buttons for analysis
        ttk.Button(tab_analysis, text="Generar Análisis Completo", command=self.generate_full_analysis).grid(row=r, column=0, pady=12, padx=6, sticky='w')
        ttk.Button(tab_analysis, text="Limpiar campos", command=self.reset_analysis_fields).grid(row=r, column=1, pady=12, padx=6, sticky='w')
        r += 1

        ttk.Label(tab_analysis, text="Reporte de Análisis:").grid(row=r, column=0, columnspan=4, sticky='w', padx=6)
        r += 1
        self.analysis_report = scrolledtext.ScrolledText(tab_analysis, width=120, height=22, wrap='word')
        self.analysis_report.grid(row=r, column=0, columnspan=4, padx=6, pady=6)
        self.analysis_report.configure(state='disabled')

    # ---------- RATIOS TAB HANDLERS ----------
    def on_category(self, event=None):
        cat = self.combo_cat.get()
        self.combo_ratio['values'] = CATEGORIES.get(cat, [])
        self.combo_ratio.set('')
        self.reset_ratio_fields(keep_locked=True)
        self.clear_ratio_report()

    def on_ratio(self, event=None):
        r = self.combo_ratio.get()
        # disable all entries, then enable needed ones
        for field, ent in self.ratio_entries.items():
            ent.configure(state='disabled')
            ent.delete(0, tk.END)
        needed = FIELDS_FOR_RATIO.get(r, [])
        for f in needed:
            ent = self.ratio_entries.get(f)
            if ent:
                ent.configure(state='normal')
        self.clear_ratio_report()

    def calculate_ratio(self):
        r = self.combo_ratio.get()
        if not r:
            messagebox.showwarning("Atención", "Selecciona categoría y ratio.")
            return
        needed = FIELDS_FOR_RATIO.get(r, [])
        values = {}
        # If Inventario promedio can be computed from inventory initial/final in Analysis tab, use that value:
        # But here we only use local inputs: check ratio_entries
        for f in needed:
            values[f] = safe_float(self.ratio_entries[f].get())
        val, exp = compute_ratio(r, values)
        interp = interpret_ratio(r, val)
        # Build report
        lines = [f"RATIO: {r}", "-"*40]
        if r in ["Pasivo/Activo","Margen neto","ROA","ROE"]:
            lines.append(f"Resultado: {fmt_pct(val)}")
        elif "Período" in r:
            lines.append(f"Resultado: {fmt_num(val)} días")
        else:
            lines.append(f"Resultado: {fmt_num(val)}")
        lines.append(f"Cálculo: {exp}")
        lines.append("")
        lines.append("Interpretación:")
        lines.append(interp)
        lines.append("")
        lines.append("Campos ingresados:")
        for f in needed:
            lines.append(f" - {f}: {fmt_num(values.get(f,0))}")
        lines.append("")
        lines.append("Conceptos rápidos:")
        lines.append(" - Inventario promedio = (Inventario inicial + Inventario final) / 2")
        lines.append(" - Período promedio de cobro = 360 / Rotación CxC (días)")
        # show report
        self.ratio_report.configure(state='normal')
        self.ratio_report.delete('1.0', tk.END)
        self.ratio_report.insert(tk.END, "\n".join(lines))
        self.ratio_report.configure(state='disabled')

    def reset_ratio_fields(self, keep_locked=False):
        for ent in self.ratio_entries.values():
            ent.configure(state='normal')
            ent.delete(0, tk.END)
            if not keep_locked:
                ent.configure(state='disabled')
        if not keep_locked:
            self.combo_cat.set('')
            self.combo_ratio.set('')
        self.clear_ratio_report()

    def clear_ratio_report(self):
        self.ratio_report.configure(state='normal')
        self.ratio_report.delete('1.0', tk.END)
        self.ratio_report.configure(state='disabled')

    # ---------- ANALYSIS TAB HANDLERS ----------
    def generate_full_analysis(self):
        try:
            # Read top summary fields current/prior
            current = {}
            prior = {}
            for k, ent in self.entries_current.items():
                current[k] = safe_float(ent.get())
            for k, ent in self.entries_prior.items():
                prior[k] = safe_float(ent.get())

            # Calculate inventory promedio automatically if initial/final provided; else rely on 'Inventario promedio' if present in BG
            invA_init = safe_float(self.inv_init_A.get())
            invA_final = safe_float(self.inv_final_A.get())
            invP_init = safe_float(self.inv_init_P.get())
            invP_final = safe_float(self.inv_final_P.get())

            if invA_init != 0 or invA_final != 0:
                invA_prom = (invA_init + invA_final) / 2.0
            else:
                # try to get from BG current 'Inventarios' field
                invA_prom = safe_float(self.bg_cur.get('Inventarios').get())

            if invP_init != 0 or invP_final != 0:
                invP_prom = (invP_init + invP_final) / 2.0
            else:
                invP_prom = safe_float(self.bg_pr.get('Inventarios').get())

            # If user didn't enter Inventario promedio in top keys, set it
            current['Inventario promedio'] = invA_prom
            prior['Inventario promedio'] = invP_prom

            # Read BG and ER detailed
            bg_cur = {}
            bg_pr = {}
            for k, ent in self.bg_cur.items():
                bg_cur[k] = safe_float(ent.get())
            for k, ent in self.bg_pr.items():
                bg_pr[k] = safe_float(ent.get())

            er_cur = {}
            er_pr = {}
            for k, ent in self.er_cur.items():
                er_cur[k] = safe_float(ent.get())
            for k, ent in self.er_pr.items():
                er_pr[k] = safe_float(ent.get())

            # Ensure current top keys have necessary values (fallback to BG/ER if missing)
            if current.get('Ventas netas', 0) == 0:
                current['Ventas netas'] = er_cur.get('Ventas netas', 0)
            if current.get('Utilidad neta', 0) == 0:
                current['Utilidad neta'] = er_cur.get('Utilidad neta', 0)
            if current.get('Costo de ventas', 0) == 0:
                current['Costo de ventas'] = er_cur.get('Costo de ventas', 0)
            if current.get('Activo total', 0) == 0:
                current['Activo total'] = bg_cur.get('Activo total', 0)
            if current.get('Patrimonio', 0) == 0:
                current['Patrimonio'] = bg_cur.get('Patrimonio', 0)
            if current.get('Pasivo total', 0) == 0:
                current['Pasivo total'] = bg_cur.get('Pasivo total', 0)

            # same for prior
            if prior.get('Ventas netas', 0) == 0:
                prior['Ventas netas'] = er_pr.get('Ventas netas', 0)
            if prior.get('Utilidad neta', 0) == 0:
                prior['Utilidad neta'] = er_pr.get('Utilidad neta', 0)
            if prior.get('Costo de ventas', 0) == 0:
                prior['Costo de ventas'] = er_pr.get('Costo de ventas', 0)
            if prior.get('Activo total', 0) == 0:
                prior['Activo total'] = bg_pr.get('Activo total', 0)
            if prior.get('Patrimonio', 0) == 0:
                prior['Patrimonio'] = bg_pr.get('Patrimonio', 0)
            if prior.get('Pasivo total', 0) == 0:
                prior['Pasivo total'] = bg_pr.get('Pasivo total', 0)

            # compute dupont and ratios
            dupont_res = calcular_ratios_y_dupont_from_inputs(current, prior, bg_cur, bg_pr, er_cur, er_pr)

            # vertical and horizontal analyses
            vert_bg_cur = analisis_vertical(bg_cur, "BG")
            vert_bg_pr = analisis_vertical(bg_pr, "BG")
            vert_er_cur = analisis_vertical(er_cur, "ER")
            vert_er_pr = analisis_vertical(er_pr, "ER")
            horiz_bg = analisis_horizontal(bg_cur, bg_pr)
            horiz_er = analisis_horizontal(er_cur, er_pr)

            # build report text
            rpt_lines = []
            rpt_lines.append("----- RESUMEN EJECUTIVO -----")
            rpt_lines.append(f"Ventas (A): {fmt_num(current.get('Ventas netas',0))}")
            rpt_lines.append(f"Utilidad neta (A): {fmt_num(current.get('Utilidad neta',0))}")
            rpt_lines.append(f"Activo total (A): {fmt_num(current.get('Activo total',0))}")
            rpt_lines.append(f"Patrimonio (A): {fmt_num(current.get('Patrimonio',0))}\n")

            rpt_lines.append("----- DU PONT (ROE) -----")
            rpt_lines.append(f"Margen neto = {fmt_pct(dupont_res['margen_neto'])}")
            rpt_lines.append(f"Rotación de activos = {fmt_num(dupont_res['rot_activos'])}")
            rpt_lines.append(f"Equity multiplier = {fmt_num(dupont_res['equity_multiplier'])}")
            rpt_lines.append(f"ROE (DuPont) = {fmt_pct(dupont_res['roe_dupont'])}")
            rpt_lines.append(f"ROA = {fmt_pct(dupont_res['roa'])}")
            rpt_lines.append(f"ROE directo = {fmt_pct(dupont_res['roe'])}\n")

            rpt_lines.append("Contribución aproximada al ROE (log-based):")
            rpt_lines.append(f" - Margen: {fmt_num(dupont_res['contribuciones'].get('margen_pct'))}%")
            rpt_lines.append(f" - Rotación: {fmt_num(dupont_res['contribuciones'].get('rot_pct'))}%")
            rpt_lines.append(f" - Apalancamiento: {fmt_num(dupont_res['contribuciones'].get('apal_pct'))}%\n")

            rpt_lines.append("----- ANÁLISIS VERTICAL (BG - % sobre Activo A) -----")
            for k,v in vert_bg_cur.items():
                rpt_lines.append(f" {k}: {fmt_pct(v)}")
            rpt_lines.append("\n----- ANÁLISIS VERTICAL (ER - % sobre Ventas A) -----")
            for k,v in vert_er_cur.items():
                rpt_lines.append(f" {k}: {fmt_pct(v)}")

            rpt_lines.append("\n----- ANÁLISIS HORIZONTAL (Variación interanual) -----")
            rpt_lines.append("Balance General (A vs P):")
            for k,v in horiz_bg.items():
                rpt_lines.append(f" {k}: {fmt_pct(v)}")
            rpt_lines.append("Estado de Resultados (A vs P):")
            for k,v in horiz_er.items():
                rpt_lines.append(f" {k}: {fmt_pct(v)}")

            # strengths / weaknesses heuristics
            rpt_lines.append("\n----- FORTALEZAS / DEBILIDADES (automático) -----")
            if dupont_res['margen_neto'] >= 0.05:
                rpt_lines.append(" - Margen neto adecuado o bueno.")
            else:
                rpt_lines.append(" - Margen neto bajo: revisar costos y precios.")
            if dupont_res['rot_activos'] >= 1:
                rpt_lines.append(" - Rotación de activos aceptable.")
            else:
                rpt_lines.append(" - Rotación de activos baja: activos infrautilizados.")
            if dupont_res['equity_multiplier'] <= 2:
                rpt_lines.append(" - Apalancamiento moderado.")
            else:
                rpt_lines.append(" - Apalancamiento alto: depender de deuda.")

            # trends
            vend_growth = horiz_er.get('Ventas netas', float('nan'))
            util_growth = horiz_er.get('Utilidad neta', float('nan'))
            rpt_lines.append("\n----- TENDENCIAS -----")
            rpt_lines.append(f" - Ventas (A vs P): {fmt_pct(vend_growth)}")
            rpt_lines.append(f" - Utilidad neta (A vs P): {fmt_pct(util_growth)}")
            if not math.isnan(vend_growth) and not math.isnan(util_growth):
                if util_growth > vend_growth:
                    rpt_lines.append(" - Rentabilidad creciendo más que ventas (mejora de margen).")
                elif util_growth < vend_growth:
                    rpt_lines.append(" - Ventas crecen pero utilidades no: revisar costos.")
                else:
                    rpt_lines.append(" - Margen estable frente a crecimiento de ventas.")

            # recommendation heuristic (simple)
            rpt_lines.append("\n----- RECOMENDACIÓN (heurística) -----")
            debt_level = dupont_res['razon_endeudamiento']
            if dupont_res['roe'] > 0.15 and debt_level < 0.6 and dupont_res['margen_neto'] > 0:
                rpt_lines.append(" Inversión favorable (ROE alto, deuda moderada, margen positivo).")
            else:
                rpt_lines.append(" Revisar con cautela: mejorar margen, reducir deuda o aumentar eficiencia.")

            # Show report
            self.analysis_report.configure(state='normal')
            self.analysis_report.delete('1.0', tk.END)
            self.analysis_report.insert(tk.END, "\n".join(rpt_lines))
            self.analysis_report.configure(state='disabled')

        except Exception as e:
            messagebox.showerror("Error", f"Verifica los campos. Error: {e}")

    def reset_analysis_fields(self):
        for ent in self.entries_current.values():
            ent.delete(0, tk.END)
        for ent in self.entries_prior.values():
            ent.delete(0, tk.END)
        for ent in [self.inv_init_A, self.inv_final_A, self.inv_init_P, self.inv_final_P]:
            ent.delete(0, tk.END)
        for ent in self.bg_cur.values():
            ent.delete(0, tk.END)
        for ent in self.bg_pr.values():
            ent.delete(0, tk.END)
        for ent in self.er_cur.values():
            ent.delete(0, tk.END)
        for ent in self.er_pr.values():
            ent.delete(0, tk.END)
        self.analysis_report.configure(state='normal')
        self.analysis_report.delete('1.0', tk.END)
        self.analysis_report.configure(state='disabled')

# ----------------- RUN -----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
