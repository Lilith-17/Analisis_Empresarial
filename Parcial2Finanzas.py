# analisis_financiero.py
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import math

# ------------------ UTILIDADES ------------------
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

# ------------------ SCROLLABLE CONTAINER ------------------
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
        self.scrollable_frame.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.scrollable_frame.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

# ------------------ RATIOS CONFIG (user code integrated) ------------------
CATEGORIES = {
    "Liquidez": ["Razón corriente", "Prueba ácida", "Capital de trabajo"],
    "Actividad": ["Rotación inventarios", "Período promedio de inventarios", "Rotación CxC", "Período promedio de cobro", "Rotación activos"],
    "Endeudamiento": ["Pasivo/Activo", "Deuda/Patrimonio", "Cobertura intereses"],
    "Rentabilidad": ["Margen neto", "ROA", "ROE"]
}

FIELDS = {
    "Razón corriente":["Activo corriente","Pasivo corriente"],
    "Prueba ácida":["Activo corriente","Inventarios","Pasivo corriente"],
    "Capital de trabajo":["Activo corriente","Pasivo corriente"],
    "Rotación inventarios":["Costo de ventas","Inventario promedio"],
    "Período promedio de inventarios":["Costo de ventas","Inventario promedio"],
    "Rotación CxC":["Ventas netas","Cuentas por cobrar"],
    "Período promedio de cobro":["Ventas netas","Cuentas por cobrar"],
    "Rotación activos":["Ventas netas","Activo total"],
    "Pasivo/Activo":["Pasivo total","Activo total"],
    "Deuda/Patrimonio":["Pasivo total","Patrimonio"],
    "Cobertura intereses":["UAII","Gastos por intereses"],
    "Margen neto":["Utilidad neta","Ventas netas"],
    "ROA":["Utilidad neta","Activo total"],
    "ROE":["Utilidad neta","Patrimonio"]
}

def interpretar(ratio, v):
    # versión mejorada y algo más elaborada (breve)
    if v != v: return "No disponible (posible división por cero o dato faltante)."
    if ratio=="Razón corriente":
        if v<1:
            return "Insuficiente: riesgo de iliquidez a corto plazo. Recomendación: aumentar activos líquidos o renegociar pasivos corrientes."
        if v<=2:
            return "Aceptable: cubre obligaciones a corto plazo. Mantener control de ciclo operativo."
        return "Exceso de liquidez: la empresa puede no estar usando eficientemente recursos (considerar inversión o reducción de inventarios)."
    if ratio=="Prueba ácida":
        if v<1:
            return "Riesgo: dependencia del inventario para cubrir pasivos. Revisar rotación de inventarios y políticas de crédito."
        return "Buena liquidez inmediata sin depender del inventario."
    if ratio=="Capital de trabajo":
        if v < 0:
            return "Capital de trabajo negativo: riesgo operativo y posible iliquidez. Priorizar gestión de corto plazo."
        return "Capital de trabajo positivo: capacidad para operar con normalidad."
    if ratio=="Rotación inventarios":
        if v < 3:
            return "Rotación baja: exceso de inventario o problemas de ventas. Revisar niveles y promociones."
        if v <= 6:
            return "Rotación adecuada: la venta y reposición están equilibradas."
        return "Rotación alta: buen ritmo de ventas, pero cuidado con rupturas de stock."
    if ratio=="Período promedio de inventarios":
        if v > 120:
            return "Inventarios permanecen mucho tiempo: riesgo de obsolescencia."
        return "Inventario se renueva en tiempo razonable."
    if ratio=="Rotación CxC":
        if v < 4:
            return "Cobranza lenta: revisar condiciones de crédito y gestión de cobranzas."
        return "Cobranza eficiente."
    if ratio=="Período promedio de cobro":
        if v > 90:
            return "Días de cobro altos: impacto negativo en liquidez; mejorar políticas o incentivos de pago."
        return "Período de cobro saludable."
    if ratio=="Rotación activos":
        if v < 1:
            return "Baja eficiencia en uso de activos: posibles inversiones improductivas o activos ociosos."
        return "Activos generan ventas eficientemente."
    if ratio=="Pasivo/Activo":
        if v > 0.6:
            return "Alto endeudamiento: riesgo financiero elevado. Revisar estructura de capital."
        if v >= 0.4:
            return "Endeudamiento moderado: equilibrio entre deuda y capital."
        return "Bajo endeudamiento: conservador; puede aumentar financiamiento para crecer."
    if ratio=="Deuda/Patrimonio":
        if v > 1:
            return "Deuda mayor que patrimonio: alta dependencia de financiamiento externo."
        return "Estructura patrimonial equilibrada."
    if ratio=="Cobertura intereses":
        if v < 2:
            return "Cobertura insuficiente: riesgo ante aumentos en tasa de interés."
        return "Capacidad adecuada para pagar intereses."
    if ratio=="Margen neto":
        if v < 0.05:
            return "Margen bajo: revisar precios y estructura de costos."
        if v <= 0.15:
            return "Margen adecuado: negocio rentable."
        return "Margen alto: excelente control de costos o posicionamiento."
    if ratio=="ROA":
        if v < 0.05:
            return "ROA bajo: los activos no generan suficiente retorno."
        return "ROA aceptable: buena utilización de activos."
    if ratio=="ROE":
        if v < 0.10:
            return "ROE bajo: bajo retorno sobre el capital propio."
        if v <= 0.20:
            return "ROE adecuado: retorno aceptable para accionistas."
        return "ROE alto: excelente retorno para accionistas."
    return "Sin interpretación."

def calcular(r, vals):
    # user code mapping
    if r=="Razón corriente": return safe_div(vals.get("Activo corriente",0), vals.get("Pasivo corriente",0))
    if r=="Prueba ácida": return safe_div(vals.get("Activo corriente",0) - vals.get("Inventarios",0), vals.get("Pasivo corriente",0))
    if r=="Capital de trabajo": return vals.get("Activo corriente",0) - vals.get("Pasivo corriente",0)
    if r=="Rotación inventarios": return safe_div(vals.get("Costo de ventas",0), vals.get("Inventario promedio",0))
    if r=="Período promedio de inventarios": return safe_div(360, safe_div(vals.get("Costo de ventas",0), vals.get("Inventario promedio",0)))
    if r=="Rotación CxC": return safe_div(vals.get("Ventas netas",0), vals.get("Cuentas por cobrar",0))
    if r=="Período promedio de cobro": return safe_div(360, safe_div(vals.get("Ventas netas",0), vals.get("Cuentas por cobrar",0)))
    if r=="Rotación activos": return safe_div(vals.get("Ventas netas",0), vals.get("Activo total",0))
    if r=="Pasivo/Activo": return safe_div(vals.get("Pasivo total",0), vals.get("Activo total",0))
    if r=="Deuda/Patrimonio": return safe_div(vals.get("Pasivo total",0), vals.get("Patrimonio",0))
    if r=="Cobertura intereses": return safe_div(vals.get("UAII",0), vals.get("Gastos por intereses",0))
    if r=="Margen neto": return safe_div(vals.get("Utilidad neta",0), vals.get("Ventas netas",0))
    if r=="ROA": return safe_div(vals.get("Utilidad neta",0), vals.get("Activo total",0))
    if r=="ROE": return safe_div(vals.get("Utilidad neta",0), vals.get("Patrimonio",0))
    return float('nan')

# ------------------ DUPONT & ANALYSIS ------------------
def calculate_dupont(current):
    ventas = current.get('Ventas netas', 0)
    utilidad_neta = current.get('Utilidad neta', 0)
    activo_total = current.get('Activo total', 0)
    patrimonio = current.get('Patrimonio', 0)
    costo_ventas = current.get('Costo de ventas', 0)
    inventario_prom = current.get('Inventario promedio', 0)
    cxc_prom = current.get('Cuentas por cobrar promedio', 0)

    margen_neto = safe_div(utilidad_neta, ventas)
    rot_activos = safe_div(ventas, activo_total)
    equity_multiplier = safe_div(activo_total, patrimonio)
    roe_dupont = margen_neto * rot_activos * equity_multiplier
    rot_inventarios = safe_div(costo_ventas, inventario_prom)
    rot_cxc = safe_div(ventas, cxc_prom)
    razon_endeudamiento = safe_div(current.get('Pasivo total',0), activo_total)
    deuda_patrimonio = safe_div(current.get('Pasivo total',0), patrimonio)
    roa = safe_div(utilidad_neta, activo_total)
    roe = safe_div(utilidad_neta, patrimonio)

    # contributions via log (only if positive)
    contrib = {'margen_pct': float('nan'), 'rot_pct': float('nan'), 'apal_pct': float('nan')}
    try:
        if margen_neto > 0 and rot_activos > 0 and equity_multiplier > 0:
            log_m = math.log(margen_neto)
            log_r = math.log(rot_activos)
            log_a = math.log(equity_multiplier)
            tot = log_m + log_r + log_a
            contrib['margen_pct'] = 100 * log_m / tot
            contrib['rot_pct'] = 100 * log_r / tot
            contrib['apal_pct'] = 100 * log_a / tot
    except Exception:
        pass

    return {
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

def analisis_vertical(stmt, kind):
    av = {}
    if kind == "BG":
        total = stmt.get('Activo total', 0)
        for k, v in stmt.items():
            if k == 'Activo total': continue
            av[k] = safe_div(v, total)
    elif kind == "ER":
        total = stmt.get('Ventas netas', 0)
        for k, v in stmt.items():
            if k == 'Ventas netas': continue
            av[k] = safe_div(v, total)
    return av

def analisis_horizontal(cur, prior):
    horiz = {}
    for k in set(cur.keys()).union(prior.keys()):
        curv = cur.get(k, 0)
        prv = prior.get(k, 0)
        horiz[k] = safe_div(curv - prv, prv)
    return horiz

# ------------------ APLICACIÓN ------------------
class UnifiedApp:
    def __init__(self, root):
        self.root = root
        root.title("Análisis Financiero - 3 Notebooks (Compartido)")
        root.geometry("1080x820")

        # Shared data dictionaries (will be filled from Analysis tab)
        self.current = {}   # año actual (A)
        self.prior = {}     # año previo (P)
        self.bg_current = {}
        self.bg_prior = {}
        self.er_current = {}
        self.er_prior = {}

        # NOTEBOOK
        notebook = ttk.Notebook(root)
        notebook.pack(fill='both', expand=True)

        # Tabs (scrollable)
        self.tab_interes = ScrollableFrame(notebook); notebook.add(self.tab_interes, text="Cálculo de Interés")
        self.tab_ratios = ScrollableFrame(notebook); notebook.add(self.tab_ratios, text="Ratios")
        self.tab_dupont = ScrollableFrame(notebook); notebook.add(self.tab_dupont, text="DuPont")
        self.tab_vh = ScrollableFrame(notebook); notebook.add(self.tab_vh, text="Vertical + Horizontal")

        # Build each tab
        self.build_tab_interes(self.tab_interes.scrollable_frame)
        self.build_tab_ratios(self.tab_ratios.scrollable_frame)
        self.build_tab_dupont(self.tab_dupont.scrollable_frame)
        self.build_tab_vh(self.tab_vh.scrollable_frame)

    # ------------------ TAB INTERES (versión PRO: robusta) ------------------
    def build_tab_interes(self, frm):
        ttk.Label(frm, text="Calculadora de Interés Simple y Compuesto",
                font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=4, pady=10)

        # 1. COMBOBOX PRINCIPAL
        ttk.Label(frm, text="Tipo de interés:").grid(row=1, column=0, sticky="w")
        self.int_tipo = ttk.Combobox(frm, values=[
            "Interés Simple",
            "Interés Simple con Variación de Tasas",
            "Interés Compuesto"
        ], state="readonly", width=40)
        self.int_tipo.grid(row=1, column=1, pady=3)
        self.int_tipo.bind("<<ComboboxSelected>>", self.actualizar_tipos_interes)

        # 2. COMBOBOX SECUNDARIO
        ttk.Label(frm, text="Tipo de cálculo:").grid(row=2, column=0, sticky="w")
        self.int_calculo = ttk.Combobox(frm, values=[], state="readonly", width=40)
        self.int_calculo.grid(row=2, column=1, pady=3)
        self.int_calculo.bind("<<ComboboxSelected>>", self.mostrar_campos_interes)

        # 3. CONTENEDOR DE CAMPOS
        self.int_fields_frame = ttk.LabelFrame(frm, text="Datos requeridos")
        self.int_fields_frame.grid(row=3, column=0, columnspan=4, pady=10, sticky="w")

        self.int_inputs = {}   # entradas dinámicas

        # 4. BOTONES
        ttk.Button(frm, text="Calcular", command=self.calcular_interes).grid(row=4, column=0, pady=10)
        ttk.Button(frm, text="Limpiar", command=self.limpiar_interes).grid(row=4, column=1, pady=10)

        # 5. REPORTE
        ttk.Label(frm, text="Resultado:", font=("Segoe UI", 10, "bold")).grid(row=5, column=0, sticky="w")
        self.int_reporte = scrolledtext.ScrolledText(frm, width=110, height=22, state="disabled")
        self.int_reporte.grid(row=6, column=0, columnspan=4, pady=8)

    def actualizar_tipos_interes(self, event=None):
        tipo = self.int_tipo.get()
        if tipo == "Interés Simple":
            opciones = [
                "Valor Futuro (VF)",
                "Valor Presente (VP)",
                "Interés generado (I)",
                "Tasa de interés (i)",
                "Tiempo (t)"
            ]
        elif tipo == "Interés Simple con Variación de Tasas":
            opciones = ["Monto con múltiples tasas (por tramos)"]
        elif tipo == "Interés Compuesto":
            opciones = [
                "Valor Futuro (VF)",
                "Valor Presente (VP)"
            ]
        else:
            opciones = []
        self.int_calculo.config(values=opciones)
        self.int_calculo.set("")
        self.limpiar_campos_interes()

    def limpiar_campos_interes(self):
        for w in self.int_fields_frame.winfo_children():
            w.destroy()
        self.int_inputs = {}

    def mostrar_campos_interes(self, event=None):
        self.limpiar_campos_interes()
        tipo = self.int_tipo.get()
        calc = self.int_calculo.get()

        # Campos comunes de tiempo (los añadimos solo cuando correspondan)
        campos_tiempo = ["Años", "Meses", "Semanas", "Días"]

        campos = []
        # INTERÉS SIMPLE
        if tipo == "Interés Simple":
            if calc == "Valor Futuro (VF)":
                campos = ["Capital (VP)", "Tasa (i)"] + campos_tiempo
            elif calc == "Valor Presente (VP)":
                campos = ["Valor Futuro (VF)", "Tasa (i)"] + campos_tiempo
            elif calc == "Interés generado (I)":
                campos = ["Capital (VP)", "Tasa (i)"] + campos_tiempo
            elif calc == "Tasa de interés (i)":
                campos = ["Valor Futuro (VF)", "Capital (VP)"] + campos_tiempo
            elif calc == "Tiempo (t)":
                campos = ["Valor Futuro (VF)", "Capital (VP)", "Tasa (i)"]
        # VARIACIÓN DE TASAS
        elif tipo == "Interés Simple con Variación de Tasas":
            campos = ["Capital (VP)", "Número de tramos"]
            # aviso explicativo
            ttk.Label(self.int_fields_frame, text="(Cada tramo: tasa % y días)").grid(row=0, column=2, sticky="w")
        # INTERÉS COMPUESTO
        elif tipo == "Interés Compuesto":
            if calc == "Valor Futuro (VF)":
                campos = ["Capital (VP)", "Tasa (i)"] + campos_tiempo
            elif calc == "Valor Presente (VP)":
                campos = ["Valor Futuro (VF)", "Tasa (i)"] + campos_tiempo

        # Generar campos
        r = 0
        for c in campos:
            ttk.Label(self.int_fields_frame, text=c + ":").grid(row=r, column=0, sticky="w")
            e = ttk.Entry(self.int_fields_frame, width=20)
            e.grid(row=r, column=1)
            self.int_inputs[c] = e
            r += 1

        # Si es variación de tasas → tabla dinámica de tramos
        if tipo == "Interés Simple con Variación de Tasas":
            self.var_tramos_frame = ttk.LabelFrame(self.int_fields_frame, text="Tramos")
            self.var_tramos_frame.grid(row=r, column=0, columnspan=4, pady=8, sticky="w")
            self.int_inputs["tabla_tramos"] = []
            # vinculamos la actualización al evento FocusOut del campo 'Número de tramos'
            n_entry = self.int_inputs.get("Número de tramos")
            if n_entry:
                n_entry.bind("<FocusOut>", self.actualizar_tabla_tramos)

    def actualizar_tabla_tramos(self, event=None):
        # limpia la subtabla
        for w in getattr(self, "var_tramos_frame", ttk.Frame()).winfo_children():
            w.destroy()
        self.int_inputs["tabla_tramos"] = []

        # leer n
        try:
            n_entry = self.int_inputs.get("Número de tramos")
            n = int(n_entry.get()) if n_entry and n_entry.get().strip() != "" else 0
        except Exception:
            messagebox.showwarning("Número inválido", "Número de tramos inválido. Ingrese un entero.")
            return

        if n <= 0:
            return

        for i in range(n):
            ttk.Label(self.var_tramos_frame, text=f"Tramo {i+1} - tasa (%) :").grid(row=i, column=0, sticky="w")
            tasa_e = ttk.Entry(self.var_tramos_frame, width=10)
            tasa_e.grid(row=i, column=1, padx=4)
            ttk.Label(self.var_tramos_frame, text="días:").grid(row=i, column=2, sticky="w", padx=(8,0))
            dias_e = ttk.Entry(self.var_tramos_frame, width=10)
            dias_e.grid(row=i, column=3, padx=4)
            self.int_inputs["tabla_tramos"].append((tasa_e, dias_e))

    def convertir_tiempo(self):
        # lee campos de tiempo que existan; devuelve años float (puede ser 0)
        try:
            años = float(self.int_inputs.get("Años", ttk.Entry()).get() or 0)
        except Exception:
            años = 0.0
        try:
            meses = float(self.int_inputs.get("Meses", ttk.Entry()).get() or 0)
        except Exception:
            meses = 0.0
        try:
            semanas = float(self.int_inputs.get("Semanas", ttk.Entry()).get() or 0)
        except Exception:
            semanas = 0.0
        try:
            dias = float(self.int_inputs.get("Días", ttk.Entry()).get() or 0)
        except Exception:
            dias = 0.0

        total_dias = años*360 + meses*30 + semanas*7 + dias
        return total_dias / 360.0

    def limpiar_interes(self):
        # limpiar todos los inputs del frame de interés (incluida tabla de tramos)
        for k, entry in list(self.int_inputs.items()):
            if k == "tabla_tramos":
                for (t_e, d_e) in entry:
                    try:
                        t_e.delete(0, "end"); d_e.delete(0, "end")
                    except:
                        pass
                self.int_inputs["tabla_tramos"] = []
            else:
                try:
                    entry.delete(0, "end")
                except:
                    pass
        self.int_reporte.configure(state="normal")
        self.int_reporte.delete("1.0", "end")
        self.int_reporte.configure(state="disabled")

    def calcular_interes(self):
        tipo = self.int_tipo.get()
        calc = self.int_calculo.get()

        # helper para leer valor; devuelve None si no válido
        def get_val(key):
            e = self.int_inputs.get(key)
            if not e: 
                return None
            s = e.get().strip()
            if s == "":
                return None
            try:
                return float(s)
            except:
                return None

        # helper para mostrar resultado
        self.int_reporte.configure(state="normal")
        self.int_reporte.delete("1.0", "end")
        lines = []

        # lectura básica
        VP = get_val("Capital (VP)")
        i = get_val("Tasa (i)")
        VF = get_val("Valor Futuro (VF)")
        I = get_val("Interés (I)")
        t_from_fields = self.convertir_tiempo()  # años (puede ser 0)

        # VALIDACIONES INICIALES por caso
        # INTERÉS SIMPLE
        if tipo == "Interés Simple":
            if calc == "Valor Futuro (VF)":
                if VP is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese Capital (VP) y Tasa (i).")
                    return
                # si tiempo está ausente (todos ceros) solicita confirmación
                if t_from_fields == 0:
                    if not messagebox.askyesno("Tiempo = 0", "No ingresaste tiempo (resultado inmediato). ¿Continuar con t=0?"):
                        return
                VF_calc = VP * (1 + i * t_from_fields)
                lines.append(f"Tipo: Interés Simple — Valor Futuro (VF)")
                lines.append(f"VP = {VP:.2f}, i = {i:.6f}, t = {t_from_fields:.6f} años")
                lines.append(f"VF = {VF_calc:.2f}")
            elif calc == "Valor Presente (VP)":
                if VF is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese Valor Futuro y Tasa (i).")
                    return
                if t_from_fields is None:
                    messagebox.showwarning("Faltan datos", "Ingrese el tiempo.")
                    return
                denom = (1 + i * t_from_fields)
                if denom == 0:
                    messagebox.showerror("Error", "División por cero en cómputo de VP (1 + i*t = 0).")
                    return
                VP_calc = VF / denom
                lines.append(f"Tipo: Interés Simple — Valor Presente (VP)")
                lines.append(f"VF = {VF:.2f}, i = {i if i is not None else 'N/A'}, t = {t_from_fields:.6f} años")
                lines.append(f"VP = {VP_calc:.2f}")
            elif calc == "Interés generado (I)":
                if VP is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese Capital (VP) y Tasa (i).")
                    return
                I_calc = VP * i * t_from_fields
                lines.append(f"Interés generado = {I_calc:.2f}")
            elif calc == "Tasa de interés (i)":
                if VP is None or VF is None:
                    messagebox.showwarning("Faltan datos", "Ingrese VP y VF.")
                    return
                if t_from_fields == 0:
                    messagebox.showerror("Error", "No puede calcularse i con t = 0.")
                    return
                i_calc = (VF / VP - 1) / t_from_fields
                lines.append(f"Tasa i = {i_calc:.6f} ({i_calc*100:.4f}%)")
            elif calc == "Tiempo (t)":
                if VP is None or VF is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese VP, VF y Tasa (i).")
                    return
                if i == 0:
                    messagebox.showerror("Error", "No puede calcularse tiempo con i = 0.")
                    return
                t_calc = (VF / VP - 1) / i
                lines.append(f"Tiempo t = {t_calc:.6f} años")
            else:
                lines.append("Cálculo no reconocido para Interés Simple.")

        # INTERÉS COMPUESTO
        elif tipo == "Interés Compuesto":
            if calc == "Valor Futuro (VF)":
                if VP is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese VP y Tasa (i).")
                    return
                VF_calc = VP * (1 + i) ** t_from_fields
                lines.append(f"Interés Compuesto — Valor Futuro")
                lines.append(f"VF = {VF_calc:.2f}")
            elif calc == "Valor Presente (VP)":
                if VF is None or i is None:
                    messagebox.showwarning("Faltan datos", "Ingrese VF y Tasa (i).")
                    return
                denom = (1 + i) ** t_from_fields
                if denom == 0:
                    messagebox.showerror("Error", "División por cero en cálculo de VP compuesto.")
                    return
                VP_calc = VF / denom
                lines.append(f"Interés Compuesto — Valor Presente")
                lines.append(f"VP = {VP_calc:.2f}")
            else:
                lines.append("Cálculo no reconocido para Interés Compuesto.")

        # VARIACIÓN DE TASAS (simple: suma de intereses por tramo sobre el capital)
        elif tipo == "Interés Simple con Variación de Tasas":
            # necesitamos capital y tabla de tramos
            if VP is None:
                messagebox.showwarning("Faltan datos", "Ingrese Capital (VP).")
                return
            tabla = self.int_inputs.get("tabla_tramos", [])
            if not tabla:
                messagebox.showwarning("Faltan tramos", "Defina Número de tramos y las tasas/días.")
                return
            total_interes = 0.0
            detalle = []
            for idx, (t_e, d_e) in enumerate(tabla):
                try:
                    tasa_pct = float(t_e.get().strip()) if t_e.get().strip() != "" else 0.0
                    dias = float(d_e.get().strip()) if d_e.get().strip() != "" else 0.0
                except:
                    messagebox.showwarning("Entrada inválida", f"Tramo {idx+1} tiene entradas inválidas.")
                    return
                tasa = tasa_pct / 100.0
                años_tramo = dias / 360.0
                interes_tramo = VP * tasa * años_tramo
                total_interes += interes_tramo
                detalle.append(f"Tramo {idx+1}: tasa={tasa_pct:.4f}%, días={dias:.0f} → interés={interes_tramo:.2f}")
            VF_calc = VP + total_interes
            lines.append("Interés Simple con Variación de Tasas (sumatoria por tramos, interés simple sobre VP):")
            lines.append(f"VP = {VP:.2f}")
            lines.extend(detalle)
            lines.append(f"Interés total = {total_interes:.2f}")
            lines.append(f"Valor Futuro (VF) = {VF_calc:.2f}")
        else:
            lines.append("Seleccione un tipo de interés y cálculo válidos.")

        # mostrar resultado
        self.int_reporte.insert("end", "\n".join(lines))
        self.int_reporte.configure(state="disabled")




    # ------------------ TAB RATIOS (user code adapted) ------------------
    def build_tab_ratios(self, frm):
        ttk.Label(frm, text="Calculadora de Ratios Financieros", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=4, pady=8)
        ttk.Label(frm, text="Categoría:").grid(row=1, column=0, sticky='w', padx=5)
        self.cat_cb = ttk.Combobox(frm, values=list(CATEGORIES.keys()), state="readonly", width=28)
        self.cat_cb.grid(row=1, column=1, sticky='w')
        self.cat_cb.bind("<<ComboboxSelected>>", self.sel_cat)

        ttk.Label(frm, text="Ratio:").grid(row=2, column=0, sticky='w', padx=5)
        self.ratio_cb = ttk.Combobox(frm, state="readonly", width=40)
        self.ratio_cb.grid(row=2, column=1, sticky='w')
        self.ratio_cb.bind("<<ComboboxSelected>>", self.sel_ratio)

        ttk.Label(frm, text="(Los datos pueden provenir de la pestaña 'Vertical + Horizontal' si ya se llenaron.)", foreground='gray').grid(row=3, column=0, columnspan=4, sticky='w', padx=6)

        # entries
        self.r_entries = {}
        row = 4
        for field in sorted({f for lst in FIELDS.values() for f in lst}):
            ttk.Label(frm, text=f"{field}:").grid(row=row, column=0, sticky='w', padx=6, pady=2)
            e = ttk.Entry(frm, width=22, state='disabled')
            e.grid(row=row, column=1, sticky='w', padx=6, pady=2)
            self.r_entries[field] = e
            row += 1
        btn_row = row + 1
        ttk.Button(frm, text="Calcular", command=self.calc_ratio).grid(row=btn_row, column=0, padx=8, pady=10, sticky='w')
        ttk.Button(frm, text="Usar datos compartidos (copiar desde Análisis)", command=self.load_shared_to_ratios).grid(row=btn_row, column=1, padx=8, pady=10, sticky='w')
        ttk.Button(frm, text="Limpiar", command=self.reset_ratios).grid(row=btn_row, column=2, padx=8, pady=10, sticky='w')

        ttk.Label(frm, text="Reporte:", font=("Segoe UI", 10, "bold")).grid(row=btn_row+1, column=0, columnspan=3, sticky='w')
        self.ratios_report = scrolledtext.ScrolledText(frm, width=115, height=18, wrap='word')
        self.ratios_report.grid(row=btn_row+2, column=0, columnspan=4, padx=6, pady=6)
        self.ratios_report.configure(state='disabled')

    def sel_cat(self, e=None):
        cat = self.cat_cb.get()
        self.ratio_cb['values'] = CATEGORIES.get(cat, [])
        self.ratio_cb.set('')
        self.disable_all_ratio_entries()

    def sel_ratio(self, e=None):
        r = self.ratio_cb.get()
        self.disable_all_ratio_entries()
        for f in FIELDS.get(r, []):
            ent = self.r_entries.get(f)
            if ent:
                ent.configure(state='normal')

    def disable_all_ratio_entries(self):
        for ent in self.r_entries.values():
            ent.configure(state='disabled')
            ent.delete(0, 'end')
        self.clear_ratios_report()

    def load_shared_to_ratios(self):
        # Copy values from shared data if present
        # Map some common keys
        mapping = {
            "Activo corriente": ("bg_cur", "Activo corriente"),
            "Pasivo corriente": ("bg_cur", "Pasivo corriente"),
            "Inventarios": ("bg_cur", "Inventarios"),
            "Inventario promedio": ("computed", "Inventario promedio"),
            "Costo de ventas": ("er_cur", "Costo de ventas"),
            "Ventas netas": ("er_cur", "Ventas netas"),
            "Cuentas por cobrar": ("bg_cur", "Cuentas por cobrar"),
            "Activo total": ("bg_cur", "Activo total"),
            "Pasivo total": ("bg_cur", "Pasivo total"),
            "Patrimonio": ("bg_cur", "Patrimonio"),
            "UAII": ("er_cur", "Utilidad operativa"),
            "Gastos por intereses": ("er_cur", "Gastos financieros"),
            "Utilidad neta": ("er_cur", "Utilidad neta")
        }
        # Compute inventory average if possible
        invA_init = safe_float(self.inv_init_A.get()) if hasattr(self, 'inv_init_A') else 0.0
        invA_final = safe_float(self.inv_final_A.get()) if hasattr(self, 'inv_final_A') else 0.0
        inv_prom = None
        if invA_init or invA_final:
            inv_prom = (invA_init + invA_final)/2.0
        else:
            inv_prom = self.bg_current.get('Inventarios', 0)

        for field, ent in self.r_entries.items():
            if field == "Inventario promedio":
                if inv_prom is not None:
                    ent.configure(state='normal')
                    ent.delete(0, 'end')
                    ent.insert(0, str(inv_prom))
                    ent.configure(state='normal')
                continue
            mapv = mapping.get(field)
            if mapv:
                source, key = mapv
                val = 0.0
                if source == "bg_cur":
                    val = self.bg_current.get(key, 0.0)
                elif source == "er_cur":
                    val = self.er_current.get(key, 0.0)
                elif source == "computed" and key == "Inventario promedio":
                    val = inv_prom or 0.0
                ent.configure(state='normal')
                ent.delete(0, 'end')
                ent.insert(0, str(val))
                ent.configure(state='normal')

        messagebox.showinfo("Listo", "Se copiaron los datos compartidos (si estaban presentes).")

    def calc_ratio(self):
        r = self.ratio_cb.get()
        if not r:
            messagebox.showwarning("Atención", "Seleccione categoría y ratio.")
            return
        vals = {}
        for f in FIELDS.get(r, []):
            s = self.r_entries[f].get()
            if s.strip() == "":
                # try shared data as fallback
                vals[f] = self.fallback_shared_value_for_field(f)
            else:
                vals[f] = safe_float(s)
        val = calcular(r, vals)
        inter = interpretar(r, val)
        lines = [f"RATIO: {r}", "-"*48]
        if r in ['Pasivo/Activo','Margen neto','ROA','ROE']:
            lines.append(f"Resultado: {fmt_pct(val)}")
        elif "Período" in r:
            lines.append(f"Resultado: {fmt_num(val)} días")
        else:
            lines.append(f"Resultado: {fmt_num(val)}")
        lines.append("")
        lines.append("Cálculo usado (valores):")
        for f in FIELDS.get(r, []):
            lines.append(f" - {f}: {fmt_num(vals.get(f, float('nan')))}")
        lines.append("")
        lines.append("Interpretación detallada:")
        lines.append(inter)
        # suggestions (basic)
        lines.append("")
        lines.append("Sugerencia práctica:")
        # Add practical suggestion based on ratio type
        if r in ["Razón corriente", "Prueba ácida", "Capital de trabajo"]:
            lines.append(" - Revisar ciclo de caja y negociar plazos con proveedores si liquidez baja.")
        if r in ["Rotación inventarios", "Período promedio de inventarios"]:
            lines.append(" - Optimizar niveles de inventario: JIT, descuentos, promociones si exceso.")
        if r in ["Rotación CxC", "Período promedio de cobro"]:
            lines.append(" - Revisar condiciones de crédito y políticas de cobranza.")
        if r in ["Pasivo/Activo", "Deuda/Patrimonio", "Cobertura intereses"]:
            lines.append(" - Evaluar refinanciación o reducción de deuda si nivel alto.")
        if r in ["Margen neto","ROA","ROE"]:
            lines.append(" - Analizar mix de productos y control de costos para mejorar rentabilidad.")
        # show
        self.ratios_report.configure(state='normal')
        self.ratios_report.delete('1.0', tk.END)
        self.ratios_report.insert(tk.END, "\n".join(lines))
        self.ratios_report.configure(state='disabled')

    def fallback_shared_value_for_field(self, field):
        # Try to return a sensible shared value if user didn't enter for ratios tab
        mapping_shared = {
            "Activo corriente": lambda: self.bg_current.get("Activo corriente", 0),
            "Pasivo corriente": lambda: self.bg_current.get("Pasivo corriente", 0),
            "Inventarios": lambda: self.bg_current.get("Inventarios", 0),
            "Inventario promedio": lambda: ((safe_float(self.inv_init_A.get()) + safe_float(self.inv_final_A.get()))/2.0) if hasattr(self, 'inv_init_A') and (self.inv_init_A.get().strip() or self.inv_final_A.get().strip()) else self.bg_current.get("Inventarios", 0),
            "Costo de ventas": lambda: self.er_current.get("Costo de ventas", 0),
            "Ventas netas": lambda: self.er_current.get("Ventas netas", 0),
            "Cuentas por cobrar": lambda: self.bg_current.get("Cuentas por cobrar", 0),
            "Activo total": lambda: self.bg_current.get("Activo total", 0),
            "Pasivo total": lambda: self.bg_current.get("Pasivo total", 0),
            "Patrimonio": lambda: self.bg_current.get("Patrimonio", 0),
            "UAII": lambda: self.er_current.get("Utilidad operativa", 0),
            "Gastos por intereses": lambda: self.er_current.get("Gastos financieros", 0),
            "Utilidad neta": lambda: self.er_current.get("Utilidad neta", 0)
        }
        fn = mapping_shared.get(field)
        if fn:
            return fn()
        return 0.0

    def reset_ratios(self):
        self.cat_cb.set('')
        self.ratio_cb.set('')
        self.disable_all_ratio_entries()

    def clear_ratios_report(self):
        self.ratios_report.configure(state='normal')
        self.ratios_report.delete('1.0', tk.END)
        self.ratios_report.configure(state='disabled')

    # ------------------ TAB DUPONT ------------------
    def build_tab_dupont(self, frm):
        ttk.Label(frm, text="Análisis DuPont (Descomposición del ROE)", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=6, pady=8, sticky='w')
        ttk.Label(frm, text="(Usa los datos ingresados en 'Vertical + Horizontal')", foreground='gray').grid(row=1, column=0, columnspan=6, sticky='w', padx=6)
        ttk.Button(frm, text="Actualizar desde datos compartidos", command=self.update_shared_from_analysis).grid(row=2, column=0, padx=6, pady=8, sticky='w')
        ttk.Button(frm, text="Calcular DuPont", command=self.calc_dupont).grid(row=2, column=1, padx=6, pady=8, sticky='w')

        ttk.Label(frm, text="Reporte DuPont:", font=("Segoe UI", 10, "bold")).grid(row=3, column=0, columnspan=6, sticky='w', padx=6)
        self.dupont_report = scrolledtext.ScrolledText(frm, width=120, height=26, wrap='word')
        self.dupont_report.grid(row=4, column=0, columnspan=6, padx=6, pady=6)
        self.dupont_report.configure(state='disabled')

    def update_shared_from_analysis(self):
        # Pull values from BG and ER entries (if they exist)
        # They are stored in self.bg_current, self.er_current by Analysis tab operations
        # Here just ensure current/prior have the best available values
        # (This function is mostly a convenience; generate_full_analysis already fills shared vars)
        messagebox.showinfo("Actualizar", "Si ya ingresaste datos en 'Vertical + Horizontal', el sistema los usará automáticamente al calcular DuPont.")
        # nothing else needed because generate_full_analysis sets self.current, etc.

    def calc_dupont(self):
        if not self.current:
            messagebox.showwarning("Datos faltantes", "No hay datos cargados. Ve a la pestaña 'Vertical + Horizontal' y complete los estados para el año actual y previo, luego presiona 'Generar análisis' para llenar la base compartida.")
            return
        res = calculate_dupont(self.current)
        # Build detailed interpretation
        lines = []
        lines.append("----- RESUMEN EJECUTIVO -----")
        # Simple executive summary heuristic
        summary = self.build_executive_summary(res)
        lines.append(summary)
        lines.append("")
        lines.append("----- DESCOMPOSICIÓN DuPont -----")
        lines.append(f"Margen neto: {fmt_pct(res['margen_neto'])}  → indica cuánto queda de cada unidad monetaria de venta como utilidad.")
        lines.append(f"Rotación de activos: {fmt_num(res['rot_activos'])} → cuántas veces los activos generan ventas.")
        lines.append(f"Apalancamiento (Activo / Patrimonio): {fmt_num(res['equity_multiplier'])} → cuánto apalanca el patrimonio.")
        lines.append(f"ROE (DuPont): {fmt_pct(res['roe_dupont'])}")
        lines.append(f"ROA: {fmt_pct(res['roa'])}")
        lines.append(f"ROE directo: {fmt_pct(res['roe'])}")
        lines.append("")
        lines.append("Contribución aproximada al ROE :")
        cm = res['contribuciones'].get('margen_pct')
        cr = res['contribuciones'].get('rot_pct')
        ca = res['contribuciones'].get('apal_pct')
        lines.append(f" - Margen: {fmt_num(cm)}%")
        lines.append(f" - Rotación: {fmt_num(cr)}%")
        lines.append(f" - Apalancamiento: {fmt_num(ca)}%")
        lines.append("")
        # Interpretation per factor
        lines.append("Interpretación detallada por factor:")
        # Margen
        if res['margen_neto'] != res['margen_neto']:
            lines.append(" - Margen: no disponible.")
        else:
            if res['margen_neto'] < 0.05:
                lines.append(" - Margen bajo: la empresa obtiene poca utilidad por venta; revisar precios y costos.")
            elif res['margen_neto'] <= 0.15:
                lines.append(" - Margen adecuado: buen control de costos relativo a la industria.")
            else:
                lines.append(" - Margen alto: muy buena gestión de costos o ventaja competitiva en precios.")
        # Rotación
        if res['rot_activos'] != res['rot_activos']:
            lines.append(" - Rotación activos: no disponible.")
        else:
            if res['rot_activos'] < 1:
                lines.append(" - Rotación de activos baja: activos no están generando suficientes ventas; revisar inversiones.")
            else:
                lines.append(" - Rotación de activos adecuada: los activos están siendo aprovechados para generar ventas.")
        # Apalancamiento
        if res['equity_multiplier'] != res['equity_multiplier']:
            lines.append(" - Apalancamiento: no disponible.")
        else:
            if res['equity_multiplier'] > 2.5:
                lines.append(" - Apalancamiento alto: el ROE puede estar elevado por deuda; revisar riesgo financiero.")
            else:
                lines.append(" - Apalancamiento moderado: estructura de capital equilibrada.")
        lines.append("")
        # Recommendations derived from DuPont
        lines.append("Recomendaciones (según DuPont):")
        if res['contribuciones'].get('margen_pct') != res['contribuciones'].get('margen_pct'):
            lines.append(" - No hay información suficiente para recomendaciones específicas.")
        else:
            # Identify main driver
            pieces = [('Margen', cm), ('Rotación', cr), ('Apalancamiento', ca)]
            # if nan, set 0
            pieces = [(name, (val if val==val else -9999)) for name,val in pieces]
            pieces_sorted = sorted(pieces, key=lambda x: x[1], reverse=True)
            main_driver = pieces_sorted[0][0]
            lines.append(f" - El factor que más está impulsando el ROE parece ser: {main_driver}.")
            if main_driver == 'Margen':
                lines.append("   * Foco: mejorar mix de producto, precios y eficiencia de costos.")
            elif main_driver == 'Rotación':
                lines.append("   * Foco: mejorar uso de activos, optimizar inventarios y aumentar ventas por activo.")
            else:
                lines.append("   * Foco: revisar estructura de deuda y coste financiero.")
        # show
        self.dupont_report.configure(state='normal')
        self.dupont_report.delete('1.0', tk.END)
        self.dupont_report.insert(tk.END, "\n".join(lines))
        self.dupont_report.configure(state='disabled')

    def build_executive_summary(self, dupont_res):
        # Short automatic summary based on indicators
        parts = []
        # Profitability
        if dupont_res['roe'] == dupont_res['roe']:
            if dupont_res['roe'] > 0.15:
                parts.append("Rentabilidad (ROE) alta.")
            elif dupont_res['roe'] >= 0.08:
                parts.append("Rentabilidad moderada.")
            else:
                parts.append("Rentabilidad baja.")
        # Debt
        if dupont_res['razon_endeudamiento'] == dupont_res['razon_endeudamiento']:
            if dupont_res['razon_endeudamiento'] < 0.4:
                parts.append("Bajo nivel de endeudamiento.")
            elif dupont_res['razon_endeudamiento'] <= 0.6:
                parts.append("Endeudamiento moderado.")
            else:
                parts.append("Endeudamiento alto; riesgo financiero.")
        # Efficiency
        if dupont_res['rot_activos'] == dupont_res['rot_activos']:
            if dupont_res['rot_activos'] >= 1:
                parts.append("Eficiencia operativa razonable (activos generan ventas).")
            else:
                parts.append("Activos infrautilizados.")
        return " ".join(parts)

    # ------------------ TAB VERTICAL + HORIZONTAL ------------------
    def build_tab_vh(self, frm):
        ttk.Label(frm, text="Análisis Vertical y Horizontal (Estados financieros A y P)", font=("Segoe UI", 14, "bold")).grid(row=0, column=0, columnspan=6, sticky='w', pady=8)
        ttk.Label(frm, text="Ingrese los estados para Año Actual (A) y Año Previo (P). Inventario promedio se calcula automáticamente si provees inventario inicial y final.", foreground='gray').grid(row=1, column=0, columnspan=6, sticky='w', padx=6)

        # Top keys
        keys = ['Ventas netas', 'Utilidad neta', 'Activo total', 'Patrimonio', 'Pasivo total', 'Costo de ventas']
        self.entries_current = {}
        self.entries_prior = {}
        ttk.Label(frm, text="Cuenta").grid(row=2, column=0, padx=4, pady=4)
        ttk.Label(frm, text="Año actual (A)").grid(row=2, column=1, padx=4, pady=4)
        ttk.Label(frm, text="Año previo (P)").grid(row=2, column=2, padx=4, pady=4)
        r = 3
        for k in keys:
            ttk.Label(frm, text=k).grid(row=r, column=0, sticky='w', padx=4, pady=2)
            a = ttk.Entry(frm, width=18); a.grid(row=r, column=1, padx=4, pady=2)
            p = ttk.Entry(frm, width=18); p.grid(row=r, column=2, padx=4, pady=2)
            self.entries_current[k] = a
            self.entries_prior[k] = p
            r += 1

        # Inventarios inicial/final A y P
        ttk.Label(frm, text="Inventario inicial A:").grid(row=r, column=0, sticky='w', padx=4); self.inv_init_A = ttk.Entry(frm, width=18); self.inv_init_A.grid(row=r, column=1, padx=4); ttk.Label(frm, text="Inventario final A:").grid(row=r, column=2, sticky='w', padx=4); self.inv_final_A = ttk.Entry(frm, width=18); self.inv_final_A.grid(row=r, column=3, padx=4); r+=1
        ttk.Label(frm, text="Inventario inicial P:").grid(row=r, column=0, sticky='w', padx=4); self.inv_init_P = ttk.Entry(frm, width=18); self.inv_init_P.grid(row=r, column=1, padx=4); ttk.Label(frm, text="Inventario final P:").grid(row=r, column=2, sticky='w', padx=4); self.inv_final_P = ttk.Entry(frm, width=18); self.inv_final_P.grid(row=r, column=3, padx=4); r+=1

        # Balance General detailed
        ttk.Label(frm, text="Balance General (A / P)", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=6, sticky='w', pady=6); r+=1
        bg_accounts = ['Activo total','Caja y bancos','Cuentas por cobrar','Inventarios','Activos fijos netos','Pasivo corriente','Pasivo no corriente','Pasivo total','Patrimonio']
        self.bg_cur = {}; self.bg_pr = {}
        ttk.Label(frm, text="Cuenta").grid(row=r, column=0, sticky='w'); ttk.Label(frm, text="Actual (A)").grid(row=r, column=1); ttk.Label(frm, text="Previo (P)").grid(row=r, column=2); r+=1
        for acc in bg_accounts:
            ttk.Label(frm, text=acc+":").grid(row=r, column=0, sticky='w', padx=4, pady=2)
            ec = ttk.Entry(frm, width=18); ec.grid(row=r, column=1, padx=4, pady=2)
            ep = ttk.Entry(frm, width=18); ep.grid(row=r, column=2, padx=4, pady=2)
            self.bg_cur[acc] = ec; self.bg_pr[acc] = ep
            r += 1

        # Estado de Resultados detailed
        ttk.Label(frm, text="Estado de Resultados (A / P)", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=6, sticky='w', pady=6); r+=1
        er_accounts = ['Ventas netas','Costo de ventas','Utilidad bruta','Gastos operativos','Utilidad operativa','Gastos financieros','Utilidad antes de impuestos','Utilidad neta']
        self.er_cur = {}; self.er_pr = {}
        ttk.Label(frm, text="Cuenta").grid(row=r, column=0, sticky='w'); ttk.Label(frm, text="Actual (A)").grid(row=r, column=1); ttk.Label(frm, text="Previo (P)").grid(row=r, column=2); r+=1
        for acc in er_accounts:
            ttk.Label(frm, text=acc+":").grid(row=r, column=0, sticky='w', padx=4, pady=2)
            ec = ttk.Entry(frm, width=18); ec.grid(row=r, column=1, padx=4, pady=2)
            ep = ttk.Entry(frm, width=18); ep.grid(row=r, column=2, padx=4, pady=2)
            self.er_cur[acc] = ec; self.er_pr[acc] = ep
            r += 1

        ttk.Button(frm, text="Generar Análisis (llenar datos compartidos)", command=self.generate_full_analysis).grid(row=r, column=0, padx=6, pady=12, sticky='w')
        ttk.Button(frm, text="Limpiar campos", command=self.clear_vh_fields).grid(row=r, column=1, padx=6, pady=12, sticky='w')
        r+=1
        ttk.Label(frm, text="Reporte (Resumen ejecutivo + Vertical + Horizontal):", font=("Segoe UI", 10, "bold")).grid(row=r, column=0, columnspan=6, sticky='w')
        r+=1
        self.vh_report = scrolledtext.ScrolledText(frm, width=120, height=28, wrap='word')
        self.vh_report.grid(row=r, column=0, columnspan=6, padx=6, pady=6)
        self.vh_report.configure(state='disabled')

    def generate_full_analysis(self):
        try:
            # Read top summary current/prior
            for k, ent in self.entries_current.items():
                self.current[k] = safe_float(ent.get())
            for k, ent in self.entries_prior.items():
                self.prior[k] = safe_float(ent.get())

            # Read BG and ER
            for k, ent in self.bg_cur.items():
                self.bg_current[k] = safe_float(ent.get())
            for k, ent in self.bg_pr.items():
                self.bg_prior[k] = safe_float(ent.get())
            for k, ent in self.er_cur.items():
                self.er_current[k] = safe_float(ent.get())
            for k, ent in self.er_pr.items():
                self.er_prior[k] = safe_float(ent.get())

            # Inventory promedio calculation
            invA_init = safe_float(self.inv_init_A.get())
            invA_final = safe_float(self.inv_final_A.get())
            if invA_init or invA_final:
                invA_prom = (invA_init + invA_final) / 2.0
            else:
                invA_prom = self.bg_current.get('Inventarios', 0.0)
            invP_init = safe_float(self.inv_init_P.get())
            invP_final = safe_float(self.inv_final_P.get())
            if invP_init or invP_final:
                invP_prom = (invP_init + invP_final) / 2.0
            else:
                invP_prom = self.bg_prior.get('Inventarios', 0.0)

            # set in current/prior dictionaries for downstream use
            self.current['Inventario promedio'] = invA_prom
            self.prior['Inventario promedio'] = invP_prom

            # Ensure fallback: if top summary missing, use ER/BG detailed
            if self.current.get('Ventas netas',0)==0:
                self.current['Ventas netas'] = self.er_current.get('Ventas netas', 0)
            if self.current.get('Utilidad neta',0)==0:
                self.current['Utilidad neta'] = self.er_current.get('Utilidad neta', 0)
            if self.current.get('Costo de ventas',0)==0:
                self.current['Costo de ventas'] = self.er_current.get('Costo de ventas', 0)
            if self.current.get('Activo total',0)==0:
                self.current['Activo total'] = self.bg_current.get('Activo total', 0)
            if self.current.get('Patrimonio',0)==0:
                self.current['Patrimonio'] = self.bg_current.get('Patrimonio', 0)
            if self.current.get('Pasivo total',0)==0:
                self.current['Pasivo total'] = self.bg_current.get('Pasivo total', 0)

            # same for prior
            if self.prior.get('Ventas netas',0)==0:
                self.prior['Ventas netas'] = self.er_prior.get('Ventas netas', 0)
            if self.prior.get('Utilidad neta',0)==0:
                self.prior['Utilidad neta'] = self.er_prior.get('Utilidad neta', 0)
            if self.prior.get('Costo de ventas',0)==0:
                self.prior['Costo de ventas'] = self.er_prior.get('Costo de ventas', 0)
            if self.prior.get('Activo total',0)==0:
                self.prior['Activo total'] = self.bg_prior.get('Activo total', 0)
            if self.prior.get('Patrimonio',0)==0:
                self.prior['Patrimonio'] = self.bg_prior.get('Patrimonio', 0)
            if self.prior.get('Pasivo total',0)==0:
                self.prior['Pasivo total'] = self.bg_prior.get('Pasivo total', 0)

            # compute dupont and analyses
            dupont_res = calculate_dupont(self.current)
            vert_bg_cur = analisis_vertical(self.bg_current, "BG")
            vert_er_cur = analisis_vertical(self.er_current, "ER")
            horiz_bg = analisis_horizontal(self.bg_current, self.bg_prior)
            horiz_er = analisis_horizontal(self.er_current, self.er_prior)

            # Build report with summary executive
            rpt = []
            rpt.append("----- RESUMEN EJECUTIVO -----")
            # simple heuristics
            summary_parts = []
            if dupont_res['roe'] == dupont_res['roe']:
                if dupont_res['roe'] > 0.15:
                    summary_parts.append(f"ROE = {fmt_pct(dupont_res['roe'])}: alta rentabilidad.")
                elif dupont_res['roe'] >= 0.08:
                    summary_parts.append(f"ROE = {fmt_pct(dupont_res['roe'])}: rentabilidad moderada.")
                else:
                    summary_parts.append(f"ROE = {fmt_pct(dupont_res['roe'])}: rentabilidad baja.")
            if dupont_res['razon_endeudamiento'] == dupont_res['razon_endeudamiento']:
                if dupont_res['razon_endeudamiento'] < 0.4:
                    summary_parts.append("Endeudamiento bajo.")
                elif dupont_res['razon_endeudamiento'] <= 0.6:
                    summary_parts.append("Endeudamiento moderado.")
                else:
                    summary_parts.append("Endeudamiento alto.")
            if dupont_res['rot_activos'] == dupont_res['rot_activos']:
                if dupont_res['rot_activos'] >= 1:
                    summary_parts.append("Eficiencia en uso de activos aceptable.")
                else:
                    summary_parts.append("Activos infrautilizados.")
            rpt.append(" ".join(summary_parts))
            rpt.append("")

            # DuPont summary lines (like earlier)
            rpt.append("----- DU PONT (resumen rápido) -----")
            rpt.append(f"Margen neto: {fmt_pct(dupont_res['margen_neto'])}")
            rpt.append(f"Rotación activos: {fmt_num(dupont_res['rot_activos'])}")
            rpt.append(f"Apalancamiento (Activo/Patrimonio): {fmt_num(dupont_res['equity_multiplier'])}")
            rpt.append(f"ROE (DuPont): {fmt_pct(dupont_res['roe_dupont'])}")
            rpt.append("")

            # Vertical BG
            rpt.append("----- ANALISIS VERTICAL (Balance General - % sobre Activo) -----")
            if vert_bg_cur:
                for k,v in vert_bg_cur.items():
                    rpt.append(f" {k}: {fmt_pct(v)}")
            else:
                rpt.append(" No hay datos de Balance General para calcular análisis vertical.")
            rpt.append("")

            # Vertical ER
            rpt.append("----- ANALISIS VERTICAL (Estado de Resultados - % sobre Ventas) -----")
            if vert_er_cur:
                for k,v in vert_er_cur.items():
                    rpt.append(f" {k}: {fmt_pct(v)}")
            else:
                rpt.append(" No hay datos de Estado de Resultados para análisis vertical.")
            rpt.append("")

            # Horizontal
            rpt.append("----- ANALISIS HORIZONTAL (Variación A vs P) -----")
            if horiz_bg:
                rpt.append("Balance General (A vs P):")
                for k,v in horiz_bg.items():
                    rpt.append(f" {k}: {fmt_pct(v)}")
            else:
                rpt.append(" No hay datos para análisis horizontal BG.")
            rpt.append("Estado de Resultados (A vs P):")
            if horiz_er:
                for k,v in horiz_er.items():
                    rpt.append(f" {k}: {fmt_pct(v)}")
            else:
                rpt.append(" No hay datos para análisis horizontal ER.")
            rpt.append("")

            # Fortalezas y debilidades (heurístico y algo más elaborado)
            rpt.append("----- FORTALEZAS Y DEBILIDADES -----")
            # Margen
            if dupont_res['margen_neto'] != dupont_res['margen_neto']:
                rpt.append(" - Margen: no disponible.")
            else:
                if dupont_res['margen_neto'] >= 0.10:
                    rpt.append(" - Margen: sólido. Buen control de costos o precio con poder de mercado.")
                elif dupont_res['margen_neto'] >= 0.05:
                    rpt.append(" - Margen: razonable, puede mejorarse con control de costos.")
                else:
                    rpt.append(" - Margen: bajo; revisar estructura de costos y política de precios.")
            # Rotación inventarios
            if dupont_res['rot_inventarios'] != dupont_res['rot_inventarios']:
                rpt.append(" - Rotación de inventarios: no disponible.")
            else:
                if dupont_res['rot_inventarios'] < 3:
                    rpt.append(" - Rotación de inventarios baja: exceso de stock o baja demanda.")
                else:
                    rpt.append(" - Rotación de inventarios adecuada.")
            # Apalancamiento
            if dupont_res['equity_multiplier'] > 2.5:
                rpt.append(" - Apalancamiento: alto; revisar estructura de financiamiento y coste de deuda.")
            else:
                rpt.append(" - Apalancamiento: controlado.")
            rpt.append("")

            # Tendencias clave
            vend_growth = horiz_er.get('Ventas netas', float('nan'))
            util_growth = horiz_er.get('Utilidad neta', float('nan'))
            rpt.append("----- TENDENCIAS -----")
            rpt.append(f" Ventas (A vs P): {fmt_pct(vend_growth)}")
            rpt.append(f" Utilidad neta (A vs P): {fmt_pct(util_growth)}")
            if vend_growth == vend_growth and util_growth == util_growth:
                if util_growth > vend_growth:
                    rpt.append(" Interpretación: la rentabilidad mejora más que las ventas (mejor gestión de costos).")
                elif util_growth < vend_growth:
                    rpt.append(" Interpretación: ventas crecen pero utilidad menos; vigilar margen.")
                else:
                    rpt.append(" Interpretación: margen estable frente a crecimiento de ventas.")
            else:
                rpt.append(" No hay datos suficientes para evaluar tendencias.")

            # Recomendación final
            rpt.append("")
            rpt.append("----- RECOMENDACIÓN (resumen) -----")
            debt_level = dupont_res['razon_endeudamiento']
            if dupont_res['roe'] > 0.15 and debt_level < 0.6 and dupont_res['margen_neto']>0:
                rpt.append(" Inversión favorable: buena rentabilidad y deuda controlada.")
            else:
                rpt.append(" Evaluar con cautela: considerar medidas para mejorar margen, reducir deuda o aumentar eficiencia.")

            # Show report and also store dupont_res to self for DuPont tab
            self.vh_report.configure(state='normal')
            self.vh_report.delete('1.0', tk.END)
            self.vh_report.insert(tk.END, "\n".join(rpt))
            self.vh_report.configure(state='disabled')

            # Also set shared 'current'/'prior' dicts already updated above (they are used by DuPont tab)
            # Copy totals for quick fallback in ratios
            # ensure bg_current and er_current remain available for ratios fallback
            # Done above by reading entries into self.bg_current/self.er_current

            messagebox.showinfo("Listo", "Análisis generado y datos base actualizados para las otras pestañas.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar análisis: {e}")

    def clear_vh_fields(self):
        for ent in self.entries_current.values(): ent.delete(0, 'end')
        for ent in self.entries_prior.values(): ent.delete(0, 'end')
        for ent in [self.inv_init_A, self.inv_final_A, self.inv_init_P, self.inv_final_P]:
            ent.delete(0, 'end')
        for ent in self.bg_cur.values(): ent.delete(0, 'end')
        for ent in self.bg_pr.values(): ent.delete(0, 'end')
        for ent in self.er_cur.values(): ent.delete(0, 'end')
        for ent in self.er_pr.values(): ent.delete(0, 'end')
        self.vh_report.configure(state='normal'); self.vh_report.delete('1.0', 'end'); self.vh_report.configure(state='disabled')
        # clear shared dicts
        self.current.clear(); self.prior.clear(); self.bg_current.clear(); self.bg_prior.clear(); self.er_current.clear(); self.er_prior.clear()
        messagebox.showinfo("Limpiar", "Campos y datos compartidos limpiados.")

# ------------------ RUN ------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = UnifiedApp(root)
    root.mainloop()
