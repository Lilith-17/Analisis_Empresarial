# =============================
# Librerías principales
# =============================
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import html
import math
import os
import datetime
import re
import numpy as np

# =============================
# Matplotlib para gráficos
# =============================
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# =============================
# ReportLab para PDF (DISEÑO PRO)
# =============================
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor

# ======================================================
# CONFIGURACIÓN DE COLORES Y ESTILO (PALETA PRO)
# ======================================================
COLOR_PRIMARIO = HexColor("#2C3E50")  # Azul Oscuro (Midnight Blue)
COLOR_SECUNDARIO = HexColor("#3498DB") # Azul Brillante
COLOR_ACENTO = HexColor("#E74C3C")    # Rojo Suave
COLOR_FONDO_TABLA = HexColor("#ECF0F1") # Gris muy claro
COLOR_TEXTO = HexColor("#34495E")     # Gris oscuro (mejor que negro puro)

# ======================================================
# CLASE DE CÁLCULO Y LÓGICA FINANCIERA
# ======================================================

class AnalisisFinanciero:
    def __init__(self, data=None):
        # Datos por defecto
        default_data = {
            "AC_2023": 2800.0, "AC_2024": 3800.0, "ANC_2023": 1450.0, "ANC_2024": 1850.0,
            "PC_2023": 550.0, "PC_2024": 1000.0, "PNC_2023": 700.0, "PNC_2024": 1000.0,
            "PN_2023": 3000.0, "PN_2024": 3650.0, "Caja_2023": 850.0, "Caja_2024": 1100.0,
            "Clientes_2023": 1200.0, "Clientes_2024": 1600.0, "InvCP_2023": 300.0, "InvCP_2024": 500.0,
            "Ingresos_2023": 8500.0, "Ingresos_2024": 11200.0, "Costo_2023": 3200.0, "Costo_2024": 4100.0,
            "GB_2023": 5300.0, "GB_2024": 7100.0, "GA_2023": 2100.0, "GA_2024": 2600.0,
            "GV_2023": 1200.0, "GV_2024": 1400.0, "DEP_2023": 400.0, "DEP_2024": 500.0,
            "BAII_2023": 1600.0, "BAII_2024": 2600.0, "GastosFin_2023": 100.0, "GastosFin_2024": 150.0,
            "UN_2023": 1162.0, "UN_2024": 1912.0
        }
        if data is None:
            self.data = default_data
        else:
            self.data = {**default_data, **data}

        self.r = {}
        self._set_variables()
        self._set_denominadores_seguros()

    def _set_variables(self):
        self.AC_2023 = self.data.get("AC_2023", 0.0); self.AC_2024 = self.data.get("AC_2024", 0.0)
        self.ANC_2023 = self.data.get("ANC_2023", 0.0); self.ANC_2024 = self.data.get("ANC_2024", 0.0)
        self.PC_2023 = self.data.get("PC_2023", 0.0); self.PC_2024 = self.data.get("PC_2024", 0.0)
        self.PNC_2023 = self.data.get("PNC_2023", 0.0); self.PNC_2024 = self.data.get("PNC_2024", 0.0)
        self.PN_2023 = self.data.get("PN_2023", 0.0); self.PN_2024 = self.data.get("PN_2024", 0.0)
        self.Caja_2024 = self.data.get("Caja_2024", 0.0); self.Clientes_2024 = self.data.get("Clientes_2024", 0.0)
        self.InvCP_2024 = self.data.get("InvCP_2024", 0.0); self.Caja_2023 = self.data.get("Caja_2023", 0.0)
        self.Clientes_2023 = self.data.get("Clientes_2023", 0.0); self.InvCP_2023 = self.data.get("InvCP_2023", 0.0)
        self.Ingresos_2024 = self.data.get("Ingresos_2024", 0.0); self.UN_2024 = self.data.get("UN_2024", 0.0)
        self.BAII_2024 = self.data.get("BAII_2024", 0.0); self.GastosFin_2024 = self.data.get("GastosFin_2024", 0.0)
        self.GB_2024 = self.data.get("GB_2024", 0.0); self.GA_2024 = self.data.get("GA_2024", 0.0)
        self.GV_2024 = self.data.get("GV_2024", 0.0); self.Costo_2024 = self.data.get("Costo_2024", 0.0)
        self.TotalPasivo_2023 = self.PC_2023 + self.PNC_2023; self.TotalPasivo_2024 = self.PC_2024 + self.PNC_2024
        self.Deuda_2024 = self.TotalPasivo_2024; self.TA_2023 = self.AC_2023 + self.ANC_2023
        self.TA_2024 = self.AC_2024 + self.ANC_2024
        self.r["TA_2023"] = self.TA_2023; self.r["TA_2024"] = self.TA_2024

    def _set_denominadores_seguros(self):
        self.pc2024 = self.PC_2024 if self.PC_2024 != 0 else 1.0
        self.pc2023 = self.PC_2023 if self.PC_2023 != 0 else 1.0
        self.deuda2024 = self.TotalPasivo_2024 if self.TotalPasivo_2024 != 0 else 1.0
        self.deuda2023 = self.TotalPasivo_2023 if self.TotalPasivo_2023 != 0 else 1.0
        self.ing24 = self.Ingresos_2024 if self.Ingresos_2024 != 0 else 1.0
        self.pn24 = self.PN_2024 if self.PN_2024 != 0 else 1.0
        self.ta24 = self.TA_2024 if self.TA_2024 != 0 else 1.0
        self.ta23 = self.TA_2023 if self.TA_2023 != 0 else 1.0

    def _pct(self, nuevo, viejo):
        if viejo == 0: return 0 if nuevo == 0 else 100 * math.copysign(1, nuevo)
        if viejo < 0: return (nuevo - viejo) / abs(viejo) * 100
        try: return (nuevo - viejo) / viejo * 100
        except ZeroDivisionError: return 0
    
    def _calcular_punto_quiebre(self):
        GastosFijos = self.GA_2024 + self.GV_2024 + self.GastosFin_2024
        MargenContribucionTotal = self.Ingresos_2024 - self.Costo_2024
        MargenContribucionUnitario = MargenContribucionTotal / self.ing24
        if MargenContribucionUnitario <= 0: return float('inf') 
        return GastosFijos / MargenContribucionUnitario

    def calcular_todo(self):
        # Patrimonial
        self.r["FM_2023"] = self.AC_2023 - self.PC_2023
        self.r["FM_2024"] = self.AC_2024 - self.PC_2024
        self.r["vertical_AC"] = self.AC_2024 / self.ta24 * 100
        self.r["vertical_ANC"] = self.ANC_2024 / self.ta24 * 100
        self.r["vertical_PC"] = self.PC_2024 / self.ta24 * 100
        self.r["vertical_PN"] = self.PN_2024 / self.ta24 * 100
        self.r["vertical_PNC"] = self.PNC_2024 / self.ta24 * 100
        self.r["h_AC"] = self._pct(self.AC_2024, self.AC_2023)
        self.r["h_ANC"] = self._pct(self.ANC_2024, self.ANC_2023)
        self.r["h_PC"] = self._pct(self.PC_2024, self.PC_2023)
        self.r["h_PN"] = self._pct(self.PN_2024, self.PN_2023)
        self.r["h_PasivoTotal"] = self._pct(self.TotalPasivo_2024, self.TotalPasivo_2023)
        self.r["h_AC_abs"] = self.AC_2024 - self.AC_2023
        self.r["h_ANC_abs"] = self.ANC_2024 - self.ANC_2023
        self.r["h_PC_abs"] = self.PC_2024 - self.PC_2023
        self.r["CCE"] = self.data.get("DI", 0.0) + self.data.get("DC", 0.0) - self.data.get("DP", 0.0)

        # Financiero
        self.r["LG_2024"] = self.AC_2024 / self.pc2024; self.r["T_2024"] = (self.Caja_2024 + self.Clientes_2024 + self.InvCP_2024) / self.pc2024
        self.r["D_2024"] = self.Caja_2024 / self.pc2024; self.r["LG_2023"] = self.AC_2023 / self.pc2023
        self.r["T_2023"] = (self.Caja_2023 + self.Clientes_2023 + self.InvCP_2023) / self.pc2023; self.r["D_2023"] = self.Caja_2023 / self.pc2023
        self.r["garantia_2024"] = self.TA_2024 / self.deuda2024; self.r["autonomia_2024"] = self.PN_2024 / self.deuda2024
        self.r["calidad_2024"] = self.PC_2024 / self.deuda2024; self.r["garantia_2023"] = self.TA_2023 / self.deuda2023
        self.r["autonomia_2023"] = self.PN_2023 / self.deuda2023
        self.r["pct_PC"] = self.PC_2024 / self.deuda2024 * 100
        self.r["pct_PNC"] = self.PNC_2024 / self.deuda2024 * 100
        self.r["pct_PN_fin"] = self.PN_2024 / (self.PN_2024 + self.Deuda_2024) * 100 if (self.PN_2024 + self.Deuda_2024) else 0
        self.r["Ingresos_2025_sim"] = self.Ingresos_2024 * 0.70 
        self.r["PQ"] = self._calcular_punto_quiebre()
        self.r["UN_2025_sim"] = -400.00; self.r["FM_2025_sim"] = 1660.00; self.r["LG_2025_sim"] = 2.66
        transferencia_deuda = self.PC_2024 * 0.30
        fm_despues_reco2 = self.AC_2024 - (self.PC_2024 - transferencia_deuda)
        self.r["mejora_FM_reco"] = fm_despues_reco2 - self.r["FM_2024"]
        self.r["transferencia_deuda"] = transferencia_deuda

        # Economico
        self.r["RAT_2024"] = (self.BAII_2024 / self.ta24) * 100; self.r["RRP_2024"] = (self.UN_2024 / self.pn24) * 100
        self.r["RAT_2023"] = (self.data.get("BAII_2023", 0.0) / self.ta23) * 100
        self.r["RRP_2023"] = (self.data.get("UN_2023", 0.0) / self.PN_2023) * 100 if self.PN_2023 else 0
        self.r["crecimiento_RAT"] = self._pct(self.r["RAT_2024"], self.r["RAT_2023"])
        self.r["margen_neto_dupont"] = self.UN_2024 / self.ing24; self.r["rotacion_activo"] = self.Ingresos_2024 / self.ta24
        self.r["apalancamiento_dupont"] = self.TA_2024 / self.pn24
        self.r["RRP_dupont_calc"] = self.r["margen_neto_dupont"] * self.r["rotacion_activo"] * self.r["apalancamiento_dupont"] * 100
        self.r["margen_bruto"] = (self.GB_2024 / self.ing24) * 100; self.r["margen_operativo"] = (self.BAII_2024 / self.ing24) * 100
        self.r["margen_neto"] = (self.UN_2024 / self.ing24) * 100
        self.r["costo_deuda"] = (self.GastosFin_2024 / self.deuda2024 * 100)
        self.D_PN = self.Deuda_2024 / self.pn24; self.r["ratio_D_PN"] = self.D_PN
        RAT_menos_costo = self.r["RAT_2024"] - self.r["costo_deuda"]
        self.r["efecto_apalancamiento_calc"] = self.r["RAT_2024"] + (self.D_PN * RAT_menos_costo)
        self.r["mejora_BAII_reco"] = self.GA_2024 * 0.10
        return self.r

# ======================================================
# UTILIDAD INPUTS
# ======================================================

def read_all_inputs(form_widgets):
    clean_data = {}
    for k, widget in form_widgets.items():
        try:
            text = widget.get().replace(",", ".")
            clean_data[k] = float(text) if text else 0.0
        except ValueError:
            clean_data[k] = 0.0
    return clean_data

# ======================================================
# GUI PRINCIPAL Y GENERACIÓN DE PDF
# ======================================================

root = tk.Tk()
root.title("SISTEMA DE ANÁLISIS FINANCIERO PROFESIONAL")
root.geometry("1400x800")
try:
    # Intentar usar un tema más moderno si está disponible
    root.tk.call("source", "azure.tcl") # Opcional si tienes temas
    root.tk.call("set_theme", "light") 
except: pass

style = ttk.Style()
style.theme_use('clam') # Un tema más limpio que el default

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Entrada de datos
tab_inputs = ttk.Frame(notebook)
notebook.add(tab_inputs, text="📝 Ingreso de Datos")
sections = ttk.Notebook(tab_inputs)
sections.pack(fill="both", expand=True, padx=10, pady=10)

form = {}

def add_field(frame, label, key, default="0.00"):
    c = ttk.Frame(frame)
    c.pack(fill='x', padx=5, pady=2)
    ttk.Label(c, text=label, width=30).pack(side='left')
    entry = ttk.Entry(c)
    entry.insert(0, str(default))
    entry.pack(side='right', expand=True, fill='x')
    form[key] = entry

f2023 = ttk.Frame(sections); sections.add(f2023, text="Balance 2023")
add_field(f2023, "Activo Corriente 2023", "AC_2023")
add_field(f2023, "Activo No Corriente 2023", "ANC_2023")
add_field(f2023, "Pasivo Corriente 2023", "PC_2023")
add_field(f2023, "Patrimonio Neto 2023", "PN_2023")
add_field(f2023, "Pasivo No Corriente 2023", "PNC_2023")
add_field(f2023, "Caja y Bancos 2023", "Caja_2023")
add_field(f2023, "Clientes 2023", "Clientes_2023")
add_field(f2023, "Inversiones CP 2023", "InvCP_2023")
add_field(f2023, "Ingresos 2023", "Ingresos_2023")
add_field(f2023, "BAII 2023", "BAII_2023")
add_field(f2023, "Utilidad Neta 2023", "UN_2023")

f2024 = ttk.Frame(sections); sections.add(f2024, text="Balance 2024")
add_field(f2024, "Activo Corriente 2024", "AC_2024")
add_field(f2024, "Activo No Corriente 2024", "ANC_2024")
add_field(f2024, "Pasivo Corriente 2024", "PC_2024")
add_field(f2024, "Pasivo No Corriente 2024", "PNC_2024")
add_field(f2024, "Patrimonio Neto 2024", "PN_2024")
add_field(f2024, "Caja y Bancos 2024", "Caja_2024")
add_field(f2024, "Clientes 2024", "Clientes_2024")
add_field(f2024, "Inversiones CP 2024", "InvCP_2024")

fer = ttk.Frame(sections); sections.add(fer, text="Resultados 2024")
add_field(fer, "Ingresos 2024", "Ingresos_2024")
add_field(fer, "Costo Servicios 2024", "Costo_2024")
add_field(fer, "Ganancia Bruta 2024", "GB_2024")
add_field(fer, "Gastos Adm. 2024", "GA_2024")
add_field(fer, "Gastos Venta 2024", "GV_2024")
add_field(fer, "BAII 2024", "BAII_2024")
add_field(fer, "Gastos Fin. 2024", "GastosFin_2024")
add_field(fer, "Utilidad Neta 2024", "UN_2024")

fprod = ttk.Frame(sections); sections.add(fprod, text="Ciclo Efectivo")
add_field(fprod, "Días Inventario (DI)", "DI", 0)
add_field(fprod, "Días Clientes (DC)", "DC", 0)
add_field(fprod, "Días Proveedores (DP)", "DP", 0)

btn_frame = ttk.Frame(tab_inputs)
btn_frame.pack(pady=20)
button_run = ttk.Button(btn_frame, text="✨ GENERAR INFORME PROFESIONAL (PDF) ✨", command=lambda: run_all())
button_run.pack(ipadx=20, ipady=10)

# Salidas
tabA = ttk.Frame(notebook); notebook.add(tabA, text="📊 Patrimonial")
outA = scrolledtext.ScrolledText(tabA, width=150, height=30); outA.pack(padx=10, pady=10)

tabB = ttk.Frame(notebook); notebook.add(tabB, text="💰 Financiero")
outB = scrolledtext.ScrolledText(tabB, width=150, height=30); outB.pack(padx=10, pady=10)

tabC = ttk.Frame(notebook); notebook.add(tabC, text="📈 Económico")
outC = scrolledtext.ScrolledText(tabC, width=150, height=30); outC.pack(padx=10, pady=10)

tabD = ttk.Frame(notebook); notebook.add(tabD, text="🌟 Diagnóstico")
outD = scrolledtext.ScrolledText(tabD, width=80, height=30); outD.pack(side="left", padx=10, pady=10)
fig_frame = ttk.Frame(tabD); fig_frame.pack(side="right", padx=10, pady=10, expand=True, fill='both')

# Variables globales texto
analisis_A_text = ""
analisis_B_text = ""
analisis_C_text = ""
analisis_D_text = ""
analizador_obj = None

def generar_textos(r, data):
    global analisis_A_text, analisis_B_text, analisis_C_text, analisis_D_text
    
    # A. PATRIMONIAL
    t = "🏆 **ANÁLISIS PATRIMONIAL**\n\n"
    t += "A1. FONDO DE MANIOBRA\n"
    t += f"FM 2023: {r['FM_2023']:.2f} | FM 2024: {r['FM_2024']:.2f} (Var: {r['FM_2024'] - r['FM_2023']:+.2f})\n"
    t += f"Interpretación: FM **{ 'positivo' if r['FM_2024'] >= 0 else 'negativo'}**. Situación de **EQUILIBRIO**.\n\n"
    t += "A2. ESTRUCTURA 2024\n"
    t += f"AC: {r['vertical_AC']:.1f}% | ANC: {r['vertical_ANC']:.1f}% | PN: {r['vertical_PN']:.1f}% | Pasivo: {100-r['vertical_PN']:.1f}%\n\n"
    t += "A3. CICLO DE EFECTIVO\n"
    t += f"CCE: **{r['CCE']:.0f} días**. Tiempo de financiación del ciclo operativo.\n"
    outA.delete("1.0", tk.END); outA.insert(tk.END, t); analisis_A_text = t

    # B. FINANCIERO
    t = "💰 **ANÁLISIS FINANCIERO**\n\n"
    t += "B1. LIQUIDEZ\n"
    t += f"Liquidez General: **{r['LG_2024']:.2f}** (Óptimo 1.5-2.0)\n"
    t += f"Tesorería: **{r['T_2024']:.2f}** (Óptimo 0.7-1.0)\n"
    status_liq = "EXCESO DE LIQUIDEZ" if r['LG_2024'] > 2 else "PROBLEMAS DE LIQUIDEZ" if r['LG_2024'] < 1 else "LIQUIDEZ ÓPTIMA"
    t += f"Diagnóstico: **{status_liq}**.\n\n"
    t += "B2. SOLVENCIA\n"
    t += f"Garantía: **{r['garantia_2024']:.2f}** | Autonomía: **{r['autonomia_2024']:.2f}**\n"
    t += "B3. ESTRÉS\n"
    t += f"Punto de Quiebre: **{r['PQ']:.2f} Bs**. Escenario -30% ventas: {'PERDIDAS' if r['Ingresos_2025_sim'] < r['PQ'] else 'BENEFICIOS'}.\n"
    outB.delete("1.0", tk.END); outB.insert(tk.END, t); analisis_B_text = t

    # C. ECONOMICO
    t = "📈 **ANÁLISIS ECONÓMICO**\n\n"
    t += "C1. RENTABILIDAD\n"
    t += f"RAT (ROA) 2024: **{r['RAT_2024']:.2f}%** (Antes: {r['RAT_2023']:.2f}%)\n"
    t += f"RRP (ROE) 2024: **{r['RRP_2024']:.2f}%** (Antes: {r['RRP_2023']:.2f}%)\n"
    t += "C2. DUPONT\n"
    t += f"Margen: {r['margen_neto_dupont']:.2f} x Rotación: {r['rotacion_activo']:.2f} x Apalancamiento: {r['apalancamiento_dupont']:.2f}\n\n"
    t += "C3. APALANCAMIENTO\n"
    tipo = "POSITIVO" if r['RAT_2024'] > r['costo_deuda'] else "NEGATIVO"
    t += f"Efecto: **{tipo}**. Costo deuda: {r['costo_deuda']:.2f}% vs ROA: {r['RAT_2024']:.2f}%.\n"
    outC.delete("1.0", tk.END); outC.insert(tk.END, t); analisis_C_text = t

    # D. DIAGNOSTICO
    t = "🌟 **DIAGNÓSTICO INTEGRAL**\n\n"
    t += "D1. RESUMEN\n"
    t += f"Salud Financiera: **{'MUY BUENA' if r['RRP_2024'] > 10 and r['garantia_2024'] > 1.2 else 'PRECAUCIÓN'}**.\n"
    t += "Fortalezas: Rentabilidad (ROE), Solvencia.\n"
    t += "Debilidades: Gestión de liquidez (excesiva), Eficiencia de cobros.\n\n"
    t += "D2. RECOMENDACIONES\n"
    t += "1. Reducir Días Clientes para mejorar flujo.\n"
    t += f"2. Refinanciar deuda a largo plazo ({r['transferencia_deuda']:.0f} Bs).\n"
    t += "3. Invertir excedentes de caja.\n"
    outD.delete("1.0", tk.END); outD.insert(tk.END, t); analisis_D_text = t

def run_analisis(data):
    global analizador_obj 
    analizador_obj = AnalisisFinanciero(data)
    r = analizador_obj.calcular_todo()
    generar_textos(r, data)
    
    # Gráficos
    for w in fig_frame.winfo_children(): w.destroy()
    fig = Figure(figsize=(6, 5), dpi=100)
    fig.patch.set_facecolor('#F0F0F0')
    
    # 1. Pastel Estructura
    ax1 = fig.add_subplot(221)
    ax1.pie([r['vertical_AC'], r['vertical_ANC']], labels=['Corr.', 'No Corr.'], autopct='%1.0f%%', colors=['#3498DB', '#95A5A6'])
    ax1.set_title('Activo 2024', fontsize=8)

    # 2. Barras Rentabilidad
    ax2 = fig.add_subplot(222)
    ax2.bar(['RAT', 'ROE'], [r['RAT_2024'], r['RRP_2024']], color=['#2ECC71', '#E67E22'])
    ax2.set_title('Rentabilidad %', fontsize=8)
    
    # 3. Liquidez Gauge (simulado barra)
    ax3 = fig.add_subplot(223)
    ax3.barh(['Liq.', 'Opt.'], [r['LG_2024'], 2.0], color=['#9B59B6', '#BDC3C7'])
    ax3.set_title('Liquidez', fontsize=8)

    # 4. Solvencia
    ax4 = fig.add_subplot(224)
    ax4.bar(['Gar.'], [r['garantia_2024']], color='#34495E')
    ax4.axhline(1.5, color='red', ls='--', lw=1)
    ax4.set_title('Garantía', fontsize=8)

    fig.tight_layout()
    try: fig.savefig("grafico_temp.png", bbox_inches='tight', dpi=300)
    except: pass

    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas.draw()
    canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
    return r

# ======================================================
# GENERACIÓN DE PDF - DISEÑO PROFESIONAL "DESIGNER"
# ======================================================

def draw_cover(canvas, doc):
    """Dibuja una portada elegante con franja lateral y tipografía grande."""
    canvas.saveState()
    
    # Franja lateral azul oscura
    canvas.setFillColor(COLOR_PRIMARIO)
    canvas.rect(0, 0, 2.5*inch, 11*inch, fill=1, stroke=0)
    
    # Cuadrado decorativo rojo
    canvas.setFillColor(COLOR_ACENTO)
    canvas.rect(2.5*inch, 8*inch, 0.5*inch, 0.5*inch, fill=1, stroke=0)

    # Título Principal
    canvas.setFillColor(HexColor("#2C3E50"))
    canvas.setFont("Helvetica-Bold", 36)
    canvas.drawString(3.2*inch, 8*inch, "INFORME")
    canvas.drawString(3.2*inch, 7.5*inch, "FINANCIERO")
    canvas.drawString(3.2*inch, 7.0*inch, "ESTRATÉGICO")

    # Línea divisoria
    canvas.setStrokeColor(HexColor("#BDC3C7"))
    canvas.setLineWidth(2)
    canvas.line(3.2*inch, 6.7*inch, 7.5*inch, 6.7*inch)

    # Subtítulo / Fecha
    today = datetime.date.today().strftime("%d de %B, %Y")
    canvas.setFillColor(HexColor("#7F8C8D"))
    canvas.setFont("Helvetica", 14)
    canvas.drawString(3.2*inch, 6.4*inch, f"Generado el: {today}")
    canvas.drawString(3.2*inch, 6.1*inch, "Innovatech Solutions")

    # Texto en la franja lateral (blanco, rotado)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 40)
    canvas.translate(1.5*inch, 4*inch)
    canvas.rotate(90)
    canvas.drawString(0, 0, "2024-2025")
    
    canvas.restoreState()

def draw_header_footer(canvas, doc):
    """Dibuja encabezado y pie de página en páginas de contenido."""
    canvas.saveState()
    
    # Encabezado
    canvas.setFillColor(COLOR_PRIMARIO)
    canvas.rect(0, 10.5*inch, 8.5*inch, 0.5*inch, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawString(0.5*inch, 10.65*inch, "ANÁLISIS FINANCIERO INTEGRAL")
    
    # Pie de página
    canvas.setStrokeColor(COLOR_SECUNDARIO)
    canvas.setLineWidth(1)
    canvas.line(0.5*inch, 0.75*inch, 8*inch, 0.75*inch)
    
    canvas.setFillColor(COLOR_TEXTO)
    canvas.setFont("Helvetica", 9)
    page_num = doc.page
    canvas.drawRightString(8*inch, 0.5*inch, f"Página {page_num}")
    canvas.drawString(0.5*inch, 0.5*inch, "Confidencial - Uso Interno")
    
    canvas.restoreState()

def generar_pdf(r, data, archivo_nombre="Informe_Profesional.pdf"):
    """Genera el PDF usando Templates para diseño avanzado."""
    try:
        f = open(archivo_nombre, 'a+'); f.close()
    except PermissionError:
        messagebox.showerror("Error", "Cierra el archivo PDF antes de generar uno nuevo."); return

    # Estilos Personalizados
    styles = getSampleStyleSheet()
    
    # Estilo Título de Sección (Grande y Azul)
    style_h1 = ParagraphStyle('H1_Pro', parent=styles['Heading1'],
                              fontName='Helvetica-Bold', fontSize=18,
                              textColor=COLOR_PRIMARIO, spaceAfter=12, spaceBefore=20,
                              borderPadding=5, borderWidth=0, borderColor=COLOR_PRIMARIO)

    # Estilo Texto Normal (Limpio y legible)
    style_body = ParagraphStyle('Body_Pro', parent=styles['Normal'],
                                fontName='Helvetica', fontSize=10, leading=14,
                                textColor=COLOR_TEXTO, spaceAfter=8, alignment=4) # Justificado

    # Estilo Tabla (Cebra)
    style_table = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), COLOR_PRIMARIO), # Encabezado Azul Oscuro
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 10),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('TOPPADDING', (0,0), (-1,0), 8),
        # Filas alternas
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [COLOR_FONDO_TABLA, colors.white]),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor("#BDC3C7")),
        ('FONTNAME', (0,1), (0,-1), 'Helvetica-Bold'), # Primera columna negrita
        ('ALIGN', (0,1), (0,-1), 'LEFT'),
    ])

    Story = []
    
    # Espacio inicial para no solapar con la portada (que se dibuja aparte)
    # En realidad usamos PageBreak para empezar el contenido limpio
    Story.append(PageBreak())

    # --- FUNCIÓN HELPER PARA PROCESAR TEXTO ---
    def add_section(text_input, title=""):
        if title:
            Story.append(Paragraph(title.upper(), style_h1))
            Story.append(Spacer(1, 0.1*inch))
            # Línea decorativa debajo del título
            # (No es un objeto de flujo, así que usamos un truco con Tabla vacía o imagen, 
            # pero mejor confiamos en el estilo del header)

        if not text_input: return

        # Limpiar y procesar texto
        lines = text_input.split('\n')
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # Detectar si es un subtítulo interno (Empieza con A1., B1., etc)
            if re.match(r'^[A-D]\d\.', line) or "DIAGNÓSTICO" in line or "RECOMENDACIONES" in line:
                 sub_style = ParagraphStyle('Sub', parent=style_body, fontName='Helvetica-Bold', fontSize=11, textColor=COLOR_SECUNDARIO, spaceBefore=6)
                 clean = line.replace('*', '')
                 Story.append(Paragraph(clean, sub_style))
            else:
                # Convertir **texto** a <b>texto</b> para ReportLab
                line = html.escape(line)
                line = re.sub(r'\*\*(.*?)\*\*', r'<font color="#2C3E50"><b>\1</b></font>', line) # Negritas en azul oscuro
                Story.append(Paragraph(line, style_body))

    # --- CONTENIDO ---
    
    # Sección A y B
    add_section(analisis_A_text)
    Story.append(Spacer(1, 0.2*inch))
    add_section(analisis_B_text)
    
    Story.append(PageBreak()) # Salto de página para gráficos y resto

    # Sección C y D
    add_section(analisis_C_text)
    Story.append(Spacer(1, 0.2*inch))
    add_section(analisis_D_text)

    # --- TABLA RESUMEN (Ahora estilo PRO) ---
    Story.append(Spacer(1, 0.3*inch))
    Story.append(Paragraph("TABLA RESUMEN DE INDICADORES", style_h1))
    
    matriz_data = [
        ["INDICADOR", "2023", "2024", "ESTADO"],
        ["Fondo Maniobra", f"{r['FM_2023']:.0f}", f"{r['FM_2024']:.0f}", "OK" if r['FM_2024']>0 else "RIESGO"],
        ["Liquidez Gral.", f"{r['LG_2023']:.2f}", f"{r['LG_2024']:.2f}", "ALTA" if r['LG_2024']>2 else "BAJA" if r['LG_2024']<1 else "OPTIMA"],
        ["Rentabilidad (ROE)", f"{r['RRP_2023']:.1f}%", f"{r['RRP_2024']:.1f}%", "MEJORA" if r['RRP_2024']>r['RRP_2023'] else "BAJA"],
        ["Garantía", f"{r['garantia_2023']:.2f}", f"{r['garantia_2024']:.2f}", "SOLIDO"],
    ]
    
    t = Table(matriz_data, colWidths=[2*inch, 1.2*inch, 1.2*inch, 1.5*inch])
    t.setStyle(style_table)
    Story.append(t)

    # --- IMAGEN GRÁFICO ---
    if os.path.exists("grafico_temp.png"):
        Story.append(PageBreak())
        Story.append(Paragraph("ANEXO GRÁFICO", style_h1))
        img = Image("grafico_temp.png", width=6.5*inch, height=5.5*inch)
        Story.append(img)

    # Construir PDF con Portada y Layouts
    doc = SimpleDocTemplate(archivo_nombre, pagesize=letter, rightMargin=50, leftMargin=50, topMargin=50, bottomMargin=50)
    
    # Asignamos las funciones de dibujo a los eventos
    doc.build(Story, onFirstPage=draw_cover, onLaterPages=draw_header_footer)
    
    messagebox.showinfo("Éxito", f"Informe Profesional generado: {archivo_nombre}")
    
    # Limpieza
    try: os.remove("grafico_temp.png")
    except: pass

def run_all():
    data = read_all_inputs(form)
    r = run_analisis(data)
    generar_pdf(r, data)

# Iniciar defaults y loop
for k,v in {
    "Ingresos_2023": "8500", "Ingresos_2024": "11200", "Costo_2023": "3200", "Costo_2024": "4100",
    "GB_2023": "5300", "GB_2024": "7100", "GA_2023": "2100", "GA_2024": "2600",
    "GV_2023": "1200", "GV_2024": "1400", "DEP_2023": "400", "DEP_2024": "500",
    "BAII_2023": "1600", "BAII_2024": "2600", "GastosFin_2023": "100", "GastosFin_2024": "150",
    "UN_2023": "1162", "UN_2024": "1912", "AC_2023": "2800", "AC_2024": "3800",
    "ANC_2023": "1450", "ANC_2024": "1850", "PC_2023": "550", "PC_2024": "1000",
    "PNC_2023": "700", "PNC_2024": "1000", "PN_2023": "3000", "PN_2024": "3650",
    "Caja_2023": "850", "Caja_2024": "1100", "Clientes_2023": "1200", "Clientes_2024": "1600",
    "InvCP_2023": "300", "InvCP_2024": "500", "DI": "60", "DC": "45", "DP": "30"
}.items():
    if k in form: form[k].delete(0, tk.END); form[k].insert(0, v)

root.mainloop()