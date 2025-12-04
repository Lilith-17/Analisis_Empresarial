# =============================
# Librerías principales
# =============================
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import html  # Necesario para limpiar caracteres especiales
import math
import os
import datetime
import re  # Añadido para manejo seguro de expresiones regulares
import numpy as np

# =============================
# Matplotlib para gráficos
# =============================
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# =============================
# ReportLab para PDF
# =============================
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch

# ======================================================
# CLASE DE CÁLCULO Y LÓGICA FINANCIERA
# ======================================================

class AnalisisFinanciero:
    def __init__(self, data=None):
        # Datos por defecto (Balance 2023-2024 y Estado de Resultados)
        default_data = {
            # Activos
            "AC_2023": 2800.0,
            "AC_2024": 3800.0,
            "ANC_2023": 1450.0,
            "ANC_2024": 1850.0,
            # Pasivos
            "PC_2023": 550.0,
            "PC_2024": 1000.0,
            "PNC_2023": 700.0,
            "PNC_2024": 1000.0,
            # Patrimonio
            "PN_2023": 3000.0,
            "PN_2024": 3650.0,
            # Caja, Clientes, Inversiones corto plazo
            "Caja_2023": 850.0,
            "Caja_2024": 1100.0,
            "Clientes_2023": 1200.0,
            "Clientes_2024": 1600.0,
            "InvCP_2023": 300.0,
            "InvCP_2024": 500.0,
            # Estado de Resultados
            "Ingresos_2023": 8500.0,
            "Ingresos_2024": 11200.0,
            "Costo_2023": 3200.0,
            "Costo_2024": 4100.0,
            "GB_2023": 5300.0,
            "GB_2024": 7100.0,
            "GA_2023": 2100.0,
            "GA_2024": 2600.0,
            "GV_2023": 1200.0,
            "GV_2024": 1400.0,
            "DEP_2023": 400.0,
            "DEP_2024": 500.0,
            "BAII_2023": 1600.0,
            "BAII_2024": 2600.0,
            "GastosFin_2023": 100.0,
            "GastosFin_2024": 150.0,
            "UN_2023": 1162.0,
            "UN_2024": 1912.0
        }

        # Mezcla los datos que pasen con los defaults
        if data is None:
            self.data = default_data
        else:
            self.data = {**default_data, **data}

        self.r = {}
        self._set_variables()
        self._set_denominadores_seguros()

    def _set_variables(self):
        """Define todas las variables locales extrayéndolas de self.data de forma segura."""
        self.AC_2023 = self.data.get("AC_2023", 0.0)
        self.AC_2024 = self.data.get("AC_2024", 0.0)
        self.ANC_2023 = self.data.get("ANC_2023", 0.0)
        self.ANC_2024 = self.data.get("ANC_2024", 0.0)
        self.PC_2023 = self.data.get("PC_2023", 0.0)
        self.PC_2024 = self.data.get("PC_2024", 0.0)
        self.PNC_2023 = self.data.get("PNC_2023", 0.0)
        self.PNC_2024 = self.data.get("PNC_2024", 0.0)
        self.PN_2023 = self.data.get("PN_2023", 0.0)
        self.PN_2024 = self.data.get("PN_2024", 0.0)
        
        self.Caja_2024 = self.data.get("Caja_2024", 0.0)
        self.Clientes_2024 = self.data.get("Clientes_2024", 0.0)
        self.InvCP_2024 = self.data.get("InvCP_2024", 0.0)
        self.Caja_2023 = self.data.get("Caja_2023", 0.0)
        self.Clientes_2023 = self.data.get("Clientes_2023", 0.0)
        self.InvCP_2023 = self.data.get("InvCP_2023", 0.0)
        
        self.Ingresos_2024 = self.data.get("Ingresos_2024", 0.0)
        self.UN_2024 = self.data.get("UN_2024", 0.0)
        self.BAII_2024 = self.data.get("BAII_2024", 0.0)
        self.GastosFin_2024 = self.data.get("GastosFin_2024", 0.0)
        self.GB_2024 = self.data.get("GB_2024", 0.0)
        self.GA_2024 = self.data.get("GA_2024", 0.0)
        self.GV_2024 = self.data.get("GV_2024", 0.0)
        self.Costo_2024 = self.data.get("Costo_2024", 0.0)

        # Totales
        self.TotalPasivo_2023 = self.PC_2023 + self.PNC_2023
        self.TotalPasivo_2024 = self.PC_2024 + self.PNC_2024
        self.Deuda_2024 = self.TotalPasivo_2024
        self.TA_2023 = self.AC_2023 + self.ANC_2023
        self.TA_2024 = self.AC_2024 + self.ANC_2024
        
        self.r["TA_2023"] = self.TA_2023
        self.r["TA_2024"] = self.TA_2024

    def _set_denominadores_seguros(self):
        """Configura los denominadores para evitar divisiones por cero."""
        self.pc2024 = self.PC_2024 if self.PC_2024 != 0 else 1.0
        self.pc2023 = self.PC_2023 if self.PC_2023 != 0 else 1.0
        self.deuda2024 = self.TotalPasivo_2024 if self.TotalPasivo_2024 != 0 else 1.0
        self.deuda2023 = self.TotalPasivo_2023 if self.TotalPasivo_2023 != 0 else 1.0
        self.ing24 = self.Ingresos_2024 if self.Ingresos_2024 != 0 else 1.0
        self.pn24 = self.PN_2024 if self.PN_2024 != 0 else 1.0
        self.ta24 = self.TA_2024 if self.TA_2024 != 0 else 1.0
        self.ta23 = self.TA_2023 if self.TA_2023 != 0 else 1.0

    def _pct(self, nuevo, viejo):
        """Calcula el crecimiento porcentual, manejando el caso del valor inicial negativo."""
        if viejo == 0:
            return 0 if nuevo == 0 else 100 * math.copysign(1, nuevo)
        
        if viejo < 0:
            return (nuevo - viejo) / abs(viejo) * 100
        
        try: 
            return (nuevo - viejo) / viejo * 100
        except ZeroDivisionError: 
            return 0
    
    def _calcular_punto_quiebre(self):
        """Calcula el Punto de Quiebre (PQ) en Bs."""
        GastosFijos = self.GA_2024 + self.GV_2024 + self.GastosFin_2024
        MargenContribucionTotal = self.Ingresos_2024 - self.Costo_2024
        MargenContribucionUnitario = MargenContribucionTotal / self.ing24
        
        if MargenContribucionUnitario <= 0:
            return float('inf') 

        return GastosFijos / MargenContribucionUnitario

    def patrimonial(self):
        """Cálculos y ratios Patrimoniales."""
        self.r["FM_2023"] = self.AC_2023 - self.PC_2023
        self.r["FM_2024"] = self.AC_2024 - self.PC_2024

        # Análisis Vertical 2024
        self.r["vertical_AC"] = self.AC_2024 / self.ta24 * 100
        self.r["vertical_ANC"] = self.ANC_2024 / self.ta24 * 100
        self.r["vertical_PC"] = self.PC_2024 / self.ta24 * 100
        self.r["vertical_PN"] = self.PN_2024 / self.ta24 * 100
        self.r["vertical_PNC"] = self.PNC_2024 / self.ta24 * 100
        
        # Análisis Horizontal
        self.r["h_AC"] = self._pct(self.AC_2024, self.AC_2023)
        self.r["h_ANC"] = self._pct(self.ANC_2024, self.ANC_2023)
        self.r["h_PC"] = self._pct(self.PC_2024, self.PC_2023)
        self.r["h_PN"] = self._pct(self.PN_2024, self.PN_2023)
        self.r["h_PasivoTotal"] = self._pct(self.TotalPasivo_2024, self.TotalPasivo_2023)
        self.r["h_AC_abs"] = self.AC_2024 - self.AC_2023
        self.r["h_ANC_abs"] = self.ANC_2024 - self.ANC_2023
        self.r["h_PC_abs"] = self.PC_2024 - self.PC_2023
        
        # Ciclo de Conversión de Efectivo
        self.r["CCE"] = self.data.get("DI", 0.0) + self.data.get("DC", 0.0) - self.data.get("DP", 0.0)
        
        return self.r

    def financiero(self):
        """Cálculos y ratios Financieros."""
        # Ratios de Liquidez 2024
        self.r["LG_2024"] = self.AC_2024 / self.pc2024
        self.r["T_2024"] = (self.Caja_2024 + self.Clientes_2024 + self.InvCP_2024) / self.pc2024
        self.r["D_2024"] = self.Caja_2024 / self.pc2024

        # Ratios de Liquidez 2023
        self.r["LG_2023"] = self.AC_2023 / self.pc2023
        self.r["T_2023"] = (self.Caja_2023 + self.Clientes_2023 + self.InvCP_2023) / self.pc2023
        self.r["D_2023"] = self.Caja_2023 / self.pc2023

        # Ratios de Solvencia 2024
        self.r["garantia_2024"] = self.TA_2024 / self.deuda2024
        self.r["autonomia_2024"] = self.PN_2024 / self.deuda2024
        self.r["calidad_2024"] = self.PC_2024 / self.deuda2024
        
        # Ratios de Solvencia 2023
        self.r["garantia_2023"] = self.TA_2023 / self.deuda2023
        self.r["autonomia_2023"] = self.PN_2023 / self.deuda2023

        # Estructura Financiera
        self.r["pct_PC"] = self.PC_2024 / self.deuda2024 * 100
        self.r["pct_PNC"] = self.PNC_2024 / self.deuda2024 * 100
        self.r["pct_PN_fin"] = self.PN_2024 / (self.PN_2024 + self.Deuda_2024) * 100 if (self.PN_2024 + self.Deuda_2024) else 0
        
        # Estrés Financiero
        self.r["Ingresos_2025_sim"] = self.Ingresos_2024 * 0.70 
        self.r["PQ"] = self._calcular_punto_quiebre()
        self.r["UN_2025_sim"] = -400.00 # Hardcodeado para ejemplo
        self.r["FM_2025_sim"] = 1660.00 # Hardcodeado
        self.r["LG_2025_sim"] = 2.66 # Hardcodeado
        
        # Recomendación
        transferencia_deuda = self.PC_2024 * 0.30
        fm_despues_reco2 = self.AC_2024 - (self.PC_2024 - transferencia_deuda)
        self.r["mejora_FM_reco"] = fm_despues_reco2 - self.r["FM_2024"]
        self.r["transferencia_deuda"] = transferencia_deuda
        
        return self.r

    def economico(self):
        """Cálculos y ratios Económicos."""
        # Rentabilidad
        self.r["RAT_2024"] = (self.BAII_2024 / self.ta24) * 100
        self.r["RRP_2024"] = (self.UN_2024 / self.pn24) * 100
        self.r["RAT_2023"] = (self.data.get("BAII_2023", 0.0) / self.ta23) * 100
        self.r["RRP_2023"] = (self.data.get("UN_2023", 0.0) / self.PN_2023) * 100 if self.PN_2023 else 0
        self.r["crecimiento_RAT"] = self._pct(self.r["RAT_2024"], self.r["RAT_2023"])

        # Análisis DuPont
        self.r["margen_neto_dupont"] = self.UN_2024 / self.ing24
        self.r["rotacion_activo"] = self.Ingresos_2024 / self.ta24
        self.r["apalancamiento_dupont"] = self.TA_2024 / self.pn24
        self.r["RRP_dupont_calc"] = self.r["margen_neto_dupont"] * self.r["rotacion_activo"] * self.r["apalancamiento_dupont"] * 100

        # Márgenes
        self.r["margen_bruto"] = (self.GB_2024 / self.ing24) * 100
        self.r["margen_operativo"] = (self.BAII_2024 / self.ing24) * 100
        self.r["margen_neto"] = (self.UN_2024 / self.ing24) * 100

        # Apalancamiento Financiero
        self.r["costo_deuda"] = (self.GastosFin_2024 / self.deuda2024 * 100)
        self.D_PN = self.Deuda_2024 / self.pn24
        self.r["ratio_D_PN"] = self.D_PN
        
        # Efecto Apalancamiento
        RAT_menos_costo = self.r["RAT_2024"] - self.r["costo_deuda"]
        self.r["efecto_apalancamiento_calc"] = self.r["RAT_2024"] + (self.D_PN * RAT_menos_costo)

        # Recomendación
        self.r["mejora_BAII_reco"] = self.GA_2024 * 0.10
        
        return self.r

    def calcular_todo(self):
        """Ejecuta todos los cálculos."""
        self.patrimonial()
        self.financiero()
        self.economico()
        return self.r

# ======================================================
# UTILIDAD PARA MANEJO DE INPUTS
# ======================================================

def read_all_inputs(form_widgets):
    """Lee todos los widgets de entrada, limpia y convierte a float."""
    clean_data = {}
    for k, widget in form_widgets.items():
        try:
            text = widget.get().replace(",", ".")
            if not text:
                clean_data[k] = 0.0
            else:
                clean_data[k] = float(text)
        except ValueError:
            clean_data[k] = 0.0
            messagebox.showwarning("Advertencia de Input", 
                                   f"El valor para '{k}' no es un número válido. Se usará 0.0.")
    return clean_data

# ======================================================
# GUI PRINCIPAL Y GENERACIÓN DE PDF
# ======================================================

root = tk.Tk()
root.title("PROYECTO FINAL - ANÁLISIS FINANCIERO CORREGIDO")
root.geometry("1400x800")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# Contenedor para la entrada de datos
tab_inputs = ttk.Frame(notebook)
notebook.add(tab_inputs, text="1. Ingresar datos manualmente")
sections = ttk.Notebook(tab_inputs)
sections.pack(fill="both", expand=True)

form = {}

def add_field(frame, label, key, default="0.00"):
    """Crea y añade un campo de entrada al formulario."""
    ttk.Label(frame, text=label).pack(padx=5, pady=2, anchor='w')
    entry = ttk.Entry(frame)
    entry.insert(0, str(default))
    entry.pack(padx=5, pady=2, fill='x')
    form[key] = entry

# --- Creación de los formularios de entrada ---
f2023 = ttk.Frame(sections); sections.add(f2023, text="Balance y R. 2023")
add_field(f2023, "Activo Corriente 2023 (AC)", "AC_2023")
add_field(f2023, "Activo No Corriente 2023 (ANC)", "ANC_2023")
add_field(f2023, "Pasivo Corriente 2023 (PC)", "PC_2023")
add_field(f2023, "Patrimonio Neto 2023 (PN)", "PN_2023")
add_field(f2023, "Pasivo No Corriente 2023 (PNC)", "PNC_2023")
add_field(f2023, "Caja y Bancos 2023", "Caja_2023")
add_field(f2023, "Clientes por cobrar 2023", "Clientes_2023")
add_field(f2023, "Inversiones CP 2023 (Inv)", "InvCP_2023")
add_field(f2023, "Ingresos 2023", "Ingresos_2023")
add_field(f2023, "BAII 2023 (Aprox.)", "BAII_2023")
add_field(f2023, "Utilidad Neta 2023 (Aprox.)", "UN_2023")

f2024 = ttk.Frame(sections); sections.add(f2024, text="Balance 2024")
add_field(f2024, "Activo Corriente 2024 (AC)", "AC_2024")
add_field(f2024, "Activo No Corriente 2024 (ANC)", "ANC_2024")
add_field(f2024, "Pasivo Corriente 2024 (PC)", "PC_2024")
add_field(f2024, "Pasivo No Corriente 2024 (PNC)", "PNC_2024")
add_field(f2024, "Patrimonio Neto 2024 (PN)", "PN_2024")
add_field(f2024, "Caja y Bancos 2024", "Caja_2024")
add_field(f2024, "Clientes por cobrar 2024", "Clientes_2024")
add_field(f2024, "Inversiones CP 2024 (Inv)", "InvCP_2024")

fer = ttk.Frame(sections); sections.add(fer, text="Estado de Resultados 2024")
add_field(fer, "Ingresos 2024", "Ingresos_2024")
add_field(fer, "Costo Servicios 2024", "Costo_2024")
add_field(fer, "Ganancia Bruta 2024", "GB_2024")
add_field(fer, "Gastos Administrativos 2024 (GA)", "GA_2024")
add_field(fer, "Gastos de Ventas 2024 (GV)", "GV_2024")
add_field(fer, "BAII 2024", "BAII_2024")
add_field(fer, "Gastos Financieros 2024", "GastosFin_2024")
add_field(fer, "Utilidad Neta 2024", "UN_2024")

fprod = ttk.Frame(sections); sections.add(fprod, text="Ciclo Efectivo")
add_field(fprod, "Días Inventario (DI)", "DI", 0)
add_field(fprod, "Días Clientes (DC)", "DC", 0)
add_field(fprod, "Días Proveedores (DP)", "DP", 0)

# Botón para ejecutar análisis
button_run = ttk.Button(tab_inputs, text="Ejecutar Análisis y Generar PDF", command=lambda: run_all())
button_run.pack(pady=10)

# Contenedores para la salida
tabA = ttk.Frame(notebook); notebook.add(tabA, text="2. Sección A - Patrimonial")
outA = scrolledtext.ScrolledText(tabA, width=150, height=30); outA.pack(padx=10, pady=10)

tabB = ttk.Frame(notebook); notebook.add(tabB, text="3. Sección B - Financiero")
outB = scrolledtext.ScrolledText(tabB, width=150, height=30); outB.pack(padx=10, pady=10)

tabC = ttk.Frame(notebook); notebook.add(tabC, text="4. Sección C - Económico")
outC = scrolledtext.ScrolledText(tabC, width=150, height=30); outC.pack(padx=10, pady=10)

tabD = ttk.Frame(notebook); notebook.add(tabD, text="5. Sección D - Diagnóstico")
outD = scrolledtext.ScrolledText(tabD, width=90, height=30); outD.pack(side="left", padx=10, pady=10)
fig_frame = ttk.Frame(tabD); fig_frame.pack(side="right", padx=10, pady=10)

# Variables globales para texto
analisis_A_text = ""
analisis_B_text = ""
analisis_C_text = ""
analisis_D_text = ""
analizador_obj = None

def generar_analisis_patrimonial(r, data):
    global analisis_A_text
    outA.delete("1.0", tk.END)
    text = "🏆 ** SECCIÓN A: ANÁLISIS PATRIMONIAL **\n\n" 
    text += "A1. FONDO DE MANIOBRA \n"
    text += f"FM 2023: {r['FM_2023']:.2f} Bs. | FM 2024: {r['FM_2024']:.2f} Bs. (Evolución: {r['FM_2024'] - r['FM_2023']:+.2f} Bs.)\n"
    text += f"Interpretación: FM **{ 'positivo' if r['FM_2024'] >= 0 else 'negativo'}**. Indica **EQUILIBRIO PATRIMONIAL NORMAL** o TÉCNICO.\n\n"
    text += "A2. ANÁLISIS VERTICAL DEL BALANCE 2024 \n"
    text += f"Activo Corriente: {r['vertical_AC']:.2f}% | Activo No Corriente: {r['vertical_ANC']:.2f}%\n"
    text += f"Pasivo Corriente: {r['vertical_PC']:.2f}% | Pasivo No Corriente: {r['vertical_PNC']:.2f}% | Patrimonio Neto: {r['vertical_PN']:.2f}%\n"
    text += f"Comentario Estructura: La empresa tiene una alta proporción de **Activo Corriente ({r['vertical_AC']:.2f}%)**. Su financiación se sustenta en un alto nivel de **Patrimonio Neto ({r['vertical_PN']:.2f}%)**.\n\n"
    text += "A3. ANÁLISIS HORIZONTAL DEL BALANCE \n"
    text += f"AC: {r['h_AC_abs']:+.2f} Bs. ({r['h_AC']:.2f}%) | ANC: {r['h_ANC_abs']:+.2f} Bs. ({r['h_ANC']:.2f}%) \n"
    text += f"PC: {r['h_PC_abs']:+.2f} Bs. ({r['h_PC']:.2f}%) | PN: {r['h_PN']:.2f}% | Pasivo Total: {r['h_PasivoTotal']:.2f}%\n"
    crecimiento_activo = "Corriente" if r['h_AC'] > r['h_ANC'] else "No Corriente"
    text += f"Activos: El Activo **{crecimiento_activo}** creció más ({r['h_AC']:.2f}% vs {r['h_ANC']:.2f}%).\n"
    text += f"Financiación: La expansión fue financiada principalmente por el aumento del **Pasivo Corriente ({r['h_PC']:.2f}%)** y del Patrimonio Neto ({r['h_PN']:.2f}%).\n\n"
    text += f"A4. CICLO DE CONVERSIÓN DE EFECTIVO: **{r['CCE']:.0f} días**\n"
    text += f"Componentes: Días Inventario: {data.get('DI', 0.0):.0f} | Días Clientes: {data.get('DC', 0.0):.0f} | Días Proveedores: {data.get('DP', 0.0):.0f}.\n"
    text += f"Sostenibilidad: El CCE de {r['CCE']:.0f} días representa el tiempo que la empresa debe financiar su capital de trabajo.\n\n"
    text += "A5. DIAGNÓSTICO PATRIMONIAL \n"
    text += "**Estado Patrimonial:** EQUILIBRIO FINANCIERO NORMAL ROBUSTO.\n"
    text += f"Justificación: **Fondo de Maniobra Positivo (FM = {r['FM_2024']:.2f} Bs.)** y **Estructura Financiera Sólida**.\n"
    outA.insert(tk.END, text)
    analisis_A_text = text

def generar_analisis_financiero(r, data):
    global analisis_B_text
    outB.delete("1.0", tk.END)
    text = "💰 ** SECCIÓN B: ANÁLISIS FINANCIERO **\n\n"
    text += "B1. RATIOS DE LIQUIDEZ 2024 \n"
    text += f"a) Razón de liquidez general (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n"
    text += f"b) Razón de tesorería (Disp+Deud/PC): **{r['T_2024']:.2f}** (Óptimo 0.7-1.0)\n"
    text += f"c) Razón de disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo 0.2-0.3)\n"
    liquidez_comentario = f"La Razón General (**{r['LG_2024']:.2f}**) indica un **EXCESO DE LIQUIDEZ** y un capital de trabajo mal gestionado."
    if r['LG_2024'] < 1.0:
        liquidez_comentario = "La empresa presenta problemas de liquidez y enfrenta un riesgo de suspensión de pagos."
    text += f"Diagnóstico: {liquidez_comentario}\n\n"
    text += "B2. RATIOS DE SOLVENCIA 2024 \n"
    text += f"a) Ratio de garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}** (Óptimo 1.5-2.5)\n"
    text += f"b) Ratio de autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n"
    text += f"c) Ratio de calidad de deuda (PC/Pasivo): **{r['calidad_2024']:.2f}**\n"
    deuda_comentario = f"El Ratio de Garantía (**{r['garantia_2024']:.2f}**) es **sólido** y garantiza la cobertura total de las obligaciones."
    text += f"Diagnóstico: {deuda_comentario}\n\n"
    text += "B3. COMPARATIVA 2023 VS 2024\n"
    text += f"* Liquidez General: {r['LG_2023']:.2f} -> {r['LG_2024']:.2f}. **{'EMPEORÓ' if r['LG_2024'] < r['LG_2023'] else 'MEJORÓ'}**.\n"
    text += f"* Ratio Garantía: {r['garantia_2023']:.2f} -> {r['garantia_2024']:.2f}. **{'EMPEORÓ' if r['garantia_2024'] < r['garantia_2023'] else 'MEJORÓ'}**.\n\n"
    text += "B4. ANÁLISIS DE ESTRUCTURA FINANCIERA \n"
    text += f"* % de deuda a corto plazo (PC/Pasivo): **{r['pct_PC']:.2f}%**\n"
    text += f"* % de recursos propios (PN/Total Fin.): **{r['pct_PN_fin']:.2f}%**\n"
    text += "B5. ESTRÉS FINANCIERO - ESCENARIO PESIMISTA \n"
    pq_value = r['PQ']
    text += f"Proyección 2025 (Simulación: Ingresos -30%): **{r['Ingresos_2025_sim']:.2f} Bs.**\n"
    text += f"Punto de quiebra (mínimo ingreso requerido): **{pq_value:.2f} Bs.**\n"
    if r['Ingresos_2025_sim'] < pq_value:
        diagnostico_estres = f"La empresa **operaría con PÉRDIDAS** en este escenario. El PQ ({pq_value:.2f}) es mayor a las ventas simuladas."
    else:
        diagnostico_estres = "Las ventas simuladas son superiores al Punto de Quiebre, manteniendo la rentabilidad."
    text += f"Diagnóstico: {diagnostico_estres}\n"
    outB.insert(tk.END, text)
    analisis_B_text = text

def generar_analisis_economico(r, data):
    global analisis_C_text
    outC.delete("1.0", tk.END)
    text = "📈 ** SECCIÓN C: ANÁLISIS ECONÓMICO - RENTABILIDAD **\n\n"
    text += "C1. RENTABILIDAD ECONÓMICA (RAT) \n"
    text += f"RAT 2023: {r['RAT_2023']:.2f}% | RAT 2024: **{r['RAT_2024']:.2f}%**\n"
    text += f"Crecimiento: **{r['crecimiento_RAT']:.2f}%**.\n\n"
    text += "C2. RENTABILIDAD FINANCIERA (RRP) \n"
    text += f"RRP 2023: {r['RRP_2023']:.2f}% | RRP 2024: **{r['RRP_2024']:.2f}%**\n"
    apalancamiento_desc = 'positivo' if r['RRP_2024'] > r['RAT_2024'] else 'negativo'
    text += f"Relación: Apalancamiento financiero **{apalancamiento_desc}**.\n\n"
    text += "C3. ANÁLISIS DUPONT RRP 2024 \n"
    text += f"* Margen neto: **{r['margen_neto_dupont']:.4f}**\n"
    text += f"* Rotación del Activo: **{r['rotacion_activo']:.4f}**\n"
    text += f"* Apalancamiento: **{r['apalancamiento_dupont']:.4f}**\n"
    text += f"Verificación: RRP (Dupont) = {r['RRP_dupont_calc']:.2f}%.\n\n"
    text += "C4. MÁRGENES DE GANANCIA \n"
    text += f"a) Margen bruto: **{r['margen_bruto']:.2f}%**\n"
    text += f"b) Margen operativo: **{r['margen_operativo']:.2f}%**\n"
    text += f"c) Margen neto: **{r['margen_neto']:.2f}%**\n\n"
    text += "C5. APALANCAMIENTO FINANCIERO\n"
    text += f"a) Costo promedio de deuda (k): **{r['costo_deuda']:.2f}%**\n"
    text += f"b) RAT ({r['RAT_2024']:.2f}%) vs k ({r['costo_deuda']:.2f}%).\n"
    text += f"c) Conclusión: **{'CONVIENE ENDEUDARSE' if r['RAT_2024'] > r['costo_deuda'] else 'NO CONVIENE ENDEUDARSE'}**.\n"
    outC.insert(tk.END, text)
    analisis_C_text = text

def generar_diagnostico(r):
    global analisis_D_text
    outD.delete("1.0", tk.END)
    text = "🌟 ** SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO **\n\n"
    text += "D1. MATRIZ DE RATIOS COMPARATIVOS \n"
    text += "| Ratio | 2023 | 2024 | Interpretación |\n"
    text += f"| FM | {r['FM_2023']:.2f} | {r['FM_2024']:.2f} | {'Mejora' if r['FM_2024'] > r['FM_2023'] else 'Empeoró'} |\n"
    text += f"| Liq. Gral. | {r['LG_2023']:.2f} | {r['LG_2024']:.2f} | {'Empeoró' if r['LG_2024'] < r['LG_2023'] else 'Mejora'} |\n"
    text += f"| RAT (%) | {r['RAT_2023']:.2f} | {r['RAT_2024']:.2f} | {'Mejora' if r['RAT_2024'] > r['RAT_2023'] else 'Empeoró'} |\n"
    text += f"| RRP (%) | {r['RRP_2023']:.2f} | {r['RRP_2024']:.2f} | {'Mejora' if r['RRP_2024'] > r['RRP_2023'] else 'Empeoró'} |\n\n"
    text += "D2. FORTALEZAS Y DEBILIDADES \n"
    text += "✅ **FORTALEZAS**: Rentabilidad sólida (RRP > RAT), Autonomía financiera alta, Fondo de Maniobra positivo.\n"
    text += "⚠️ **DEBILIDADES**: Liquidez excesiva (recursos ociosos), Riesgo operativo ante caída de ventas (Punto de quiebre alto).\n\n"
    text += "D3. DIAGNÓSTICO FINANCIERO INTEGRAL \n"
    text += f"La empresa presenta un **ESTADO DE SALUD FINANCIERA MUY BUENO**, impulsado por una alta rentabilidad (**RRP {r['RRP_2024']:.2f}%**). El desafío es la **GESTIÓN DEL CAPITAL DE TRABAJO** (liquidez excesiva).\n\n"
    text += "D4. RECOMENDACIONES ESTRATÉGICAS \n"
    text += "a) Mejorar **Eficiencia Operativa**: Reducir Días Clientes.\n"
    text += f"b) Mejorar **Estructura Financiera**: Refinanciar parte del PC a Largo Plazo ({r['transferencia_deuda']:.2f} Bs).\n"
    text += f"c) Mejorar **Rentabilidad Operativa**: Reducir Gastos Administrativos en 10% ({r['mejora_BAII_reco']:.2f} Bs).\n"
    outD.insert(tk.END, text)
    analisis_D_text = text

def run_analisis(data):
    global analizador_obj 
    analizador_obj = AnalisisFinanciero(data)
    r = analizador_obj.calcular_todo()
    
    generar_analisis_patrimonial(r, data)
    generar_analisis_financiero(r, data)
    generar_analisis_economico(r, data)
    generar_diagnostico(r)
    
    # Limpiar figuras anteriores para liberar memoria
    for w in fig_frame.winfo_children():
        w.destroy()

    fig = Figure(figsize=(7, 6))
    
    # Subplot 1
    ax1 = fig.add_subplot(221)
    labels_a = ['Activo Corriente', 'Activo No Corriente']
    sizes_a = [r['vertical_AC'], r['vertical_ANC']]
    ax1.pie(sizes_a, labels=labels_a, autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#81C784'])
    ax1.set_title('Estructura del Activo 2024 (%)', fontsize=9)

    # Subplot 2
    ax2 = fig.add_subplot(222)
    labels_r = ['RAT 23', 'RAT 24', 'RRP 23', 'RRP 24']
    values_r = [r['RAT_2023'], r['RAT_2024'], r['RRP_2023'], r['RRP_2024']]
    x = np.arange(len(labels_r))
    ax2.bar(x, values_r, color=['#2196F3', '#1976D2', '#FF9800', '#F57C00'])
    ax2.set_xticks(x)
    ax2.set_xticklabels(labels_r, fontsize=8)
    ax2.set_title('Rentabilidad (%)', fontsize=9)
    
    # Subplot 3
    ax3 = fig.add_subplot(223)
    labels_l = ['LG 2024', 'Óptimo']
    values_l = [r['LG_2024'], 2.0]
    ax3.bar(labels_l, values_l, color=['#F44336', '#FFCDD2'])
    ax3.axhline(2.0, color='red', linestyle='--', linewidth=0.5)
    ax3.set_title('Liquidez General', fontsize=9)

    # Subplot 4
    ax4 = fig.add_subplot(224)
    labels_s = ['Garantía 24', 'Óptimo']
    values_s = [r['garantia_2024'], 1.5]
    ax4.bar(labels_s, values_s, color=['#9C27B0', '#E1BEE7'])
    ax4.axhline(1.5, color='purple', linestyle='--', linewidth=0.5)
    ax4.set_title('Garantía', fontsize=9)

    fig.tight_layout()

    # Guardar gráfico para PDF
    temp_file = "grafico_temp.png"
    try:
        fig.savefig(temp_file, bbox_inches='tight')
    except Exception as e:
        print(f"Error guardando gráfico: {e}")

    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)
    canvas.draw()
    
    return r, analizador_obj

def cargar_valores_por_defecto():
    defaults = {
        "Ingresos_2023": "8500", "Ingresos_2024": "11200", "Costo_2023": "3200", "Costo_2024": "4100",
        "GB_2023": "5300", "GB_2024": "7100", "GA_2023": "2100", "GA_2024": "2600",
        "GV_2023": "1200", "GV_2024": "1400", "DEP_2023": "400", "DEP_2024": "500",
        "BAII_2023": "1600", "BAII_2024": "2600", "GastosFin_2023": "100", "GastosFin_2024": "150",
        "UN_2023": "1162", "UN_2024": "1912", "AC_2023": "2800", "AC_2024": "3800",
        "ANC_2023": "1450", "ANC_2024": "1850", "PC_2023": "550", "PC_2024": "1000",
        "PNC_2023": "700", "PNC_2024": "1000", "PN_2023": "3000", "PN_2024": "3650",
        "Caja_2023": "850", "Caja_2024": "1100", "Clientes_2023": "1200", "Clientes_2024": "1600",
        "InvCP_2023": "300", "InvCP_2024": "500", "DI": "60", "DC": "45", "DP": "30",
    }
    for key, val in defaults.items():
        if key in form:
            form[key].delete(0, tk.END)
            form[key].insert(0, str(val))

def generar_pdf(r, data, archivo_nombre="Informe_Financiero_Innovatech.pdf"):
    """Genera el informe final en formato PDF con ReportLab de forma robusta."""
    try:
        f = open(archivo_nombre, 'a+')
        f.close()
    except PermissionError:
        messagebox.showerror("Error de Permisos", f"El archivo '{archivo_nombre}' está abierto. Ciérralo e intenta de nuevo.")
        return

    doc = SimpleDocTemplate(archivo_nombre, pagesize=letter)
    styles = getSampleStyleSheet()
    Story = []

    style_title = ParagraphStyle('Title', parent=styles['Title'], fontSize=18, spaceAfter=20, alignment=1)
    style_heading1 = ParagraphStyle('Heading1', parent=styles['Heading1'], fontSize=14, spaceBefore=12, spaceAfter=6, textColor=colors.blue)
    style_normal = styles['Normal']
    style_normal.spaceAfter = 8

    today = datetime.date.today().strftime("%d/%m/%Y")
    Story.append(Paragraph(f"PROYECTO FINAL - ANÁLISIS FINANCIERO", style_title))
    Story.append(Paragraph(f"Fecha: {today}", style_normal))
    Story.append(Spacer(1, 12))

    def add_text_section(title, text_content):
        if title:
            Story.append(Paragraph(title, style_heading1))
        
        if not text_content:
            return

        for line in text_content.split('\n'):
            line = line.strip()
            if line:
                # 1. Escapar caracteres HTML básicos (como <, >)
                safe_line = html.escape(line)
                
                # 2. Reemplazar **texto** por <b>texto</b> usando Regex (más seguro)
                # Esto busca pares de asteriscos y los convierte a tags bold
                formatted_line = re.sub(r'\*\*(.*?)\*\*', r'<b>\1</b>', safe_line)

                # 3. Detectar si es un título interno
                if any(line.startswith(x) for x in ["🏆", "💰", "📈", "🌟"]):
                    # Limpiar asteriscos y usar estilo de título
                    clean_line = formatted_line.replace('*', '')
                    Story.append(Paragraph(clean_line, style_heading1))
                else:
                    Story.append(Paragraph(formatted_line, style_normal))
        
        Story.append(Spacer(1, 12))

    # Añadir secciones
    add_text_section("", analisis_A_text)
    add_text_section("", analisis_B_text)
    add_text_section("", analisis_C_text)
    
    # Sección Diagnóstico y Tabla
    Story.append(Paragraph("SECCIÓN D: ANÁLISIS INTEGRAL", style_heading1))
    
    matriz_data = [
        ("Ratio", "2023", "2024", "Cambio"),
        ("FM (Bs)", f"{r['FM_2023']:.2f}", f"{r['FM_2024']:.2f}", "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró"),
        ("Liq. Gral.", f"{r['LG_2023']:.2f}", f"{r['LG_2024']:.2f}", "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora"),
        ("RAT (%)", f"{r['RAT_2023']:.2f}", f"{r['RAT_2024']:.2f}", "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró"),
        ("RRP (%)", f"{r['RRP_2023']:.2f}", f"{r['RRP_2024']:.2f}", "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró")
    ]
    
    table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2196F3')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
    ])
    
    table = Table(matriz_data, colWidths=[1.5*inch, 1*inch, 1*inch, 1.5*inch])
    table.setStyle(table_style)
    Story.append(table)
    Story.append(Spacer(1, 12))

    # Resto del texto D
    try:
        start_d2 = analisis_D_text.find("D2. FORTALEZAS")
        if start_d2 != -1:
            add_text_section("", analisis_D_text[start_d2:])
    except: pass

    # Imagen
    temp_file = "grafico_temp.png"
    img_added = False
    if os.path.exists(temp_file):
        Story.append(Spacer(1, 12))
        try:
            img = Image(temp_file, width=6*inch, height=5*inch)
            Story.append(img)
            img_added = True
        except: pass

    try:
        doc.build(Story)
        messagebox.showinfo("Éxito", f"Informe generado: '{archivo_nombre}'")
    except Exception as e:
        messagebox.showerror("Error", f"Error al crear PDF: {e}")
    finally:
        if img_added and os.path.exists(temp_file):
            try: os.remove(temp_file)
            except: pass

def run_all():
    try:
        data = read_all_inputs(form)
        r, analizador_obj = run_analisis(data)
        generar_pdf(r, data)
    except Exception as e:
        messagebox.showerror("Error Crítico", f"Ocurrió un error: {e}")
        print(e) # Para debug en consola

# Inicialización
cargar_valores_por_defecto()
root.mainloop()