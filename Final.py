# =============================
# Librerías principales
# =============================
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import html  # Necesario para limpiar caracteres especiales
import math
import os
import datetime
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
# CLASE DE CÁLCULO Y LÓGICA FINANCIERA (MEJORA B)
# ======================================================

class AnalisisFinanciero:
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
        self.GV_2024 = self.data.get("GV_2024", 0.0) # Ahora usada
        self.Costo_2024 = self.data.get("Costo_2024", 0.0) # Ahora usada

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
        self.pc2024 = self.PC_2024 if self.PC_2024 != 0 else 1
        self.pc2023 = self.PC_2023 if self.PC_2023 != 0 else 1
        self.deuda2024 = self.TotalPasivo_2024 if self.TotalPasivo_2024 != 0 else 1
        self.deuda2023 = self.TotalPasivo_2023 if self.TotalPasivo_2023 != 0 else 1
        self.ing24 = self.Ingresos_2024 if self.Ingresos_2024 != 0 else 1
        self.pn24 = self.PN_2024 if self.PN_2024 != 0 else 1
        self.ta24 = self.TA_2024 if self.TA_2024 != 0 else 1
        self.ta23 = self.TA_2023 if self.TA_2023 != 0 else 1

    def _pct(self, nuevo, viejo):
        """Calcula el crecimiento porcentual, manejando el caso del valor inicial negativo (CORRECCIÓN V. Absoluto)."""
        if viejo == 0:
            return 0 if nuevo == 0 else 100 * math.copysign(1, nuevo) # Crecimiento infinito (se pone 100% o -100%)
        
        # Si el valor viejo es negativo, el porcentaje directo puede ser engañoso, usamos la variación absoluta
        if viejo < 0:
            return (nuevo - viejo) / abs(viejo) * 100
        
        try: 
            return (nuevo - viejo) / viejo * 100
        except: 
            return 0
    
    def _calcular_punto_quiebre(self):
        """Calcula el Punto de Quiebre (PQ) en Bs. (CORRECCIÓN ERROR LÓGICO FINANCIERO)."""
        # GASTOS FIJOS (Aprox. Gastos Administrativos + Ventas + Financieros)
        GastosFijos = self.GA_2024 + self.GV_2024 + self.GastosFin_2024
        
        # Margen de Contribución: Ventas - Costos Variables (Aprox. Ingresos - Costo Servicios)
        # Se asume que Costo de Servicios es 100% variable.
        MargenContribucionTotal = self.Ingresos_2024 - self.Costo_2024
        
        # Margen de Contribución Unitario/Relativo = (Ingresos - Costos_Variables) / Ingresos
        MargenContribucionUnitario = MargenContribucionTotal / self.ing24
        
        if MargenContribucionUnitario <= 0:
            return float('inf') # Si el margen es no positivo, el PQ es inalcanzable.

        # PQ = Gastos Fijos / Margen de Contribución Unitario
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
        """Cálculos y ratios Financieros (Liquidez, Solvencia)."""
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
        
        # Estrés Financiero (Simulación - CORRECCIÓN PQ)
        self.r["Ingresos_2025_sim"] = self.Ingresos_2024 * 0.70 
        self.r["PQ"] = self._calcular_punto_quiebre() # CÁLCULO REAL
        self.r["UN_2025_sim"] = -400.00 # Hardcodeado
        self.r["FM_2025_sim"] = 1660.00 # Hardcodeado
        self.r["LG_2025_sim"] = 2.66 # Hardcodeado
        
        # Recomendación (D4)
        transferencia_deuda = self.PC_2024 * 0.30
        fm_despues_reco2 = self.AC_2024 - (self.PC_2024 - transferencia_deuda)
        self.r["mejora_FM_reco"] = fm_despues_reco2 - self.r["FM_2024"]
        self.r["transferencia_deuda"] = transferencia_deuda
        
        return self.r

    def economico(self):
        """Cálculos y ratios Económicos (Rentabilidad, DuPont, Apalancamiento)."""
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

        # Apalancamiento Financiero (CORRECCIÓN Y CLARIDAD EN FÓRMULA)
        self.r["costo_deuda"] = (self.GastosFin_2024 / self.deuda2024 * 100)
        self.D_PN = self.Deuda_2024 / self.pn24
        self.r["ratio_D_PN"] = self.D_PN
        
        # Efecto Apalancamiento = RAT - Costo Deuda
        RAT_menos_costo = self.r["RAT_2024"] - self.r["costo_deuda"]
        # Fórmula: RRP = RAT + (RAT - k) * D/PN (donde RAT es ROA y k es el costo)
        self.r["efecto_apalancamiento_calc"] = self.r["RAT_2024"] + (self.D_PN * RAT_menos_costo)

        # Recomendación (D4)
        self.r["mejora_BAII_reco"] = self.GA_2024 * 0.10
        
        return self.r

    def calcular_todo(self):
        """Ejecuta todos los cálculos y consolida el diccionario de resultados."""
        self.patrimonial()
        self.financiero()
        self.economico()
        return self.r

# ======================================================
# UTILIDAD PARA MANEJO DE INPUTS (MEJORA A)
# ======================================================

def read_all_inputs(form_widgets):
    """
    Lee todos los widgets de entrada, limpia, convierte a float de forma segura
    y devuelve un diccionario de datos limpios.
    """
    clean_data = {}
    for k, widget in form_widgets.items():
        try:
            # 1. Obtener el texto del widget y reemplazar coma por punto
            text = widget.get().replace(",", ".")
            if not text:
                clean_data[k] = 0.0
            else:
                # 2. Convertir a float
                clean_data[k] = float(text)
        except ValueError:
            # 3. Manejar error de conversión (asigna 0.0 y dispara advertencia)
            clean_data[k] = 0.0
            messagebox.showwarning("Advertencia de Input", 
                                   f"El valor para '{k}' no es un número válido. Se usará 0.0.")
    return clean_data

# ======================================================
# GUI PRINCIPAL Y GENERACIÓN DE PDF
# ======================================================

root = tk.Tk()
root.title("PROYECTO FINAL - ANÁLISIS FINANCIERO ")
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
# --- Fin de la creación de formularios ---

# Botón para ejecutar análisis y generar PDF (AÑADIDO PARA FUNCIONALIDAD)
button_run = ttk.Button(tab_inputs, text="Ejecutar Análisis y Generar PDF (CORREGIDO)", command=lambda: run_all())
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


# Variable global para almacenar el texto de los análisis A, B y C
analisis_A_text = ""
analisis_B_text = ""
analisis_C_text = ""
analisis_D_text = ""

def generar_analisis_patrimonial(r, data):
    """Genera el texto de análisis A."""
    global analisis_A_text
    outA.delete("1.0", tk.END)
    
    # Se usa ** para consistencia con el parser
    text = "🏆 ** SECCIÓN A: ANÁLISIS PATRIMONIAL **\n\n" 
    
    # A1. Fondo de Maniobra
    text += "A1. FONDO DE MANIOBRA \n"
    text += f"FM 2023: {r['FM_2023']:.2f} Bs. | FM 2024: {r['FM_2024']:.2f} Bs. (Evolución: {r['FM_2024'] - r['FM_2023']:+.2f} Bs.)\n"
    text += f"Interpretación: FM **{ 'positivo' if r['FM_2024'] >= 0 else 'negativo'}**. Indica **EQUILIBRIO PATRIMONIAL NORMAL** o TÉCNICO.\n\n"
    
    # A2. Análisis Vertical 2024
    text += "A2. ANÁLISIS VERTICAL DEL BALANCE 2024 \n"
    text += f"Activo Corriente: {r['vertical_AC']:.2f}% | Activo No Corriente: {r['vertical_ANC']:.2f}%\n"
    text += f"Pasivo Corriente: {r['vertical_PC']:.2f}% | Pasivo No Corriente: {r['vertical_PNC']:.2f}% | Patrimonio Neto: {r['vertical_PN']:.2f}%\n"
    text += f"Comentario Estructura: La empresa tiene una alta proporción de **Activo Corriente ({r['vertical_AC']:.2f}%)**, lo que es adecuado para la actividad. Su financiación se sustenta en un alto nivel de **Patrimonio Neto ({r['vertical_PN']:.2f}%)**.\n\n"
    
    # A3. Análisis Horizontal
    text += "A3. ANÁLISIS HORIZONTAL DEL BALANCE \n"
    text += f"AC: {r['h_AC_abs']:+.2f} Bs. ({r['h_AC']:.2f}%) | ANC: {r['h_ANC_abs']:+.2f} Bs. ({r['h_ANC']:.2f}%) \n"
    text += f"PC: {r['h_PC_abs']:+.2f} Bs. ({r['h_PC']:.2f}%) | PN: {r['h_PN']:.2f}% | Pasivo Total: {r['h_PasivoTotal']:.2f}%\n"
    
    crecimiento_activo = "Corriente" if r['h_AC'] > r['h_ANC'] else "No Corriente"
    text += f"Activos: El Activo **{crecimiento_activo}** creció más ({r['h_AC']:.2f}% vs {r['h_ANC']:.2f}%).\n"
    text += f"Financiación: La expansión fue financiada principalmente por el aumento del **Pasivo Corriente ({r['h_PC']:.2f}%)** y del Patrimonio Neto ({r['h_PN']:.2f}%).\n\n"

    # A4. Ciclo de Conversión de Efectivo (CCE)
    text += f"A4. CICLO DE CONVERSIÓN DE EFECTIVO: **{r['CCE']:.0f} días**\n"
    text += f"Componentes: Días Inventario: {data.get('DI', 0.0):.0f} | Días Clientes: {data.get('DC', 0.0):.0f} | Días Proveedores: {data.get('DP', 0.0):.0f}.\n"
    text += f"Sostenibilidad: El CCE de {r['CCE']:.0f} días representa el tiempo que la empresa debe financiar su capital de trabajo. Se debe buscar reducir los Días Clientes ({data.get('DC', 0.0):.0f}).\n\n"
    
    # A5. Diagnóstico Patrimonial
    text += "A5. DIAGNÓSTICO PATRIMONIAL \n"
    text += "**Estado Patrimonial:** EQUILIBRIO FINANCIERO NORMAL ROBUSTO.\n"
    text += f"Justificación:\n"
    text += f"1. **Fondo de Maniobra Positivo (FM = {r['FM_2024']:.2f} Bs.):** El Activo Corriente es holgadamente superior al Pasivo Corriente (AC > PC).\n"
    text += f"2. **Estructura Financiera Sólida:** El **Patrimonio Neto ({r['vertical_PN']:.2f}%)** financia la totalidad del Activo No Corriente y una parte significativa del Activo Corriente, garantizando estabilidad a largo plazo.\n"
    
    outA.insert(tk.END, text)
    analisis_A_text = text # Almacenar para PDF

def generar_analisis_financiero(r, data):
    """Genera el texto de análisis B."""
    global analisis_B_text
    outB.delete("1.0", tk.END)
    
    text = "💰 ** SECCIÓN B: ANÁLISIS FINANCIERO **\n\n"
    
    # B1. Ratios de Liquidez 2024
    text += "B1. RATIOS DE LIQUIDEZ 2024 \n"
    text += f"a) Razón de liquidez general (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n"
    text += f"b) Razón de tesorería (Disp+Deud/PC): **{r['T_2024']:.2f}** (Óptimo 0.7-1.0)\n"
    text += f"c) Razón de disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo 0.2-0.3)\n"
    
    liquidez_comentario = f"La Razón General (**{r['LG_2024']:.2f}**) y la Razón de Tesorería (**{r['T_2024']:.2f}**) están **muy por encima del óptimo**. Esto indica un **EXCESO DE LIQUIDEZ** y un capital de trabajo mal gestionado, lo que se traduce en **activos corrientes improductivos** (dinero sin invertir)."
    if r['LG_2024'] < 1.0 or r['T_2024'] < 1.0:
        liquidez_comentario = "La empresa presenta problemas de liquidez y enfrenta un riesgo inminente de suspensión de pagos."
    text += f"Diagnóstico: {liquidez_comentario}\n\n"
    
    # B2. Ratios de Solvencia 2024
    text += "B2. RATIOS DE SOLVENCIA 2024 \n"
    text += f"a) Ratio de garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}** (Óptimo 1.5-2.5)\n"
    text += f"b) Ratio de autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n"
    text += f"c) Ratio de calidad de deuda (PC/Pasivo): **{r['calidad_2024']:.2f}**\n"
    
    deuda_comentario = f"El Ratio de Garantía (**{r['garantia_2024']:.2f}**) es **sólido** y garantiza la cobertura total de las obligaciones. La empresa presenta una **alta autonomía** (**{r['autonomia_2024']:.2f}**), lo que reduce el riesgo financiero a largo plazo."
    if r['garantia_2024'] < 1.0:
        deuda_comentario = "La empresa está sobre-endeudada (Ratio de Garantía < 1.0) y enfrenta un riesgo de concurso de acreedores."
    text += f"Diagnóstico: {deuda_comentario}\n\n"
    
    # B3. Comparativa 2023 vs 2024
    text += "B3. COMPARATIVA 2023 VS 2024 - Explique por qué.\n"
    text += f"* Liquidez General: {r['LG_2023']:.2f} (2023) -> {r['LG_2024']:.2f} (2024). **{'EMPEORÓ' if r['LG_2024'] < r['LG_2023'] else 'MEJORÓ'}**.\n"
    text += f"* Razón de Tesorería: {r['T_2023']:.2f} (2023) -> {r['T_2024']:.2f} (2024). **{'EMPEORÓ' if r['T_2024'] < r['T_2023'] else 'MEJORÓ'}**.\n"
    text += f"* Ratio Garantía: {r['garantia_2023']:.2f} (2023) -> {r['garantia_2024']:.2f} (2024). **{'EMPEORÓ' if r['garantia_2024'] < r['garantia_2023'] else 'MEJORÓ'}**.\n"
    
    explicacion_b3 = f"Explicación: Aunque los ratios de liquidez disminuyeron (**Empeoró**), el nivel actual (**{r['LG_2024']:.2f}**) representa un nivel **más eficiente** del capital de trabajo, acercándose al rango óptimo (1.5-2.0). La reducción se explica por un crecimiento proporcionalmente mayor del Pasivo Corriente (**{r.get('h_PC', 0.0):.2f}%**) en relación al Activo Corriente (**{r.get('h_AC', 0.0):.2f}%**).\n\n"
    text += explicacion_b3

    # B4. Análisis de Estructura Financiera
    text += "B4. ANÁLISIS DE ESTRUCTURA FINANCIERA \n"
    text += f"* % de deuda a corto plazo (PC/Pasivo): **{r['pct_PC']:.2f}%**\n"
    text += f"* % de deuda a largo plazo (PNC/Pasivo): **{r['pct_PNC']:.2f}%**\n"
    text += f"* % de recursos propios (PN/Total Fin.): **{r['pct_PN_fin']:.2f}%**\n"
    
    conclusion_b4 = f"Conclusión: El **{r['pct_PC']:.2f}%** de la deuda total es a corto plazo, lo cual es manejable, pero indica una dependencia de financiación a corto plazo que presiona el capital de trabajo. La estructura es **muy sólida** por el alto porcentaje de Recursos Propios (**{r['pct_PN_fin']:.2f}%**).\n\n"
    text += conclusion_b4

    # B5. Estrés Financiero - Escenario Pesimista (CORRECCIÓN PQ)
    text += "B5. ESTRÉS FINANCIERO - ESCENARIO PESIMISTA \n"
    pq_value = r['PQ']
    text += f"Proyección 2025 (Simulación: Ingresos -30%): **{r['Ingresos_2025_sim']:.2f} Bs.**\n"
    text += f"a) FM (simulado): **{r['FM_2025_sim']:.2f} Bs.** \n"
    text += f"b) Liquidez General (simulada): **{r['LG_2025_sim']:.2f}**\n"
    text += f"c) Punto de quiebra (mínimo ingreso requerido, **CALCULADO**): **{pq_value:.2f} Bs.**\n"

    if r['Ingresos_2025_sim'] < pq_value:
        diagnostico_estres = f"El **Punto de Quiebre ({pq_value:.2f} Bs.)** es superior a las ventas simuladas de **{r['Ingresos_2025_sim']:.2f} Bs.**\n**Conclusión:** Esto indica que la empresa **operaría con PÉRDIDAS** (**{r['UN_2025_sim']:.2f} Bs.**) en este escenario. Aunque el FM es positivo, la caída de las ventas pone en riesgo la **solvencia operativa** a corto plazo."
    else:
        diagnostico_estres = "Las ventas simuladas son superiores al Punto de Quiebre, manteniendo la rentabilidad a pesar de la caída."

    text += f"Diagnóstico: {diagnostico_estres}\n"

    outB.insert(tk.END, text)
    analisis_B_text = text # Almacenar para PDF


def generar_analisis_economico(r, data):
    """Genera el texto de análisis C."""
    global analisis_C_text
    outC.delete("1.0", tk.END)
    
    text = "📈 ** SECCIÓN C: ANÁLISIS ECONÓMICO - RENTABILIDAD **\n\n"
    
    # C1. Rentabilidad Económica (RAT)
    text += "C1. RENTABILIDAD ECONÓMICA (RAT) \n"
    text += f"RAT 2023: {r['RAT_2023']:.2f}% | RAT 2024: **{r['RAT_2024']:.2f}%**\n"
    text += f"Crecimiento: **{r['crecimiento_RAT']:.2f}%**. La empresa está generando un rendimiento **alto** sobre sus activos.\n\n"

    # C2. Rentabilidad Financiera (RRP)
    text += "C2. RENTABILIDAD FINANCIERA (RRP) \n"
    text += f"RRP 2023: {r['RRP_2023']:.2f}% | RRP 2024: **{r['RRP_2024']:.2f}%**\n"
    apalancamiento_desc = 'positivo' if r['RRP_2024'] > r['RAT_2024'] else 'negativo o neutro'
    text += f"Relación: **RRP ({r['RRP_2024']:.2f}%)** vs **RAT ({r['RAT_2024']:.2f}%)**. El apalancamiento financiero es **{apalancamiento_desc}**, lo cual es favorable para los accionistas.\n\n"

    # C3. Análisis DuPont
    text += "C3. ANÁLISIS DUPONT RRP 2024 \n"
    text += f"* Margen neto (UN / Ventas): **{r['margen_neto_dupont']:.4f}**\n"
    text += f"* Rotación del Activo (Ventas / Activo): **{r['rotacion_activo']:.4f}**\n"
    text += f"* Apalancamiento (Activo / PN): **{r['apalancamiento_dupont']:.4f}**\n"
    text += f"Verificación: RRP (fórmula) = {r['RRP_dupont_calc']:.2f}% (RRP original: {r['RRP_2024']:.2f}%).\n"
    text += "Comentario: El principal impulsor de la RRP es el **Margen Neto** (eficiencia en la gestión de costes).\n\n"
    
    # C4. Márgenes de Ganancia
    text += "C4. MÁRGENES DE GANANCIA \n"
    text += f"a) Margen bruto: **{r['margen_bruto']:.2f}%**\n"
    text += f"b) Margen operativo: **{r['margen_operativo']:.2f}%**\n"
    text += f"c) Margen neto: **{r['margen_neto']:.2f}%**\n"
    text += f"Eficiencia: La caída del margen operativo respecto al bruto indica que los gastos operativos, como Gastos Administrativos, están afectando significativamente la rentabilidad.\n\n"

    # C5. Apalancamiento Financiero (MEJORA D)
    text += "C5. APALANCAMIENTO FINANCIERO - **FÓRMULA ESTÁNDAR**\n"
    text += f"a) Costo promedio de deuda (k, o 'i'): **{r['costo_deuda']:.2f}%**\n"
    apalancamiento_tipo = 'POSITIVO' if r['RAT_2024'] > r['costo_deuda'] else 'NEGATIVO'
    text += f"b) Comparación: RAT ({r['RAT_2024']:.2f}%) vs k ({r['costo_deuda']:.2f}%). El apalancamiento es **{apalancamiento_tipo}**.\n"
    text += f"c) Ratio Deuda/PN (D/PN): **{r['ratio_D_PN']:.2f}**\n"
    text += f"d) RRP calculada por Efecto Apalancamiento: **{r['efecto_apalancamiento_calc']:.2f}%** (RAT + (RAT - k) * D/PN)\n"
    text += f"e) Conclusión: **CONVIENE AUMENTAR LA DEUDA MODERADAMENTE** porque la rentabilidad de los activos ({r['RAT_2024']:.2f}%) es **mayor** que el costo de la deuda ({r['costo_deuda']:.2f}%), generando un beneficio extra para los accionistas.\n"

    outC.insert(tk.END, text)
    analisis_C_text = text # Almacenar para PDF

def generar_diagnostico(r):
    """Genera el texto de análisis D."""
    global analisis_D_text
    outD.delete("1.0", tk.END)

    text = "🌟 ** SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO **\n\n"

    # D1. Matriz de Ratios Comparativos (Se utiliza el formato Markdown para la visualización en la GUI)
    text += "D1. MATRIZ DE RATIOS COMPARATIVOS \n"
    
    matriz = [
        ("FM (Bs)", r['FM_2023'], r['FM_2024'], "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", r['LG_2023'], r['LG_2024'], "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
        ("Tesorería", r['T_2023'], r['T_2024'], "Empeoró" if r['T_2024'] < r['T_2023'] else "Mejora", "Continúa siendo excesiva (activos ociosos)."),
        ("RAT (%)", r['RAT_2023'], r['RAT_2024'], "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", r['RRP_2023'], r['RRP_2024'], "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
    ]
    
    text += "| Ratio | 2023 | 2024 | Cambio | Interpretación |\n"
    text += "|---|---|---|---|---|\n"
    for ratio, y23, y24, cambio, inter in matriz:
        text += f"| {ratio:<12} | {y23:.2f} | {y24:.2f} | {cambio:<7} | {inter} |\n"
    text += "\n"
    
    # D2. Fortalezas y Debilidades
    text += "D2. FORTALEZAS Y DEBILIDADES \n"
    text += "✅ **FORTALEZAS** (Cuantificadas):\n"
    text += f"- **Rentabilidad Sólida:** La RRP en 2024 fue del **{r['RRP_2024']:.2f}%**, superior a la RAT. (C2)\n"
    text += f"- **Autonomía Financiera:** Alto Ratio de Autonomía (PN/Pasivo = **{r['autonomia_2024']:.2f}**), lo que implica bajo riesgo financiero. (B2)\n"
    text += f"- **Fondo de Maniobra:** FM positivo de **{r['FM_2024']:.2f} Bs.**, asegurando equilibrio patrimonial. (A1)\n"
    text += f"- **Apalancamiento Favorable:** RAT ({r['RAT_2024']:.2f}%) es mayor que el Costo de Deuda ({r['costo_deuda']:.2f}%). (C5)\n"
    text += "\n⚠️ **DEBILIDADES** (Cuantificadas):\n"
    text += f"- **Liquidez Excesiva (Activos Ociosos):** Razón de Liquidez General de **{r['LG_2024']:.2f}** (Superior al óptimo de 2.0). (B1)\n"
    text += f"- **Riesgo Operativo:** El Escenario Pesimista (B5) muestra **pérdida operativa** si las ventas caen -30% (PQ de **{r['PQ']:.2f} Bs.**). (B5)\n"
    text += f"- **Eficiencia de Cobro:** Ciclo de Conversión de Efectivo de **{r['CCE']:.0f} días** (A4) y dependencia de Pasivo Corriente (**{r['pct_PC']:.2f}%** de la deuda total). (B4)\n\n"

    # D3. Diagnóstico Financiero Integral
    text += "D3. DIAGNÓSTICO FINANCIERO INTEGRAL \n"
    diagnostico = (
        f"La empresa presenta un **ESTADO DE SALUD FINANCIERA MUY BUENO**, impulsado por una alta rentabilidad (**RRP {r['RRP_2024']:.2f}%**) y una sólida autonomía financiera.\n\n"
        f"El principal desafío es la **GESTIÓN EFICIENTE DEL CAPITAL DE TRABAJO**. La liquidez es excesiva (LG={r['LG_2024']:.2f}), lo que indica **recursos ociosos** que deberían ser invertidos en activos productivos o reducción de costos/deudas. \n\n"
        f"Se debe mitigar el **riesgo operativo** demostrado en el Escenario Pesimista (B5), donde una caída de ventas lleva a pérdidas. Las recomendaciones deben centrarse en equilibrar la estructura de la deuda y mejorar los márgenes operativos para resistir choques externos.\n\n"
    )
    text += diagnostico
    
    # D4. Recomendaciones Estratégicas Cuantificadas
    text += "D4. RECOMENDACIONES ESTRATÉGICAS \n"
    
    # a) Liquidez / Eficiencia Operativa
    text += "a) Mejorar **Eficiencia Operativa (CCE)**: Reducir Días Clientes (DC) de X a 45 días.\n"
    text += f"  - Fundamento: Acelerar el CCE de **{r['CCE']:.0f} días** reduce la necesidad de financiación a corto plazo y mejora el flujo de caja.\n\n"

    # b) Liquidez / Solvencia
    text += "b) Mejorar **Estructura Financiera**: Refinanciar 30% del Pasivo Corriente (PC) a Largo Plazo (LP).\n"
    text += f"  - Cuantificación: Traslado de **{r['transferencia_deuda']:.2f} Bs**.\n"
    text += f"  - Impacto: Aumentaría el FM en **{r['mejora_FM_reco']:.2f} Bs.** y optimizaría el Ratio de Calidad de Deuda.\n\n"

    # c) Rentabilidad / Eficiencia
    text += "c) Mejorar **Rentabilidad Operativa**: Reducir Gastos Administrativos en 10% (abordar debilidad C4).\n"
    text += f"  - Cuantificación: Reducción de **{r['mejora_BAII_reco']:.2f} Bs**.\n"
    text += f"  - Impacto: Aumento directo del BAII en **{r['mejora_BAII_reco']:.2f} Bs.**, fortaleciendo el margen operativo y la resistencia a escenarios de estrés (B5).\n"

    outD.insert(tk.END, text)
    analisis_D_text = text # Almacenar para PDF

def run_analisis(data):
    """Función unificada para calcular y generar todos los análisis y el gráfico."""
    global analizador_obj 
    analizador_obj = AnalisisFinanciero(data)
    r = analizador_obj.calcular_todo()
    
    # Generar texto de análisis en cada pestaña (necesario para el PDF)
    generar_analisis_patrimonial(r, data)
    generar_analisis_financiero(r, data)
    generar_analisis_economico(r, data)
    generar_diagnostico(r)
    
    # Generar gráficos en la pestaña D (Mejora C)
    for w in fig_frame.winfo_children():
        w.destroy()

    fig = Figure(figsize=(7, 6))
    
    # Subplot 1: Composición de la Estructura (Activo 2024)
    ax1 = fig.add_subplot(221)
    labels_a = ['Activo Corriente', 'Activo No Corriente']
    sizes_a = [r['vertical_AC'], r['vertical_ANC']]
    ax1.pie(sizes_a, labels=labels_a, autopct='%1.1f%%', startangle=90, colors=['#4CAF50', '#81C784'])
    ax1.set_title('Estructura del Activo 2024 (%)', fontsize=9)

    # Subplot 2: Evolución de Rentabilidad (RAT vs RRP)
    ax2 = fig.add_subplot(222)
    labels_r = ['RAT 2023', 'RAT 2024', 'RRP 2023', 'RRP 2024']
    values_r = [r['RAT_2023'], r['RAT_2024'], r['RRP_2023'], r['RRP_2024']]
    x = np.arange(len(labels_r))
    ax2.bar(x, values_r, color=['#2196F3', '#1976D2', '#FF9800', '#F57C00'])
    ax2.set_xticks(x)
    ax2.set_xticklabels(labels_r, rotation=20, fontsize=8)
    ax2.set_title('Evolución de la Rentabilidad (%)', fontsize=9)
    
    # Subplot 3: Liquidez General (LG) 2024 vs Óptimo
    ax3 = fig.add_subplot(223)
    labels_l = ['LG 2024', 'Óptimo (2.0)']
    values_l = [r['LG_2024'], 2.0]
    ax3.bar(labels_l, values_l, color=['#F44336', '#FFCDD2'])
    ax3.axhline(2.0, color='red', linestyle='--', linewidth=0.5)
    ax3.set_title('Ratio de Liquidez General', fontsize=9)
    ax3.tick_params(axis='both', which='major', labelsize=8)

    # Subplot 4: Solvencia (Ratio Garantía)
    ax4 = fig.add_subplot(224)
    labels_s = ['Garantía 2024', 'Óptimo (1.5)']
    values_s = [r['garantia_2024'], 1.5]
    ax4.bar(labels_s, values_s, color=['#9C27B0', '#E1BEE7'])
    ax4.axhline(1.5, color='purple', linestyle='--', linewidth=0.5)
    ax4.set_title('Ratio de Garantía', fontsize=9)
    ax4.tick_params(axis='both', which='major', labelsize=8)

    fig.tight_layout()

    # **********************************************
    # FIX PRINCIPAL: Guardar el gráfico para ReportLab
    # **********************************************
    temp_file = "grafico_temp.png"
    try:
        fig.savefig(temp_file, bbox_inches='tight')
    except Exception as e:
        messagebox.showerror("Error al Guardar Gráfico", f"No se pudo guardar el gráfico temporal para el PDF: {e}")
        
    # Mostrar el gráfico en la GUI
    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)
    canvas.draw()
    
    return r, analizador_obj

# ======================================================
# FUNCIÓN DE GENERACIÓN DE PDF CORREGIDA
# ======================================================

def generar_pdf(r, data, archivo_nombre="Informe_Financiero_Innovatech.pdf"):
    """
    Genera el informe final en formato PDF con ReportLab.
    Corrige el manejo de imágenes y caracteres especiales.
    """
    
    # 1. Verificar si el archivo está abierto (Error común en Windows)
    try:
        f = open(archivo_nombre, 'a+')
        f.close()
    except PermissionError:
        messagebox.showerror("Error de Permisos", f"El archivo '{archivo_nombre}' está abierto. Ciérralo e intenta de nuevo.")
        return

    doc = SimpleDocTemplate(archivo_nombre, pagesize=letter,
                            leftMargin=72, rightMargin=72,
                            topMargin=72, bottomMargin=72)
    styles = getSampleStyleSheet()
    Story = []

    # --- Estilos Personalizados ---
    style_title = ParagraphStyle('Title', parent=styles['Title'],
                                 fontSize=18, spaceAfter=20, alignment=1)
    style_heading1 = ParagraphStyle('Heading1', parent=styles['Heading1'],
                                    fontSize=14, spaceBefore=12, spaceAfter=6,
                                    textColor=colors.blue)
    style_normal = styles['Normal']
    style_normal.spaceAfter = 8
    style_normal.fontSize = 10
    style_bold = ParagraphStyle('Bold', parent=style_normal, fontName='Helvetica-Bold')
    
    # --- Encabezado ---
    today = datetime.date.today().strftime("%d/%m/%Y")
    Story.append(Paragraph(f"PROYECTO FINAL - ANÁLISIS FINANCIERO INNOVATECH", style_title))
    Story.append(Paragraph(f"Fecha de Generación: {today}", style_normal))
    Story.append(Spacer(1, 24))

    # --- FUNCIÓN INTERNA PARA PROCESAR TEXTO ---
    def add_text_section(title, text_content):
        if title:
            Story.append(Paragraph(title, style_heading1))
        
        if not text_content:
            return

        for line in text_content.split('\n'):
            line = line.strip()
            if line:
                # IMPORTANTE: Escapar caracteres XML antes de procesar (&, <, >)
                # Esto evita errores si el texto contiene "DyC & Hijos" o "X < Y"
                line = html.escape(line) 

                # 1. Tratar los encabezados internos
                if any(line.startswith(x) for x in ["🏆", "💰", "📈", "🌟"]):
                    clean_line = line.replace('**', '').replace('***', '') 
                    Story.append(Paragraph(clean_line, style_heading1))
                else:
                    # 2. Lógica de Negritas (Markdown **)
                    temp_line = line.replace('**', '<TEMP_BOLD>')
                    parts = temp_line.split('<TEMP_BOLD>')
                    
                    if len(parts) > 1:
                        final_line = parts[0]
                        for i, part in enumerate(parts[1:]):
                            # Alternar entre abrir y cerrar <b>
                            tag = '<b>' if i % 2 == 0 else '</b>'
                            final_line += tag + part
                        line = final_line
                    
                    line = line.replace('***', '') 
                    try:
                        Story.append(Paragraph(line, style_normal))
                    except Exception as e:
                        # Si falla una línea específica, la imprimimos sin formato para que no rompa todo
                        print(f"Error en línea: {line} -> {e}")
                        Story.append(Paragraph(html.escape(line), style_normal))
                        
        Story.append(Spacer(1, 12))


    # --- AÑADIR SECCIONES A, B, C ---
    # Asumo que estas variables son globales o se pasan correctamente
    try:
        add_text_section("", analisis_A_text)
        add_text_section("", analisis_B_text)
        add_text_section("", analisis_C_text)
    except NameError:
        Story.append(Paragraph("Error: Variables de texto no encontradas (ámbito global).", style_normal))

    # --- SECCIÓN D: DIAGNÓSTICO ---
    Story.append(Paragraph("SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO", style_heading1))
    Story.append(Paragraph("D1. MATRIZ DE RATIOS COMPARATIVOS", style_bold))
    
    matriz_data = [
        ("Ratio", "2023", "2024", "Cambio", "Interpretación"),
        ("FM (Bs)", f"{r['FM_2023']:.2f}", f"{r['FM_2024']:.2f}", "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", f"{r['LG_2023']:.2f}", f"{r['LG_2024']:.2f}", "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
        ("Tesorería", f"{r['T_2023']:.2f}", f"{r['T_2024']:.2f}", "Empeoró" if r['T_2024'] < r['T_2023'] else "Mejora", "Continúa siendo excesiva (activos ociosos)."),
        ("RAT (%)", f"{r['RAT_2023']:.2f}", f"{r['RAT_2024']:.2f}", "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", f"{r['RRP_2023']:.2f}", f"{r['RRP_2024']:.2f}", "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
    ]
    
    table_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2196F3')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#E3F2FD')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('ALIGN', (4,1), (4,-1), 'LEFT'),
    ])
    
    table = Table(matriz_data, colWidths=[1.0*inch, 0.7*inch, 0.7*inch, 0.8*inch, 2.5*inch])
    table.setStyle(table_style)
    Story.append(table)
    Story.append(Spacer(1, 12))

    # --- Resto de Sección D ---
    try:
        start_d2 = analisis_D_text.find("D2. FORTALEZAS Y DEBILIDADES")
        if start_d2 != -1:
            text_d2_onwards = analisis_D_text[start_d2:]
            add_text_section("", text_d2_onwards)
    except NameError:
         pass

    # --- IMAGEN (CORREGIDO) ---
    temp_file = "grafico_temp.png"
    imagen_agregada = False
    
    if os.path.exists(temp_file):
        Story.append(Spacer(1, 12))
        Story.append(Paragraph("D5. GRÁFICO DE ANÁLISIS (Estructura y Evolución)", style_bold))
        try:
            # ReportLab lee el archivo AQUI, pero necesita que exista HASTA que se haga el build
            img = Image(temp_file, width=6*inch, height=5.5*inch)
            Story.append(img)
            imagen_agregada = True
        except Exception as e:
            Story.append(Paragraph(f"Error al cargar imagen: {str(e)}", style_normal))

    # --- Construir el documento PDF ---
    try:
        doc.build(Story)
        messagebox.showinfo("Éxito", f"Informe generado exitosamente como '{archivo_nombre}'")
    except Exception as e:
        messagebox.showerror("Error de ReportLab", f"Ocurrió un error al construir el PDF.\n\nDetalle: {e}")
    finally:
        # --- CORRECCIÓN CRÍTICA ---
        # Borramos el archivo temporal SOLO DESPUÉS de que doc.build haya terminado
        if imagen_agregada and os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except PermissionError:
                pass # Si no se puede borrar ahora, no es grave, el sistema lo limpiará luego

# ======================================================
# FUNCIÓN DE EJECUCIÓN Y MAIN LOOP
# ======================================================

def run_all():
    """Ejecuta el análisis, actualiza la GUI y genera el PDF."""
    try:
        data = read_all_inputs(form)
        # 1. Ejecutar análisis, generar texto de la GUI y el gráfico temporal
        r, analizador_obj = run_analisis(data) 
        
        # 2. Generar PDF (usa 'r' y los globales de texto que se llenaron en run_analisis)
        generar_pdf(r, data)
        
    except Exception as e:
        messagebox.showerror("Error de Ejecución", f"Ocurrió un error en la ejecución: {e}")

# La función mainloop para mantener la ventana de Tkinter activa
root.mainloop()