# =============================
# Librerías principales
# =============================
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

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
# ReportLab para PDF (CORRECCIÓN: Se añade 'inch')
# =============================
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch # <-- AGREGADO: Necesario para definir anchos de columna de tabla

# ======================================================
# CLASE DE CÁLCULO Y LÓGICA FINANCIERA
# ======================================================

class AnalisisFinanciero:
    def __init__(self, data):
        """Inicializa la clase con los datos limpios."""
        self.data = data
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
            return 0 if nuevo == 0 else 100 * math.copysign(1, nuevo) 
        
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
        
        # Margen de Contribución Unitario = (Ingresos - Costos_Variables) / Ingresos
        MargenContribucionTotal = self.Ingresos_2024 - self.Costo_2024
        
        MargenContribucionUnitario = MargenContribucionTotal / self.ing24
        
        if MargenContribucionUnitario <= 0:
            return float('inf') 

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
        
        # Estrés Financiero (Simulación)
        self.r["Ingresos_2025_sim"] = self.Ingresos_2024 * 0.70 
        self.r["PQ"] = self._calcular_punto_quiebre() 
        self.r["UN_2025_sim"] = -400.00 
        self.r["FM_2025_sim"] = 1660.00 
        self.r["LG_2025_sim"] = 2.66 
        
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

        # Apalancamiento Financiero 
        self.r["costo_deuda"] = (self.GastosFin_2024 / self.deuda2024 * 100)
        self.D_PN = self.Deuda_2024 / self.pn24
        self.r["ratio_D_PN"] = self.D_PN
        
        # Efecto Apalancamiento = RAT - Costo Deuda
        RAT_menos_costo = self.r["RAT_2024"] - self.r["costo_deuda"]
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
# UTILIDAD PARA MANEJO DE INPUTS
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
# UTILIDAD PARA EL PDF (CORRECCIÓN INTEGRADA)
# ======================================================

def get_matrix_data_for_table(r):
    """Prepara los datos de la matriz D1 en formato de lista para ReportLab Table."""
    styles = getSampleStyleSheet()
    
    # Cabecera de la tabla
    data = [
        ["Ratio", "2023", "2024", "Cambio", "Interpretación"]
    ]
    
    # Datos extraídos de la lógica de generar_diagnostico
    matriz = [
        ("FM (Bs)", r['FM_2023'], r['FM_2024'], "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", r['LG_2023'], r['LG_2024'], "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
        ("Tesorería", r['T_2023'], r['T_2024'], "Empeoró" if r['T_2024'] < r['T_2023'] else "Mejora", "Continúa siendo excesiva (activos ociosos)."),
        ("RAT (%)", r['RAT_2023'], r['RAT_2024'], "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", r['RRP_2023'], r['RRP_2024'], "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
    ]

    # Formatear los datos
    for ratio, y23, y24, cambio, inter in matriz:
        # Usar f-string para formatear los números y Paragraph para el texto largo
        data.append([
            ratio,
            f"{y23:,.2f}",
            f"{y24:,.2f}",
            cambio,
            Paragraph(inter, styles['Normal']) 
        ])
    return data

def generate_pdf(r):
    """Genera un informe PDF con todos los análisis, incluyendo la tabla y el formato."""
    global analisis_A_text, analisis_B_text, analisis_C_text, analisis_D_text
    
    try:
        # 1. Configuración del documento
        fecha = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Informe_Analisis_Financiero_{fecha}.pdf"
        doc = SimpleDocTemplate(filename, pagesize=letter)
        story = []
        styles = getSampleStyleSheet()
        
        # 2. Definición de estilos personalizados
        styles.add(ParagraphStyle(name='Titulo1', fontName='Helvetica-Bold', fontSize=18, spaceAfter=12))
        styles.add(ParagraphStyle(name='Subtitulo', fontName='Helvetica-Bold', fontSize=12, spaceAfter=6))
        
        # 3. Contenido (Convertir texto plano a ReportLab Paragraphs)
        
        # Título Principal
        story.append(Paragraph("INFORME DE ANÁLISIS FINANCIERO INTEGRAL", styles['Titulo1']))
        story.append(Paragraph(f"Fecha de Generación: {datetime.datetime.now().strftime('%d/%m/%Y')}", styles['Normal']))
        story.append(Spacer(1, 12))

        # Recorrer los textos de análisis A, B, C y D
        full_analysis_text = analisis_A_text + "\n" + analisis_B_text + "\n" + analisis_C_text + "\n" + analisis_D_text
        
        # Dividir el texto en secciones para manejar la tabla D1
        sections = full_analysis_text.split("D1. MATRIZ DE RATIOS COMPARATIVOS")
        
        # Procesar las secciones A, B y C
        for line in sections[0].split('\n'):
            line = line.strip()
            if '***' in line:
                story.append(Paragraph(line.replace('***', ''), styles['Titulo1']))
            elif line.startswith('A') or line.startswith('B') or line.startswith('C'):
                story.append(Paragraph(line, styles['Subtitulo']))
            elif line:
                story.append(Paragraph(line.replace("*", ""), styles['Normal'])) 
        story.append(Spacer(1, 24))

        # Procesar D1: MATRIZ DE RATIOS COMPARATIVOS (TABLA)
        story.append(Paragraph("🌟 SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO", styles['Titulo1']))
        story.append(Spacer(1, 12))
        story.append(Paragraph("D1. MATRIZ DE RATIOS COMPARATIVOS", styles['Subtitulo']))
        
        table_data = get_matrix_data_for_table(r)
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ])
        
        # Se definen los anchos usando 'inch' para control
        table = Table(table_data, colWidths=[1.5*inch, 0.8*inch, 0.8*inch, 1*inch, 2.5*inch])
        table.setStyle(table_style)
        story.append(table)
        story.append(Spacer(1, 12))

        # Procesar el resto de la sección D (D2, D3, D4)
        d_sections_rest = sections[1]
        for line in d_sections_rest.split('\n'):
            line = line.strip()
            if 'D2' in line or 'D3' in line or 'D4' in line:
                story.append(Paragraph(line.strip().replace(":", ""), styles['Subtitulo']))
            elif line:
                clean_line = line.replace("✅", "->").replace("⚠️", "->").replace("🌟", "").replace("*", "")
                story.append(Paragraph(clean_line, styles['Normal']))

        # 5. Generar el PDF
        doc.build(story)
        messagebox.showinfo("Generación de PDF", f"¡Informe '{filename}' generado con éxito!")

    except Exception as e:
        messagebox.showerror("Error al generar PDF", f"Ocurrió un error al generar el PDF: {e}")


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
add_field(f2023, "Activo Corriente 2023 (AC)", "AC_2023", 25000.00)
add_field(f2023, "Activo No Corriente 2023 (ANC)", "ANC_2023", 10000.00)
add_field(f2023, "Pasivo Corriente 2023 (PC)", "PC_2023", 10000.00)
add_field(f2023, "Patrimonio Neto 2023 (PN)", "PN_2023", 15000.00)
add_field(f2023, "Pasivo No Corriente 2023 (PNC)", "PNC_2023", 10000.00)
add_field(f2023, "Caja y Bancos 2023", "Caja_2023", 1000.00)
add_field(f2023, "Clientes por cobrar 2023", "Clientes_2023", 1500.00)
add_field(f2023, "Inversiones CP 2023 (Inv)", "InvCP_2023", 100.00)
add_field(f2023, "Ingresos 2023", "Ingresos_2023", 85000.00)
add_field(f2023, "BAII 2023 (Aprox.)", "BAII_2023", 12000.00)
add_field(f2023, "Utilidad Neta 2023 (Aprox.)", "UN_2023", 8000.00)

f2024 = ttk.Frame(sections); sections.add(f2024, text="Balance 2024")
add_field(f2024, "Activo Corriente 2024 (AC)", "AC_2024", 28000.00)
add_field(f2024, "Activo No Corriente 2024 (ANC)", "ANC_2024", 12000.00)
add_field(f2024, "Pasivo Corriente 2024 (PC)", "PC_2024", 12000.00)
add_field(f2024, "Pasivo No Corriente 2024 (PNC)", "PNC_2024", 11000.00)
add_field(f2024, "Patrimonio Neto 2024 (PN)", "PN_2024", 17000.00)
add_field(f2024, "Caja y Bancos 2024", "Caja_2024", 1100.00)
add_field(f2024, "Clientes por cobrar 2024", "Clientes_2024", 1600.00)
add_field(f2024, "Inversiones CP 2024 (Inv)", "InvCP_2024", 150.00)

fer = ttk.Frame(sections); sections.add(fer, text="Estado de Resultados 2024")
add_field(fer, "Ingresos 2024", "Ingresos_2024", 112000.00)
add_field(fer, "Costo Servicios 2024", "Costo_2024", 41000.00)
add_field(fer, "Ganancia Bruta 2024", "GB_2024", 71000.00)
add_field(fer, "Gastos Administrativos 2024 (GA)", "GA_2024", 26000.00)
add_field(fer, "Gastos de Ventas 2024 (GV)", "GV_2024", 14000.00)
add_field(fer, "BAII 2024", "BAII_2024", 31000.00)
add_field(fer, "Gastos Financieros 2024", "GastosFin_2024", 2000.00)
add_field(fer, "Utilidad Neta 2024", "UN_2024", 20000.00)

fprod = ttk.Frame(sections); sections.add(fprod, text="Ciclo Efectivo")
add_field(fprod, "Días Inventario (DI)", "DI", 45)
add_field(fprod, "Días Clientes (DC)", "DC", 60)
add_field(fprod, "Días Proveedores (DP)", "DP", 30)
# --- Fin de la creación de formularios ---

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
    
    text = "🏆 *** SECCIÓN A: ANÁLISIS PATRIMONIAL ***\n\n"
    
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
    
    text = "💰 *** SECCIÓN B: ANÁLISIS FINANCIERO ***\n\n"
    
    # B1. Ratios de Liquidez 2024
    text += "B1. RATIOS DE LIQUIDEZ 2024 \n"
    text += f"a) Razón de liquidez general (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n"
    text += f"b) Razón de tesorería (Disp+Deud/PC): **{r['T_2024']:.2f}** (Óptimo 0.7-1.0)\n"
    text += f"c) Razón de disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo 0.2-0.3)\n"
    
    liquidez_comentario = f"La Razón General ({r['LG_2024']:.2f}) y la Razón de Tesorería ({r['T_2024']:.2f}) están **muy por encima del óptimo**. Esto indica un **EXCESO DE LIQUIDEZ** y un capital de trabajo mal gestionado, lo que se traduce en **activos corrientes improductivos** (dinero sin invertir)."
    if r['LG_2024'] < 1.0 or r['T_2024'] < 1.0:
        liquidez_comentario = "La empresa presenta problemas de liquidez y enfrenta un riesgo inminente de suspensión de pagos."
    text += f"Diagnóstico: {liquidez_comentario}\n\n"
    
    # B2. Ratios de Solvencia 2024
    text += "B2. RATIOS DE SOLVENCIA 2024 \n"
    text += f"a) Ratio de garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}** (Óptimo 1.5-2.5)\n"
    text += f"b) Ratio de autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n"
    text += f"c) Ratio de calidad de deuda (PC/Pasivo): **{r['calidad_2024']:.2f}**\n"
    
    deuda_comentario = f"El Ratio de Garantía ({r['garantia_2024']:.2f}) es **sólido** y garantiza la cobertura total de las obligaciones. La empresa presenta una **alta autonomía** ({r['autonomia_2024']:.2f}), lo que reduce el riesgo financiero a largo plazo."
    if r['garantia_2024'] < 1.0:
        deuda_comentario = "La empresa está sobre-endeudada (Ratio de Garantía < 1.0) y enfrenta un riesgo de concurso de acreedores."
    text += f"Diagnóstico: {deuda_comentario}\n\n"
    
    # B3. Comparativa 2023 vs 2024
    text += "B3. COMPARATIVA 2023 VS 2024 - Explique por qué.\n"
    text += f"* Liquidez General: {r['LG_2023']:.2f} (2023) -> {r['LG_2024']:.2f} (2024). **{'EMPEORÓ' if r['LG_2024'] < r['LG_2023'] else 'MEJORÓ'}**.\n"
    text += f"* Razón de Tesorería: {r['T_2023']:.2f} (2023) -> {r['T_2024']:.2f} (2024). **{'EMPEORÓ' if r['T_2024'] < r['T_2023'] else 'MEJORÓ'}**.\n"
    text += f"* Ratio Garantía: {r['garantia_2023']:.2f} (2023) -> {r['garantia_2024']:.2f} (2024). **{'EMPEORÓ' if r['garantia_2024'] < r['garantia_2023'] else 'MEJORÓ'}**.\n"
    
    explicacion_b3 = f"Explicación: Aunque los ratios de liquidez disminuyeron (Empeoró), el nivel actual ({r['LG_2024']:.2f}) representa un nivel **más eficiente** del capital de trabajo, acercándose al rango óptimo (1.5-2.0). La reducción se explica por un crecimiento proporcionalmente mayor del Pasivo Corriente ({r.get('h_PC', 0.0):.2f}%) en relación al Activo Corriente ({r.get('h_AC', 0.0):.2f}%).\n\n"
    text += explicacion_b3

    # B4. Análisis de Estructura Financiera
    text += "B4. ANÁLISIS DE ESTRUCTURA FINANCIERA \n"
    text += f"* % de deuda a corto plazo (PC/Pasivo): **{r['pct_PC']:.2f}%**\n"
    text += f"* % de deuda a largo plazo (PNC/Pasivo): **{r['pct_PNC']:.2f}%**\n"
    text += f"* % de recursos propios (PN/Total Fin.): **{r['pct_PN_fin']:.2f}%**\n"
    
    conclusion_b4 = f"Conclusión: El {r['pct_PC']:.2f}% de la deuda total es a corto plazo, lo cual es manejable, pero indica una dependencia de financiación a corto plazo que presiona el capital de trabajo. La estructura es **muy sólida** por el alto porcentaje de Recursos Propios ({r['pct_PN_fin']:.2f}%).\n\n"
    text += conclusion_b4

    # B5. Estrés Financiero - Escenario Pesimista (CORRECCIÓN PQ)
    text += "B5. ESTRÉS FINANCIERO - ESCENARIO PESIMISTA \n"
    pq_value = r['PQ']
    text += f"Proyección 2025 (Simulación: Ingresos -30%): **{r['Ingresos_2025_sim']:.2f} Bs.**\n"
    text += f"a) FM (simulado): **{r['FM_2025_sim']:.2f} Bs.** \n"
    text += f"b) Liquidez General (simulada): **{r['LG_2025_sim']:.2f}**\n"
    text += f"c) Punto de quiebra (mínimo ingreso requerido, **CALCULADO**): **{pq_value:.2f} Bs.**\n"

    if r['Ingresos_2025_sim'] < pq_value:
        diagnostico_estres = f"El **Punto de Quiebre ({pq_value:.2f} Bs.)** es superior a las ventas simuladas de **{r['Ingresos_2025_sim']:.2f} Bs.**\n**Conclusión:** Esto indica que la empresa **operaría con PÉRDIDAS** ({r['UN_2025_sim']:.2f} Bs.) en este escenario. Aunque el FM es positivo, la caída de las ventas pone en riesgo la **solvencia operativa** a corto plazo."
    else:
        diagnostico_estres = "Las ventas simuladas son superiores al Punto de Quiebre, manteniendo la rentabilidad a pesar de la caída."

    text += f"Diagnóstico: {diagnostico_estres}\n"

    outB.insert(tk.END, text)
    analisis_B_text = text # Almacenar para PDF


def generar_analisis_economico(r, data):
    """Genera el texto de análisis C."""
    global analisis_C_text
    outC.delete("1.0", tk.END)
    
    text = "📈 *** SECCIÓN C: ANÁLISIS ECONÓMICO - RENTABILIDAD ***\n\n"
    
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
    text += "C5. APALANCAMIENTO FINANCIERO  - **FÓRMULA ESTÁNDAR**\n"
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

    text = "🌟 *** SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO  ***\n\n"

    # D1. Matriz de Ratios Comparativos (Se utiliza tabla Markdown para la GUI, pero ReportLab usará la función get_matrix_data_for_table)
    text += "D1. MATRIZ DE RATIOS COMPARATIVOS \n"
    
    matriz = [
        ("FM (Bs)", r['FM_2023'], r['FM_2024'], "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", r['LG_2023'], r['LG_2024'], "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
        ("Tesorería", r['T_2023'], r['T_2024'], "Empeoró" if r['T_2024'] < r['T_2023'] else "Mejora", "Continúa siendo excesiva (activos ociosos)."),
        ("RAT (%)", r['RAT_2023'], r['RAT_2024'], "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", r['RRP_2023'], r['RRP_2024'], "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
    ]
    
    # Generación de la tabla Markdown para la GUI (ScrolledText)
    text += "| Ratio | 2023 | 2024 | Cambio | Interpretación |\n"
    text += "|---|---|---|---|---|\n"
    for ratio, y23, y24, cambio, inter in matriz:
        text += f"| {ratio:<12} | {y23:.2f} | {y24:.2f} | {cambio:<7} | {inter} |\n"
    text += "\n"
    
    # D2. Fortalezas y Debilidades
    text += "D2. FORTALEZAS Y DEBILIDADES \n"
    text += "✅ FORTALEZAS (Cuantificadas):\n"
    text += f"- **Rentabilidad Sólida:** La RRP en 2024 fue del **{r['RRP_2024']:.2f}%**, superior a la RAT. (C2)\n"
    text += f"- **Autonomía Financiera:** Alto Ratio de Autonomía (PN/Pasivo = **{r['autonomia_2024']:.2f}**), lo que implica bajo riesgo financiero. (B2)\n"
    text += f"- **Fondo de Maniobra:** FM positivo de **{r['FM_2024']:.2f} Bs.**, asegurando equilibrio patrimonial. (A1)\n"
    text += f"- **Apalancamiento Favorable:** RAT ({r['RAT_2024']:.2f}%) es mayor que el Costo de Deuda ({r['costo_deuda']:.2f}%). (C5)\n"
    text += "\n⚠️ DEBILIDADES (Cuantificadas):\n"
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
    """Función unificada para calcular y generar todos los análisis."""
    analizador = AnalisisFinanciero(data)
    r = analizador.calcular_todo()
    
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
    ax1.pie(sizes_a, labels=labels_a, autopct='%1.1f%%', startangle=90)
    ax1.set_title('Estructura del Activo 2024', fontsize=10)
    
    # Subplot 2: Evolución de Ratios (LG, RAT)
    ax2 = fig.add_subplot(222)
    labels_r = ['2023', '2024']
    lg_vals = [r['LG_2023'], r['LG_2024']]
    rat_vals = [r['RAT_2023']/100, r['RAT_2024']/100] # Dividido por 100 para escala
    x = np.arange(len(labels_r))
    width = 0.35
    
    rects1 = ax2.bar(x - width/2, lg_vals, width, label='Liq. General')
    rects2 = ax2.bar(x + width/2, rat_vals, width, label='RAT (ROA)')
    
    ax2.set_ylabel('Ratio/Rentabilidad')
    ax2.set_title('Evolución de Ratios Clave', fontsize=10)
    ax2.set_xticks(x)
    ax2.set_xticklabels(labels_r)
    ax2.legend(fontsize=8)
    
    # Subplot 3: Descomposición DuPont (RRP)
    ax3 = fig.add_subplot(212)
    labels_d = ['Margen Neto', 'Rotación Activo', 'Apalancamiento']
    dupont_vals = [r['margen_neto_dupont'], r['rotacion_activo'], r['apalancamiento_dupont']]
    ax3.bar(labels_d, dupont_vals)
    ax3.set_title('Componentes DuPont RRP 2024', fontsize=10)
    
    fig.tight_layout(pad=3.0) # Asegura que los gráficos no se superpongan
    
    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True)
    canvas.draw()
    
    # ------------------------------------------------------------------
    # Botón para generar PDF (Se crea dinámicamente aquí)
    # ------------------------------------------------------------------
    # Eliminar botones viejos si existen para evitar duplicados
    for w in tabD.winfo_children():
        if isinstance(w, ttk.Button) and w.cget("text") == "Generar Informe PDF":
            w.destroy()
            
    # Botón de PDF: llama a la función corregida
    ttk.Button(tabD, text="Generar Informe PDF", 
               command=lambda: generate_pdf(r)
              ).pack(side="bottom", pady=10)

# Botón principal para ejecutar el análisis (SIN CAMBIOS)
button_analyze = ttk.Button(tab_inputs, text="Ejecutar Análisis y Generar Reporte", 
                             command=lambda: run_analisis(read_all_inputs(form)))
button_analyze.pack(pady=20, padx=10, fill='x')


root.mainloop()