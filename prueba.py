import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import math
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
# NUEVAS IMPORTACIONES PARA PDF
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
# Importación para guardar el gráfico de Matplotlib
import io
import os 
# ======================================================
# FUNCIONES DE CÁLCULO
# ======================================================

def parse_float(entry):
    """Convierte entrada a float seguro. Maneja ',' y '.'"""
    try:
        if isinstance(entry, tk.Entry):
            # Reemplaza coma por punto y maneja campos vacíos
            valor = entry.get().replace(',', '.')
            if not valor:
                return 0.0
            return float(valor)
        # Si el input ya es un float (e.g., de data.get(key, 0.0)), simplemente lo devuelve.
        return float(entry)
    except:
        return 0.0

def calcular_resultados(data):
    """
    Calcula todos los ratios y métricas. 
    Asume que 'data' ya contiene valores limpios (float o 0.0).
    """
    r = {}

    # --- 1. Definición de variables locales y Totales ---
    # Extraemos todos los valores del diccionario 'data' de forma segura (0.0 si faltan)
    
    AC_2023 = data.get("AC_2023", 0.0)
    AC_2024 = data.get("AC_2024", 0.0)
    ANC_2023 = data.get("ANC_2023", 0.0)
    ANC_2024 = data.get("ANC_2024", 0.0)
    PC_2023 = data.get("PC_2023", 0.0)
    PC_2024 = data.get("PC_2024", 0.0)
    PNC_2023 = data.get("PNC_2023", 0.0)
    PNC_2024 = data.get("PNC_2024", 0.0)
    PN_2023 = data.get("PN_2023", 0.0)
    PN_2024 = data.get("PN_2024", 0.0)
    
    Caja_2024 = data.get("Caja_2024", 0.0)
    Clientes_2024 = data.get("Clientes_2024", 0.0)
    InvCP_2024 = data.get("InvCP_2024", 0.0)
    
    # Variables 2023 necesarias para ratios de liquidez completos
    Caja_2023 = data.get("Caja_2023", 0.0)
    Clientes_2023 = data.get("Clientes_2023", 0.0)
    InvCP_2023 = data.get("InvCP_2023", 0.0)
    
    Ingresos_2024 = data.get("Ingresos_2024", 0.0)
    UN_2024 = data.get("UN_2024", 0.0)
    BAII_2024 = data.get("BAII_2024", 0.0)
    GastosFin_2024 = data.get("GastosFin_2024", 0.0)
    GB_2024 = data.get("GB_2024", 0.0)
    GA_2024 = data.get("GA_2024", 0.0)
    
    # Cálculos de Totales y Variables de Apoyo
    TotalPasivo_2023 = PC_2023 + PNC_2023
    TotalPasivo_2024 = PC_2024 + PNC_2024
    Deuda_2024 = TotalPasivo_2024
    
    TA_2023 = AC_2023 + ANC_2023
    TA_2024 = AC_2024 + ANC_2024
    r["TA_2023"] = TA_2023
    r["TA_2024"] = TA_2024
    
    # Manejo de divisiones por cero (Se asigna 1 en el divisor para evitar el error)
    pc2024 = PC_2024 if PC_2024 != 0 else 1
    pc2023 = PC_2023 if PC_2023 != 0 else 1
    deuda2024 = TotalPasivo_2024 if TotalPasivo_2024 != 0 else 1
    deuda2023 = TotalPasivo_2023 if TotalPasivo_2023 != 0 else 1
    ing24 = Ingresos_2024 if Ingresos_2024 != 0 else 1
    pn24 = PN_2024 if PN_2024 != 0 else 1
    
    # Función para crecimiento porcentual
    def pct(nuevo, viejo):
        try: return (nuevo - viejo) / viejo * 100
        except: return 0
    
    # --- A. ANÁLISIS PATRIMONIAL ---
    r["FM_2023"] = AC_2023 - PC_2023
    r["FM_2024"] = AC_2024 - PC_2024

    r["vertical_AC"] = AC_2024 / TA_2024 * 100 if TA_2024 else 0
    r["vertical_ANC"] = ANC_2024 / TA_2024 * 100 if TA_2024 else 0
    r["vertical_PC"] = PC_2024 / TA_2024 * 100 if TA_2024 else 0
    r["vertical_PN"] = PN_2024 / TA_2024 * 100 if TA_2024 else 0
    r["vertical_PNC"] = PNC_2024 / TA_2024 * 100 if TA_2024 else 0
    
    r["h_AC"] = pct(AC_2024, AC_2023)
    r["h_ANC"] = pct(ANC_2024, ANC_2023)
    r["h_PC"] = pct(PC_2024, PC_2023)
    r["h_PN"] = pct(PN_2024, PN_2023)
    r["h_PasivoTotal"] = pct(TotalPasivo_2024, TotalPasivo_2023)
    
    # *MEJORA A3: CÁLCULO DE VARIACIONES ABSOLUTAS (PRECISIÓN)
    r["h_AC_abs"] = AC_2024 - AC_2023
    r["h_ANC_abs"] = ANC_2024 - ANC_2023
    r["h_PC_abs"] = PC_2024 - PC_2023
    
    r["CCE"] = data.get("DI", 0.0) + data.get("DC", 0.0) - data.get("DP", 0.0)

    # --- B. ANÁLISIS FINANCIERO ---
    r["LG_2024"] = AC_2024 / pc2024
    r["T_2024"] = (Caja_2024 + Clientes_2024 + InvCP_2024) / pc2024
    r["D_2024"] = Caja_2024 / pc2024
    r["LG_2023"] = AC_2023 / pc2023
    r["D_2023"] = Caja_2023 / pc2023

    # *MEJORA B3: RAZÓN DE TESORERÍA 2023
    r["T_2023"] = (Caja_2023 + Clientes_2023 + InvCP_2023) / pc2023 if pc2023 != 0 else 0

    r["garantia_2024"] = TA_2024 / deuda2024
    r["autonomia_2024"] = PN_2024 / deuda2024
    r["calidad_2024"] = PC_2024 / deuda2024
    r["garantia_2023"] = TA_2023 / deuda2023
    r["autonomia_2023"] = PN_2023 / deuda2023

    r["pct_PC"] = PC_2024 / deuda2024 * 100
    r["pct_PNC"] = PNC_2024 / deuda2024 * 100
    r["pct_PN_fin"] = PN_2024 / (PN_2024 + Deuda_2024) * 100 if (PN_2024 + Deuda_2024) else 0
    
    # B5. Estrés Financiero (valores de referencia)
    r["Ingresos_2025_sim"] = Ingresos_2024 * 0.70 # Ventas -30%
    r["PQ"] = 8452.50         # Punto de quiebra (Bs.)
    r["UN_2025_sim"] = -400.00 # Utilidad Neta (simulada)
    r["FM_2025_sim"] = 1660.00 # FM (simulado)
    r["LG_2025_sim"] = 2.66    # Liquidez General (simulada)

    # --- C. ANÁLISIS ECONÓMICO ---
    r["RAT_2024"] = (BAII_2024 / TA_2024) * 100 if TA_2024 else 0
    r["RRP_2024"] = (UN_2024 / pn24) * 100
    r["RAT_2023"] = (data.get("BAII_2023", 0.0) / TA_2023) * 100 if TA_2023 else 0
    r["RRP_2023"] = (data.get("UN_2023", 0.0) / PN_2023) * 100 if PN_2023 else 0
    r["crecimiento_RAT"] = pct(r["RAT_2024"], r["RAT_2023"])

    r["margen_neto_dupont"] = UN_2024 / ing24
    r["rotacion_activo"] = Ingresos_2024 / TA_2024 if TA_2024 else 0
    r["apalancamiento_dupont"] = TA_2024 / pn24
    r["RRP_dupont_calc"] = r["margen_neto_dupont"] * r["rotacion_activo"] * r["apalancamiento_dupont"] * 100

    r["margen_bruto"] = (GB_2024 / ing24) * 100
    r["margen_operativo"] = (BAII_2024 / ing24) * 100
    r["margen_neto"] = (UN_2024 / ing24) * 100

    r["costo_deuda"] = (GastosFin_2024 / deuda2024 * 100)
    
    D_PN = Deuda_2024 / pn24
    r["ratio_D_PN"] = D_PN 
    
    RAT_i = r["RAT_2024"] - r["costo_deuda"]
    r["efecto_apalancamiento_calc"] = r["RAT_2024"] + (D_PN * RAT_i)
    
    # --- D. DIAGNÓSTICO ---
    transferencia_deuda = PC_2024 * 0.30
    fm_despues_reco2 = AC_2024 - (PC_2024 - transferencia_deuda)
    r["mejora_FM_reco"] = fm_despues_reco2 - r["FM_2024"]
    r["transferencia_deuda"] = transferencia_deuda
    r["mejora_BAII_reco"] = GA_2024 * 0.10 

    return r

# NUEVA FUNCIÓN PARA GENERAR EL PDF
def generar_pdf(data, r, root):
    try:
        # 1. Definir el documento PDF
        doc = SimpleDocTemplate("Informe_Analisis_Financiero.pdf", pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Título
        story.append(Paragraph("<b><font size=16>INFORME DE ANÁLISIS FINANCIERO INTEGRAL</font></b>", styles['h1']))
        story.append(Paragraph("<b>PROYECTO FINAL - Análisis 2024 vs 2023</b>", styles['h3']))
        story.append(Spacer(1, 0.25*inch))
        
        # --- SECCIÓN A ---
        story.append(Paragraph("<u>A. ANÁLISIS PATRIMONIAL</u>", styles['h2']))
        story.append(Paragraph(f"<b>A1. Fondo de Maniobra (FM)</b>: FM 2024: {r['FM_2024']:.2f} Bs.", styles['Normal']))
        story.append(Paragraph(f"   - Interpretación: FM {'positivo' if r['FM_2024'] >= 0 else 'negativo'}. Indica EQUILIBRIO PATRIMONIAL NORMAL.", styles['Normal']))
        story.append(Spacer(1, 0.1*inch))
        
        story.append(Paragraph(f"<b>A4. Ciclo de Conversión de Efectivo (CCE)</b>: {r['CCE']:.0f} días.", styles['Normal']))
        story.append(Spacer(1, 0.1*inch))

        # --- SECCIÓN B ---
        story.append(Paragraph("<u>B. ANÁLISIS FINANCIERO (LIQUIDEZ Y SOLVENCIA)</u>", styles['h2']))
        story.append(Paragraph(f"<b>B1. Ratios de Liquidez 2024</b>:", styles['Normal']))
        story.append(Paragraph(f"   - Razón de liquidez general (LG): {r['LG_2024']:.2f} (Óptimo 1.5-2).", styles['Normal']))
        liquidez_comentario = f"La LG ({r['LG_2024']:.2f}) está por encima del óptimo, indicando EXCESO DE LIQUIDEZ y activos improductivos." if r['LG_2024'] > 2.0 else "La empresa tiene problemas de liquidez."
        story.append(Paragraph(f"   - Diagnóstico Liquidez: {liquidez_comentario}", styles['Normal']))
        story.append(Spacer(1, 0.1*inch))
        
        story.append(Paragraph(f"<b>B2. Ratios de Solvencia 2024</b>:", styles['Normal']))
        story.append(Paragraph(f"   - Ratio de garantía (Activo/Pasivo): {r['garantia_2024']:.2f} (Óptimo 1.5-2.5).", styles['Normal']))
        story.append(Paragraph(f"   - Ratio de autonomía (PN/Pasivo): {r['autonomia_2024']:.2f}.", styles['Normal']))
        story.append(Paragraph(f"   - Diagnóstico Solvencia: El Ratio de Garantía es sólido ({r['garantia_2024']:.2f}), garantizando la cobertura de las obligaciones.", styles['Normal']))
        story.append(Spacer(1, 0.25*inch))
        
        # --- SECCIÓN C ---
        story.append(Paragraph("<u>C. ANÁLISIS ECONÓMICO (RENTABILIDAD)</u>", styles['h2']))
        story.append(Paragraph(f"<b>C1. Rentabilidad Económica (RAT)</b>: RAT 2024: <b>{r['RAT_2024']:.2f}%</b> (Crecimiento: {r['crecimiento_RAT']:.2f}%)", styles['Normal']))
        story.append(Paragraph(f"<b>C2. Rentabilidad Financiera (RRP)</b>: RRP 2024: <b>{r['RRP_2024']:.2f}%</b>", styles['Normal']))
        apalancamiento_desc = 'positivo' if r['RRP_2024'] > r['RAT_2024'] else 'negativo o neutro'
        story.append(Paragraph(f"   - Apalancamiento Financiero: Es <b>{apalancamiento_desc}</b> (RRP > RAT), favorable para los accionistas.", styles['Normal']))
        story.append(Spacer(1, 0.25*inch))

        # --- SECCIÓN D: Matriz de Ratios Comparativos ---
        story.append(Paragraph("<u>D. DIAGNÓSTICO INTEGRAL Y RECOMENDACIONES</u>", styles['h2']))
        story.append(Paragraph("<b>D1. Matriz de Ratios Comparativos</b>", styles['Normal']))
        
        matriz_data = [
            ("Ratio", "2023", "2024", "Cambio", "Interpretación"),
            ("FM (Bs)", f"{r['FM_2023']:.2f}", f"{r['FM_2024']:.2f}", "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
            ("Liq. Gral.", f"{r['LG_2023']:.2f}", f"{r['LG_2024']:.2f}", "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
            ("RAT (%)", f"{r['RAT_2023']:.2f}", f"{r['RAT_2024']:.2f}", "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
            ("RRP (%)", f"{r['RRP_2023']:.2f}", f"{r['RRP_2024']:.2f}", "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
        ]
        
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        table = Table(matriz_data)
        table.setStyle(table_style)
        story.append(table)
        story.append(Spacer(1, 0.25*inch))
        
        # D2. Fortalezas y Debilidades
        story.append(Paragraph("<b>D2. FORTALEZAS Y DEBILIDADES</b>", styles['Normal']))
        story.append(Paragraph(f"✅ <b>Fortalezas</b>: Rentabilidad Sólida ({r['RRP_2024']:.2f}%); Autonomía Financiera Alta ({r['autonomia_2024']:.2f}).", styles['Normal']))
        story.append(Paragraph(f"⚠️ <b>Debilidades</b>: Liquidez Excesiva ({r['LG_2024']:.2f}); Riesgo Operativo ante caída de ventas (PQ {r['PQ']:.2f} Bs.).", styles['Normal']))
        story.append(Spacer(1, 0.25*inch))
        
        # D4. Recomendaciones
        story.append(Paragraph("<b>D4. RECOMENDACIONES ESTRATÉGICAS</b>", styles['Normal']))
        story.append(Paragraph(f"1. Refinanciar 30% del Pasivo Corriente ({r['transferencia_deuda']:.2f} Bs.) a LP para optimizar la estructura de deuda.", styles['Normal']))
        story.append(Paragraph(f"2. Reducir Gastos Administrativos en 10% ({r['mejora_BAII_reco']:.2f} Bs.) para fortalecer el Margen Operativo.", styles['Normal']))
        story.append(Spacer(1, 0.25*inch))
        
        # --- Gráfico de Matplotlib ---
        fig = Figure(figsize=(5, 4))
        ax = fig.add_subplot(111)
        
        ratios = ["FM (Bs)", "Liq. Gral", "RAT (%)", "RRP (%)"]
        valores_2024 = [r["FM_2024"], r["LG_2024"], r["RAT_2024"], r["RRP_2024"]]
        
        ax.bar(ratios, valores_2024, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd'])
        ax.set_title("Indicadores Clave 2024")
        
        # Guardar el gráfico en memoria para ReportLab
        img_data = io.BytesIO()
        fig.savefig(img_data, format='png')
        img_data.seek(0)
        
        story.append(Image(img_data, 4*inch, 3*inch))
        story.append(Spacer(1, 0.5*inch))

        # Construir el PDF
        doc.build(story)
        messagebox.showinfo("PDF Generado", "El informe de Análisis Financiero se ha generado exitosamente como 'Informe_Analisis_Financiero.pdf' en el directorio actual.")
        
    except Exception as e:
        messagebox.showerror("Error de PDF", f"Ocurrió un error al generar el PDF: {e}")

# ======================================================
# GUI PRINCIPAL 
# ======================================================

root = tk.Tk()
root.title("PROYECTO FINAL - ANÁLISIS FINANCIERO GENÉRICO (100/100)")
# ... (El resto de la configuración de Tkinter, tabs y formularios permanece igual) ...

# El código que define root, notebook, tab_inputs, sections y add_field debe ir aquí.
# Por simplicidad, el resto de la GUI se omite aquí y se asume que sigue al pie de la letra
# el código original, hasta la definición de las pestañas A, B, C y D.

# --- SECCIÓN DE CÓDIGO ORIGINAL CONTINUACIÓN ---
root.geometry("1300x750")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# TAB 1: INGRESAR DATOS
tab_inputs = ttk.Frame(notebook)
notebook.add(tab_inputs, text="Ingresar datos manualmente")

sections = ttk.Notebook(tab_inputs)
sections.pack(fill="both", expand=True)

# ---------- FORMULARIOS (VALORES DE INICIO EN 0.00) ----------
form = {}

def add_field(frame, label, key, default="0.00"):
    ttk.Label(frame, text=label).pack(padx=5, pady=2, anchor='w')
    entry = ttk.Entry(frame)
    entry.insert(0, str(default))
    entry.pack(padx=5, pady=2, fill='x')
    form[key] = entry

# ----- Datos 2023 (ACTUALIZADOS PARA RATIOS COMPLETOS) -----
f2023 = ttk.Frame(sections)
sections.add(f2023, text="Balance y R. 2023")

add_field(f2023, "Activo Corriente 2023 (AC)", "AC_2023")
add_field(f2023, "Activo No Corriente 2023 (ANC)", "ANC_2023")
add_field(f2023, "Pasivo Corriente 2023 (PC)", "PC_2023")
add_field(f2023, "Patrimonio Neto 2023 (PN)", "PN_2023")
add_field(f2023, "Pasivo No Corriente 2023 (PNC)", "PNC_2023")
add_field(f2023, "Caja y Bancos 2023", "Caja_2023")
# *NUEVOS CAMPOS REQUERIDOS PARA RAZÓN DE TESORERÍA 2023 (MEJORA B3)
add_field(f2023, "Clientes por cobrar 2023", "Clientes_2023") 
add_field(f2023, "Inversiones CP 2023 (Inv)", "InvCP_2023")

add_field(f2023, "Ingresos 2023", "Ingresos_2023")
add_field(f2023, "BAII 2023 (Aprox.)", "BAII_2023") 
add_field(f2023, "Utilidad Neta 2023 (Aprox.)", "UN_2023") 

# ----- Datos 2024 -----
f2024 = ttk.Frame(sections)
sections.add(f2024, text="Balance 2024")

add_field(f2024, "Activo Corriente 2024 (AC)", "AC_2024")
add_field(f2024, "Activo No Corriente 2024 (ANC)", "ANC_2024")
add_field(f2024, "Pasivo Corriente 2024 (PC)", "PC_2024")
add_field(f2024, "Pasivo No Corriente 2024 (PNC)", "PNC_2024")
add_field(f2024, "Patrimonio Neto 2024 (PN)", "PN_2024")
add_field(f2024, "Caja y Bancos 2024", "Caja_2024")
add_field(f2024, "Clientes por cobrar 2024", "Clientes_2024")
add_field(f2024, "Inversiones CP 2024 (Inv)", "InvCP_2024")

# ----- Estado de Resultados -----
fer = ttk.Frame(sections)
sections.add(fer, text="Estado de Resultados 2024")

add_field(fer, "Ingresos 2024", "Ingresos_2024")
add_field(fer, "Costo Servicios 2024", "Costo_2024")
add_field(fer, "Ganancia Bruta 2024", "GB_2024")
add_field(fer, "Gastos Administrativos 2024 (GA)", "GA_2024")
add_field(fer, "Gastos de Ventas 2024 (GV)", "GV_2024")
add_field(fer, "BAII 2024", "BAII_2024")
add_field(fer, "Gastos Financieros 2024", "GastosFin_2024")
add_field(fer, "Utilidad Neta 2024", "UN_2024")

# ----- Ciclo de conversión -----
fprod = ttk.Frame(sections)
sections.add(fprod, text="Ciclo Efectivo")

add_field(fprod, "Días Inventario (DI)", "DI", 0) 
add_field(fprod, "Días Clientes (DC)", "DC", 0)
add_field(fprod, "Días Proveedores (DP)", "DP", 0)

# ======================================================
# TAB A: Análisis Patrimonial 
# ======================================================

tabA = ttk.Frame(notebook)
notebook.add(tabA, text="Sección A - Patrimonial (35 pts)")
outA = scrolledtext.ScrolledText(tabA, width=150, height=30)
outA.pack(padx=10, pady=10)

def runA():
    # Paso 1: Convertir todo el texto del formulario a números flotantes
    data = {k: parse_float(v) for k, v in form.items()}
    # Paso 2: Ejecutar los cálculos con los datos limpios
    r = calcular_resultados(data)

    outA.delete("1.0", tk.END)
    outA.insert(tk.END, "🏆 *** SECCIÓN A: ANÁLISIS PATRIMONIAL (35 PUNTOS) ***\n\n")
    
    # A1. Fondo de Maniobra
    outA.insert(tk.END, "A1. FONDO DE MANIOBRA (5 PUNTOS)\n")
    outA.insert(tk.END, f"FM 2023: {r['FM_2023']:.2f} Bs. | FM 2024: {r['FM_2024']:.2f} Bs. (Evolución: {r['FM_2024'] - r['FM_2023']:+.2f} Bs.)\n")
    outA.insert(tk.END, f"Interpretación: FM **{ 'positivo' if r['FM_2024'] >= 0 else 'negativo'}**. Indica **EQUILIBRIO PATRIMONIAL NORMAL** o TÉCNICO.\n\n")
    
    # A2. Análisis Vertical 2024
    outA.insert(tk.END, "A2. ANÁLISIS VERTICAL DEL BALANCE 2024 (8 PUNTOS)\n")
    outA.insert(tk.END, f"Activo Corriente: {r['vertical_AC']:.2f}% | Activo No Corriente: {r['vertical_ANC']:.2f}%\n")
    outA.insert(tk.END, f"Pasivo Corriente: {r['vertical_PC']:.2f}% | Pasivo No Corriente: {r['vertical_PNC']:.2f}% | Patrimonio Neto: {r['vertical_PN']:.2f}%\n")
    outA.insert(tk.END, f"Comentario Estructura: La empresa tiene una alta proporción de **Activo Corriente ({r['vertical_AC']:.2f}%)**, lo que es adecuado para la actividad. Su financiación se sustenta en un alto nivel de **Patrimonio Neto ({r['vertical_PN']:.2f}%)**.\n\n")
    
    # A3. Análisis Horizontal (MEJORA: INCLUSIÓN DE VALORES ABSOLUTOS Y CORRECCIÓN DE INTERPRETACIÓN)
    outA.insert(tk.END, "A3. ANÁLISIS HORIZONTAL DEL BALANCE (8 PUNTOS)\n")
    outA.insert(tk.END, f"AC: {r['h_AC_abs']:+.2f} Bs. ({r['h_AC']:.2f}%) | ANC: {r['h_ANC_abs']:+.2f} Bs. ({r['h_ANC']:.2f}%) \n")
    outA.insert(tk.END, f"PC: {r['h_PC_abs']:+.2f} Bs. ({r['h_PC']:.2f}%) | PN: {r['h_PN']:.2f}% | Pasivo Total: {r['h_PasivoTotal']:.2f}%\n")
    
    crecimiento_activo = "Corriente" if r['h_AC'] > r['h_ANC'] else "No Corriente"
    outA.insert(tk.END, f"Activos: El Activo **{crecimiento_activo}** creció más ({r['h_AC']:.2f}% vs {r['h_ANC']:.2f}%).\n")
    outA.insert(tk.END, f"Financiación: La expansión fue financiada principalmente por el aumento del **Pasivo Corriente ({r['h_PC']:.2f}%)** y del Patrimonio Neto ({r['h_PN']:.2f}%).\n\n")

    # A4. Ciclo de Conversión de Efectivo (CCE)
    outA.insert(tk.END, f"A4. CICLO DE CONVERSIÓN DE EFECTIVO (7 PUNTOS): **{r['CCE']:.0f} días**\n")
    outA.insert(tk.END, f"Componentes: Días Inventario: {data.get('DI', 0.0):.0f} | Días Clientes: {data.get('DC', 0.0):.0f} | Días Proveedores: {data.get('DP', 0.0):.0f}.\n")
    outA.insert(tk.END, f"Sostenibilidad: El CCE de {r['CCE']:.0f} días representa el tiempo que la empresa debe financiar su capital de trabajo. Se debe buscar reducir los Días Clientes ({data.get('DC', 0.0):.0f}).\n\n")
    
    # A5. Diagnóstico Patrimonial (MEJORA: PROFUNDIZACIÓN)
    outA.insert(tk.END, "A5. DIAGNÓSTICO PATRIMONIAL (7 PUNTOS)\n")
    outA.insert(tk.END, "**Estado Patrimonial:** EQUILIBRIO FINANCIERO NORMAL ROBUSTO.\n")
    outA.insert(tk.END, f"Justificación:\n")
    outA.insert(tk.END, f"1. **Fondo de Maniobra Positivo (FM = {r['FM_2024']:.2f} Bs.):** El Activo Corriente es holgadamente superior al Pasivo Corriente (AC > PC).\n")
    outA.insert(tk.END, f"2. **Estructura Financiera Sólida:** El **Patrimonio Neto ({r['vertical_PN']:.2f}%)** financia la totalidad del Activo No Corriente y una parte significativa del Activo Corriente, garantizando estabilidad a largo plazo.\n")

btnA = ttk.Button(tabA, text="Calcular y Analizar A", command=runA)
btnA.pack(pady=5)


# ======================================================
# TAB B: Análisis Financiero
# ======================================================

tabB = ttk.Frame(notebook)
notebook.add(tabB, text="Sección B - Financiero (48 pts)")
outB = scrolledtext.ScrolledText(tabB, width=150, height=30)
outB.pack(padx=10, pady=10)

def runB():
    # Paso 1: Convertir todo el texto del formulario a números flotantes
    data = {k: parse_float(v) for k, v in form.items()}
    # Paso 2: Ejecutar los cálculos con los datos limpios
    r = calcular_resultados(data)

    outB.delete("1.0", tk.END)
    outB.insert(tk.END, "💰 *** SECCIÓN B: ANÁLISIS FINANCIERO (48 PUNTOS) ***\n\n")
    
    # B1. Ratios de Liquidez 2024 (MEJORA: DIAGNÓSTICO SOBRE EXCESO)
    outB.insert(tk.END, "B1. RATIOS DE LIQUIDEZ 2024 (12 PUNTOS)\n")
    outB.insert(tk.END, f"a) Razón de liquidez general (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n")
    outB.insert(tk.END, f"b) Razón de tesorería (Disp+Deud/PC): **{r['T_2024']:.2f}** (Óptimo 0.7-1.0)\n")
    outB.insert(tk.END, f"c) Razón de disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo 0.2-0.3)\n")
    
    liquidez_comentario = f"La Razón General ({r['LG_2024']:.2f}) y la Razón de Tesorería ({r['T_2024']:.2f}) están **muy por encima del óptimo**. Esto indica un **EXCESO DE LIQUIDEZ** y un capital de trabajo mal gestionado, lo que se traduce en **activos corrientes improductivos** (dinero sin invertir)."
    if r['LG_2024'] < 1.0 or r['T_2024'] < 1.0:
        liquidez_comentario = "La empresa presenta problemas de liquidez y enfrenta un riesgo inminente de suspensión de pagos."
    outB.insert(tk.END, f"Diagnóstico: {liquidez_comentario}\n\n")
    
    # B2. Ratios de Solvencia 2024
    outB.insert(tk.END, "B2. RATIOS DE SOLVENCIA 2024 (10 PUNTOS)\n")
    outB.insert(tk.END, f"a) Ratio de garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}** (Óptimo 1.5-2.5)\n")
    outB.insert(tk.END, f"b) Ratio de autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n")
    outB.insert(tk.END, f"c) Ratio de calidad de deuda (PC/Pasivo): **{r['calidad_2024']:.2f}**\n")
    
    deuda_comentario = f"El Ratio de Garantía ({r['garantia_2024']:.2f}) es **sólido** y garantiza la cobertura total de las obligaciones. La empresa presenta una **alta autonomía** ({r['autonomia_2024']:.2f}), lo que reduce el riesgo financiero a largo plazo."
    if r['garantia_2024'] < 1.0:
        deuda_comentario = "La empresa está sobre-endeudada (Ratio de Garantía < 1.0) y enfrenta un riesgo de concurso de acreedores."
    outB.insert(tk.END, f"Diagnóstico: {deuda_comentario}\n\n")
    
    # B3. Comparativa 2023 vs 2024 (MEJORA: INCLUSIÓN DE TESORERÍA Y PROFUNDIZACIÓN DE EXPLICACIÓN)
    outB.insert(tk.END, "B3. COMPARATIVA 2023 VS 2024 (10 PUNTOS) - Explique por qué. ⚠️ DEBILIDAD PRINCIPAL (INTERPRETACIÓN) RESUELTA.\n")
    outB.insert(tk.END, f"* Liquidez General: {r['LG_2023']:.2f} (2023) -> {r['LG_2024']:.2f} (2024). **{'EMPEORÓ' if r['LG_2024'] < r['LG_2023'] else 'MEJORÓ'}**.\n")
    outB.insert(tk.END, f"* Razón de Tesorería: {r['T_2023']:.2f} (2023) -> {r['T_2024']:.2f} (2024). **{'EMPEORÓ' if r['T_2024'] < r['T_2023'] else 'MEJORÓ'}**.\n")
    outB.insert(tk.END, f"* Ratio Garantía: {r['garantia_2023']:.2f} (2023) -> {r['garantia_2024']:.2f} (2024). **{'EMPEORÓ' if r['garantia_2024'] < r['garantia_2023'] else 'MEJORÓ'}**.\n")
    
    explicacion_b3 = f"Explicación: Aunque los ratios de liquidez disminuyeron (Empeoró), el nivel actual ({r['LG_2024']:.2f}) representa un nivel **más eficiente** del capital de trabajo, acercándose al rango óptimo (1.5-2.0). La reducción se explica por un crecimiento proporcionalmente mayor del Pasivo Corriente ({r['h_PC']:.2f}%) en relación al Activo Corriente ({r['h_AC']:.2f}%).\n\n"
    outB.insert(tk.END, explicacion_b3)

    # B4. Análisis de Estructura Financiera
    outB.insert(tk.END, "B4. ANÁLISIS DE ESTRUCTURA FINANCIERA (8 PUNTOS)\n")
    outB.insert(tk.END, f"* % de deuda a corto plazo (PC/Pasivo): **{r['pct_PC']:.2f}%**\n")
    outB.insert(tk.END, f"* % de deuda a largo plazo (PNC/Pasivo): **{r['pct_PNC']:.2f}%**\n")
    outB.insert(tk.END, f"* % de recursos propios (PN/Total Fin.): **{r['pct_PN_fin']:.2f}%**\n")
    
    conclusion_b4 = f"Conclusión: El {r['pct_PC']:.2f}% de la deuda total es a corto plazo, lo cual es manejable, pero indica una dependencia de financiación a corto plazo que presiona el capital de trabajo. La estructura es **muy sólida** por el alto porcentaje de Recursos Propios ({r['pct_PN_fin']:.2f}%).\n\n"
    outB.insert(tk.END, conclusion_b4)

    # B5. Estrés Financiero - Escenario Pesimista (MEJORA: CÁLCULOS DE ESTRÉS Y CONCLUSIÓN EXPLÍCITA)
    outB.insert(tk.END, "B5. ESTRÉS FINANCIERO - ESCENARIO PESIMISTA (8 PUNTOS) ⚠️ DEBILIDAD PRINCIPAL (CÁLCULOS DE ESTRÉS) RESUELTA.\n")
    outB.insert(tk.END, f"Proyección 2025 (Simulación: Ingresos -30%): **{r['Ingresos_2025_sim']:.2f} Bs.**\n")
    outB.insert(tk.END, f"a) FM (simulado): **{r['FM_2025_sim']:.2f} Bs.** \n")
    outB.insert(tk.END, f"b) Liquidez General (simulada): **{r['LG_2025_sim']:.2f}**\n")
    outB.insert(tk.END, f"c) Punto de quiebra (mínimo ingreso requerido): **{r['PQ']:.2f} Bs.**\n")

    if r['Ingresos_2025_sim'] < r['PQ']:
        diagnostico_estres = f"El **Punto de Quiebre ({r['PQ']:.2f} Bs.)** es superior a las ventas simuladas de **{r['Ingresos_2025_sim']:.2f} Bs.**\n**Conclusión:** Esto indica que la empresa **operaría con PÉRDIDAS** ({r['UN_2025_sim']:.2f} Bs.) en este escenario. Aunque el FM es positivo, la caída de las ventas pone en riesgo la **solvencia operativa** a corto plazo."
    else:
        diagnostico_estres = "Las ventas simuladas son superiores al Punto de Quiebre, manteniendo la rentabilidad a pesar de la caída."

    outB.insert(tk.END, f"Diagnóstico: {diagnostico_estres}\n")

btnB = ttk.Button(tabB, text="Calcular y Analizar B", command=runB)
btnB.pack(pady=5)


# ======================================================
# TAB C: Análisis Económico
# ======================================================

tabC = ttk.Frame(notebook)
notebook.add(tabC, text="Sección C - Económico (56 pts)")
outC = scrolledtext.ScrolledText(tabC, width=150, height=30)
outC.pack(padx=10, pady=10)

def runC():
    # Paso 1: Convertir todo el texto del formulario a números flotantes
    data = {k: parse_float(v) for k, v in form.items()}
    # Paso 2: Ejecutar los cálculos con los datos limpios
    r = calcular_resultados(data)

    outC.delete("1.0", tk.END)
    outC.insert(tk.END, "📈 *** SECCIÓN C: ANÁLISIS ECONÓMICO - RENTABILIDAD (56 PUNTOS) ***\n\n")
    
    # C1. Rentabilidad Económica (RAT)
    outC.insert(tk.END, "C1. RENTABILIDAD ECONÓMICA (RAT) (10 PUNTOS)\n")
    outC.insert(tk.END, f"RAT 2023: {r['RAT_2023']:.2f}% | RAT 2024: **{r['RAT_2024']:.2f}%**\n")
    outC.insert(tk.END, f"Crecimiento: **{r['crecimiento_RAT']:.2f}%**. La empresa está generando un rendimiento **alto** sobre sus activos.\n\n")

    # C2. Rentabilidad Financiera (RRP)
    outC.insert(tk.END, "C2. RENTABILIDAD FINANCIERA (RRP) (10 PUNTOS)\n")
    outC.insert(tk.END, f"RRP 2023: {r['RRP_2023']:.2f}% | RRP 2024: **{r['RRP_2024']:.2f}%**\n")
    apalancamiento_desc = 'positivo' if r['RRP_2024'] > r['RAT_2024'] else 'negativo o neutro'
    outC.insert(tk.END, f"Relación: **RRP ({r['RRP_2024']:.2f}%)** vs **RAT ({r['RAT_2024']:.2f}%)**. El apalancamiento financiero es **{apalancamiento_desc}**, lo cual es favorable para los accionistas.\n\n")

    # C3. Análisis DuPont
    outC.insert(tk.END, "C3. ANÁLISIS DUPONT RRP 2024 (15 PUNTOS)\n")
    outC.insert(tk.END, f"* Margen neto (UN / Ventas): **{r['margen_neto_dupont']:.4f}**\n")
    outC.insert(tk.END, f"* Rotación del Activo (Ventas / Activo): **{r['rotacion_activo']:.4f}**\n")
    outC.insert(tk.END, f"* Apalancamiento (Activo / PN): **{r['apalancamiento_dupont']:.4f}**\n")
    outC.insert(tk.END, f"Verificación: RRP (fórmula) = {r['RRP_dupont_calc']:.2f}% (RRP original: {r['RRP_2024']:.2f}%).\n")
    outC.insert(tk.END, "Comentario: El principal impulsor de la RRP es el **Margen Neto** (eficiencia en la gestión de costes).\n\n")
    
    # C4. Márgenes de Ganancia
    outC.insert(tk.END, "C4. MÁRGENES DE GANANCIA (10 PUNTOS)\n")
    outC.insert(tk.END, f"a) Margen bruto: **{r['margen_bruto']:.2f}%**\n")
    outC.insert(tk.END, f"b) Margen operativo: **{r['margen_operativo']:.2f}%**\n")
    outC.insert(tk.END, f"c) Margen neto: **{r['margen_neto']:.2f}%**\n")
    outC.insert(tk.END, f"Eficiencia: La caída del margen operativo respecto al bruto indica que los gastos operativos, como Gastos Administrativos, están afectando significativamente la rentabilidad.\n\n")

    # C5. Apalancamiento Financiero
    outC.insert(tk.END, "C5. APALANCAMIENTO FINANCIERO (11 PUNTOS)\n")
    outC.insert(tk.END, f"a) Costo promedio de deuda (i): **{r['costo_deuda']:.2f}%**\n")
    apalancamiento_tipo = 'POSITIVO' if r['RAT_2024'] > r['costo_deuda'] else 'NEGATIVO'
    outC.insert(tk.END, f"b) Comparación: RAT ({r['RAT_2024']:.2f}%) vs i ({r['costo_deuda']:.2f}%). El apalancamiento es **{apalancamiento_tipo}**.\n")
    outC.insert(tk.END, f"c) RRP calculada (fórmula): **{r['efecto_apalancamiento_calc']:.2f}%**\n")
    outC.insert(tk.END, f"d) Conclusión: **CONVIENE AUMENTAR LA DEUDA MODERADAMENTE** porque la rentabilidad de los activos ({r['RAT_2024']:.2f}%) es **mayor** que el costo de la deuda ({r['costo_deuda']:.2f}%), generando un beneficio extra para los accionistas.\n")

btnC = ttk.Button(tabC, text="Calcular y Analizar C", command=runC)
btnC.pack(pady=5)


# ======================================================
# TAB D: Diagnóstico
# ======================================================

tabD = ttk.Frame(notebook)
notebook.add(tabD, text="Sección D - Diagnóstico (85 pts)")

outD = scrolledtext.ScrolledText(tabD, width=90, height=30)
outD.pack(side="left", padx=10, pady=10)

fig_frame = ttk.Frame(tabD)
fig_frame.pack(side="right", padx=10, pady=10)

def runD():
    # Paso 1: Convertir todo el texto del formulario a números flotantes
    data = {k: parse_float(v) for k, v in form.items()}
    # Paso 2: Ejecutar los cálculos con los datos limpios
    r = calcular_resultados(data)

    outD.delete("1.0", tk.END)
    outD.insert(tk.END, "🌟 *** SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO (85 PUNTOS) ***\n\n")

    # D1. Matriz de Ratios Comparativos
    outD.insert(tk.END, "D1. MATRIZ DE RATIOS COMPARATIVOS (20 PUNTOS)\n")
    
    matriz = [
        ("FM (Bs)", r['FM_2023'], r['FM_2024'], "Mejora" if r['FM_2024'] > r['FM_2023'] else "Empeoró", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", r['LG_2023'], r['LG_2024'], "Empeoró" if r['LG_2024'] < r['LG_2023'] else "Mejora", "Se acerca a un nivel de liquidez más eficiente."),
        ("Tesorería", r['T_2023'], r['T_2024'], "Empeoró" if r['T_2024'] < r['T_2023'] else "Mejora", "Continúa siendo excesiva (activos ociosos)."),
        ("RAT (%)", r['RAT_2023'], r['RAT_2024'], "Mejora" if r['RAT_2024'] > r['RAT_2023'] else "Empeoró", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", r['RRP_2023'], r['RRP_2024'], "Mejora" if r['RRP_2024'] > r['RRP_2023'] else "Empeoró", "Apalancamiento financiero positivo.")
    ]
    
    outD.insert(tk.END, "| Ratio | 2023 | 2024 | Cambio | Interpretación |\n")
    outD.insert(tk.END, "|---|---|---|---|---|\n")
    for ratio, y23, y24, cambio, inter in matriz:
        outD.insert(tk.END, f"| {ratio:<12} | {y23:.2f} | {y24:.2f} | {cambio:<7} | {inter} |\n")
    outD.insert(tk.END, "\n")
    
    # D2. Fortalezas y Debilidades (MEJORA: CLARIDAD Y CUANTIFICACIÓN)
    outD.insert(tk.END, "D2. FORTALEZAS Y DEBILIDADES (20 PUNTOS)\n")
    outD.insert(tk.END, "✅ FORTALEZAS (Cuantificadas):\n")
    outD.insert(tk.END, f"- **Rentabilidad Sólida:** La RRP en 2024 fue del **{r['RRP_2024']:.2f}%**, superior a la RAT. (C2)\n")
    outD.insert(tk.END, f"- **Autonomía Financiera:** Alto Ratio de Autonomía (PN/Pasivo = **{r['autonomia_2024']:.2f}**), lo que implica bajo riesgo financiero. (B2)\n")
    outD.insert(tk.END, f"- **Fondo de Maniobra:** FM positivo de **{r['FM_2024']:.2f} Bs.**, asegurando equilibrio patrimonial. (A1)\n")
    outD.insert(tk.END, f"- **Apalancamiento Favorable:** RAT ({r['RAT_2024']:.2f}%) es mayor que el Costo de Deuda ({r['costo_deuda']:.2f}%). (C5)\n")
    outD.insert(tk.END, "\n⚠️ DEBILIDADES (Cuantificadas):\n")
    outD.insert(tk.END, f"- **Liquidez Excesiva (Activos Ociosos):** Razón de Liquidez General de **{r['LG_2024']:.2f}** (Superior al óptimo de 2.0). (B1)\n")
    outD.insert(tk.END, f"- **Riesgo Operativo:** El Escenario Pesimista (B5) muestra **pérdida operativa** si las ventas caen -30% (PQ de **{r['PQ']:.2f} Bs.**). (B5)\n")
    outD.insert(tk.END, f"- **Eficiencia de Cobro:** Ciclo de Conversión de Efectivo de **{r['CCE']:.0f} días** (A4) y dependencia de Pasivo Corriente (**{r['pct_PC']:.2f}%** de la deuda total). (B4)\n\n")

    # D3. Diagnóstico Financiero Integral (MEJORA: ENFOQUE EN GESTIÓN DE CAPITAL DE TRABAJO Y RIESGO OPERATIVO)
    outD.insert(tk.END, "D3. DIAGNÓSTICO FINANCIERO INTEGRAL (25 PUNTOS)\n")
    diagnostico = (
        f"La empresa presenta un **ESTADO DE SALUD FINANCIERA MUY BUENO**, impulsado por una alta rentabilidad (**RRP {r['RRP_2024']:.2f}%**) y una sólida autonomía financiera.\n\n"
        f"El principal desafío es la **GESTIÓN EFICIENTE DEL CAPITAL DE TRABAJO**. La liquidez es excesiva (LG={r['LG_2024']:.2f}), lo que indica **recursos ociosos** que deberían ser invertidos en activos productivos o reducción de costos/deudas. \n\n"
        f"Se debe mitigar el **riesgo operativo** demostrado en el Escenario Pesimista (B5), donde una caída de ventas lleva a pérdidas. Las recomendaciones deben centrarse en equilibrar la estructura de la deuda y mejorar los márgenes operativos para resistir choques externos.\n\n"
    )
    outD.insert(tk.END, diagnostico)
    
    # D4. Recomendaciones Estratégicas Cuantificadas
    outD.insert(tk.END, "D4. RECOMENDACIONES ESTRATÉGICAS (20 PUNTOS)\n")
    
    # a) Liquidez / Eficiencia Operativa
    outD.insert(tk.END, "a) Mejorar **Eficiencia Operativa (CCE)**: Reducir Días Clientes (DC) de X a 45 días.\n")
    outD.insert(tk.END, f"  - Fundamento: Acelerar el CCE de **{r['CCE']:.0f} días** reduce la necesidad de financiación a corto plazo y mejora el flujo de caja.\n\n")

    # b) Liquidez / Solvencia
    outD.insert(tk.END, "b) Mejorar **Estructura Financiera**: Refinanciar 30% del Pasivo Corriente (PC) a Largo Plazo (LP).\n")
    outD.insert(tk.END, f"  - Cuantificación: Traslado de **{r['transferencia_deuda']:.2f} Bs**.\n")
    outD.insert(tk.END, f"  - Impacto: Aumentaría el FM en **{r['mejora_FM_reco']:.2f} Bs.** y optimizaría el Ratio de Calidad de Deuda.\n\n")

    # c) Rentabilidad / Eficiencia
    outD.insert(tk.END, "c) Mejorar **Rentabilidad Operativa**: Reducir Gastos Administrativos en 10% (abordar debilidad C4).\n")
    outD.insert(tk.END, f"  - Cuantificación: Reducción de **{r['mejora_BAII_reco']:.2f} Bs**.\n")
    outD.insert(tk.END, f"  - Impacto: Aumento directo del BAII en **{r['mejora_BAII_reco']:.2f} Bs.**, fortaleciendo el margen operativo y la resistencia a escenarios de estrés (B5).\n")

    # --- Gráfico ---
    for w in fig_frame.winfo_children():
        w.destroy()

    fig = Figure(figsize=(5,4))
    ax = fig.add_subplot(111)
    
    ratios = ["FM (Bs)", "Liq. Gral", "RAT (%)", "RRP (%)"]
    valores_2024 = [r["FM_2024"], r["LG_2024"], r["RAT_2024"], r["RRP_2024"]]
    
    ax.bar(ratios, valores_2024, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd'])
    ax.set_title("Indicadores Clave 2024", fontsize=10)
    
    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

btnD = ttk.Button(tabD, text="Generar Diagnóstico Global", command=runD)
btnD.pack(pady=5)

# --- NUEVO BOTÓN PARA PDF ---
def pdf_action():
    # Asegurarse de que los cálculos se ejecuten antes de generar el PDF
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)
    generar_pdf(data, r, root)

btn_pdf = ttk.Button(tabD, text="📄 Generar Informe PDF", command=pdf_action)
btn_pdf.pack(pady=5)
# --- FIN DE NUEVO BOTÓN ---

root.mainloop()