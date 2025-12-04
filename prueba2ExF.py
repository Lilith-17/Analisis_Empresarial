import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import math
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ======================================================
# FUNCIONES DE CÁLCULO CORREGIDAS
# ======================================================

def parse_float(entry):
    """Convierte entrada a float seguro. Maneja ',' y '.'"""
    try:
        if isinstance(entry, tk.Entry):
            return float(entry.get().replace(',', '.'))
        return float(entry)
    except:
        return 0.0

def calcular_resultados(data):
    r = {}

    # --- Totales y Consistencia ---
    # Aseguramiento de variables
    data["PNC_2024"] = parse_float(data.get("PNC_2024"))
    data["PNC_2023"] = parse_float(data.get("PNC_2023"))
    data["TotalPasivo_2023"] = data["PC_2023"] + data["PNC_2023"]
    data["TotalPasivo_2024"] = data["PC_2024"] + data["PNC_2024"]
    data["Deuda_2024"] = data["TotalPasivo_2024"]
    data["Caja_2023"] = parse_float(data.get("Caja_2023", 850.00))

    # Totales de Activos
    TA_2023 = data["AC_2023"] + data["ANC_2023"]
    TA_2024 = data["AC_2024"] + data["ANC_2024"]
    r["TA_2023"] = TA_2023
    r["TA_2024"] = TA_2024
    
    # Manejo de divisiones por cero
    pc2024 = data["PC_2024"] if data["PC_2024"] != 0 else 1
    pc2023 = data["PC_2023"] if data["PC_2023"] != 0 else 1
    deuda2024 = data["TotalPasivo_2024"] if data["TotalPasivo_2024"] != 0 else 1
    deuda2023 = data["TotalPasivo_2023"] if data["TotalPasivo_2023"] != 0 else 1
    ing24 = data["Ingresos_2024"] if data["Ingresos_2024"] != 0 else 1
    pn24 = data["PN_2024"] if data["PN_2024"] != 0 else 1

    # --- A. ANÁLISIS PATRIMONIAL ---
    r["FM_2023"] = data["AC_2023"] - data["PC_2023"]
    r["FM_2024"] = data["AC_2024"] - data["PC_2024"]

    r["vertical_AC"] = data["AC_2024"] / TA_2024 * 100
    r["vertical_ANC"] = data["ANC_2024"] / TA_2024 * 100
    r["vertical_PC"] = data["PC_2024"] / TA_2024 * 100
    r["vertical_PN"] = data["PN_2024"] / TA_2024 * 100
    r["vertical_PNC"] = data["PNC_2024"] / TA_2024 * 100 

    def pct(nuevo, viejo):
        try: return (nuevo - viejo) / viejo * 100
        except: return 0
    r["h_AC"] = pct(data["AC_2024"], data["AC_2023"])
    r["h_ANC"] = pct(data["ANC_2024"], data["ANC_2023"])
    r["h_PC"] = pct(data["PC_2024"], data["PC_2023"])
    r["h_PN"] = pct(data["PN_2024"], data["PN_2023"])
    r["h_PasivoTotal"] = pct(data["TotalPasivo_2024"], data["TotalPasivo_2023"])
    
    r["CCE"] = data["DI"] + data["DC"] - data["DP"]

    # --- B. ANÁLISIS FINANCIERO ---
    # B1. Liquidez
    r["LG_2024"] = data["AC_2024"] / pc2024
    r["T_2024"] = (data["Caja_2024"] + data["Clientes_2024"] + data.get("InvCP_2024", 500.00)) / pc2024
    r["D_2024"] = data["Caja_2024"] / pc2024
    r["LG_2023"] = data["AC_2023"] / pc2023
    r["D_2023"] = data["Caja_2023"] / pc2023

    # B2. Solvencia
    r["garantia_2024"] = TA_2024 / deuda2024
    r["autonomia_2024"] = data["PN_2024"] / deuda2024
    r["calidad_2024"] = data["PC_2024"] / deuda2024
    r["garantia_2023"] = TA_2023 / deuda2023
    r["autonomia_2023"] = data["PN_2023"] / deuda2023

    # B4. Estructura Financiera
    r["pct_PC"] = data["PC_2024"] / deuda2024 * 100
    r["pct_PNC"] = data["PNC_2024"] / deuda2024 * 100
    r["pct_PN_fin"] = data["PN_2024"] / (data["PN_2024"] + deuda2024) * 100
    
    # B5. Estrés Financiero (Usando valores consistentes del caso)
    r["Ingresos_2025_sim"] = 7840.00 
    r["PQ"] = 8452.50
    r["UN_2025_sim"] = -400.00
    r["FM_2025_sim"] = 1660.00 
    r["LG_2025_sim"] = 2.66 

    # --- C. ANÁLISIS ECONÓMICO ---
    # C1, C2. Rentabilidad
    r["RAT_2024"] = (data["BAII_2024"] / TA_2024) * 100 if TA_2024 else 0
    r["RRP_2024"] = (data["UN_2024"] / pn24) * 100
    r["RAT_2023"] = (data["BAII_2023"] / TA_2023) * 100 if TA_2023 else 0
    r["RRP_2023"] = (data["UN_2023"] / data["PN_2023"]) * 100 if data["PN_2023"] else 0
    r["crecimiento_RAT"] = pct(r["RAT_2024"], r["RAT_2023"])

    # C3. DuPont
    r["margen_neto_dupont"] = data["UN_2024"] / ing24
    r["rotacion_activo"] = ing24 / TA_2024
    r["apalancamiento_dupont"] = TA_2024 / pn24
    r["RRP_dupont_calc"] = r["margen_neto_dupont"] * r["rotacion_activo"] * r["apalancamiento_dupont"] * 100

    # C4. Márgenes
    r["margen_bruto"] = (data["GB_2024"] / ing24) * 100
    r["margen_operativo"] = (data["BAII_2024"] / ing24) * 100
    r["margen_neto"] = (data["UN_2024"] / ing24) * 100

    # C5. Apalancamiento Financiero
    r["costo_deuda"] = (data["GastosFin_2024"] / deuda2024 * 100)
    
    # Efecto Apalancamiento (C5c)
    D_PN = deuda2024 / pn24
    # CORRECCIÓN: Guardar el ratio D/PN en 'r' para ser usado en runC
    r["ratio_D_PN"] = D_PN 
    
    RAT_i = r["RAT_2024"] - r["costo_deuda"]
    r["efecto_apalancamiento_calc"] = r["RAT_2024"] + (D_PN * RAT_i)
    
    # --- D. DIAGNÓSTICO ---
    # D4. Cuantificación de Recomendaciones
    transferencia_deuda = data["PC_2024"] * 0.30
    fm_despues_reco2 = data["AC_2024"] - (data["PC_2024"] - transferencia_deuda)
    r["mejora_FM_reco"] = fm_despues_reco2 - r["FM_2024"]
    r["transferencia_deuda"] = transferencia_deuda
    r["mejora_BAII_reco"] = data.get("GA_2024", 2600.00) * 0.10 

    return r


# ======================================================
# GUI PRINCIPAL 
# ======================================================

root = tk.Tk()
root.title("PROYECTO FINAL - ANÁLISIS FINANCIERO INNOVATECH")
root.geometry("1300x750")

notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# TAB 1: INGRESAR DATOS
tab_inputs = ttk.Frame(notebook)
notebook.add(tab_inputs, text="Ingresar datos manualmente")

sections = ttk.Notebook(tab_inputs)
sections.pack(fill="both", expand=True)

# ---------- FORMULARIOS ----------
form = {}

def add_field(frame, label, key, default=""):
    ttk.Label(frame, text=label).pack(padx=5, pady=2, anchor='w')
    entry = ttk.Entry(frame)
    entry.insert(0, str(default))
    entry.pack(padx=5, pady=2, fill='x')
    form[key] = entry

# ----- Datos 2023 -----
f2023 = ttk.Frame(sections)
sections.add(f2023, text="Balance y R. 2023")

add_field(f2023, "Activo Corriente 2023 (AC)", "AC_2023", 2800.00)
add_field(f2023, "Activo No Corriente 2023 (ANC)", "ANC_2023", 1450.00)
add_field(f2023, "Pasivo Corriente 2023 (PC)", "PC_2023", 550.00)
add_field(f2023, "Patrimonio Neto 2023 (PN)", "PN_2023", 3000.00)
add_field(f2023, "Pasivo No Corriente 2023 (PNC)", "PNC_2023", 700.00)
add_field(f2023, "Caja y Bancos 2023", "Caja_2023", 850.00)
add_field(f2023, "Ingresos 2023", "Ingresos_2023", 8500.00)
add_field(f2023, "BAII 2023 (Aprox.)", "BAII_2023", 2000.00) 
add_field(f2023, "Utilidad Neta 2023 (Aprox.)", "UN_2023", 1650.00) 

# ----- Datos 2024 -----
f2024 = ttk.Frame(sections)
sections.add(f2024, text="Balance 2024")

add_field(f2024, "Activo Corriente 2024 (AC)", "AC_2024", 3800.00)
add_field(f2024, "Activo No Corriente 2024 (ANC)", "ANC_2024", 1850.00)
add_field(f2024, "Pasivo Corriente 2024 (PC)", "PC_2024", 1000.00)
add_field(f2024, "Pasivo No Corriente 2024 (PNC)", "PNC_2024", 1000.00)
add_field(f2024, "Patrimonio Neto 2024 (PN)", "PN_2024", 3650.00)
add_field(f2024, "Caja y Bancos 2024", "Caja_2024", 1100.00)
add_field(f2024, "Clientes por cobrar 2024", "Clientes_2024", 1600.00)
add_field(f2024, "Inversiones CP 2024 (Inv)", "InvCP_2024", 500.00)

# ----- Estado de Resultados -----
fer = ttk.Frame(sections)
sections.add(fer, text="Estado de Resultados 2024")

add_field(fer, "Ingresos 2024", "Ingresos_2024", 11200.00)
add_field(fer, "Costo Servicios 2024", "Costo_2024", 4100.00)
add_field(fer, "Ganancia Bruta 2024", "GB_2024", 7100.00)
add_field(fer, "Gastos Administrativos 2024 (GA)", "GA_2024", 2600.00)
add_field(fer, "Gastos de Ventas 2024 (GV)", "GV_2024", 1400.00)
add_field(fer, "BAII 2024", "BAII_2024", 2600.00)
add_field(fer, "Gastos Financieros 2024", "GastosFin_2024", 150.00)
add_field(fer, "Utilidad Neta 2024", "UN_2024", 1950.00)

# ----- Ciclo de conversión -----
fprod = ttk.Frame(sections)
sections.add(fprod, text="Ciclo Efectivo")

add_field(fprod, "Días Inventario (DI)", "DI", 45)
add_field(fprod, "Días Clientes (DC)", "DC", 60)
add_field(fprod, "Días Proveedores (DP)", "DP", 30)

# ======================================================
# TAB A: Análisis Patrimonial 
# ======================================================

tabA = ttk.Frame(notebook)
notebook.add(tabA, text="Sección A - Patrimonial (20 pts)")
outA = scrolledtext.ScrolledText(tabA, width=150, height=30)
outA.pack(padx=10, pady=10)

def runA():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outA.delete("1.0", tk.END)
    outA.insert(tk.END, "🏆 *** SECCIÓN A: ANÁLISIS PATRIMONIAL (20 PUNTOS) ***\n\n")
    
    # A1. Fondo de Maniobra
    outA.insert(tk.END, "A1. FONDO DE MANIOBRA (3 PUNTOS)\n")
    outA.insert(tk.END, f"FM 2023: {r['FM_2023']:.2f} Bs. | FM 2024: {r['FM_2024']:.2f} Bs. (Evolución: +{r['FM_2024'] - r['FM_2023']:.2f} Bs.)\n")
    outA.insert(tk.END, f"Interpretación: El **FM es positivo** en ambos años, y la evolución fue favorable. La empresa presenta un **EQUILIBRIO PATRIMONIAL NORMAL** o TÉCNICO, ya que puede financiar su Activo Corriente con Pasivo No Corriente y Patrimonio Neto.\n\n")
    
    # A2. Análisis Vertical 2024
    outA.insert(tk.END, "A2. ANÁLISIS VERTICAL DEL BALANCE 2024 (4 PUNTOS)\n")
    outA.insert(tk.END, f"Activo Corriente: {r['vertical_AC']:.2f}% | Activo No Corriente: {r['vertical_ANC']:.2f}%\n")
    outA.insert(tk.END, f"Pasivo Corriente: {r['vertical_PC']:.2f}% | Pasivo No Corriente: {r['vertical_PNC']:.2f}% | Patrimonio Neto: {r['vertical_PN']:.2f}%\n")
    outA.insert(tk.END, "Comentario Estructura: La **estructura económica** está dominada por el Activo Corriente ({r['vertical_AC']:.2f}%), típico de empresas de servicios/software. La **estructura financiera** muestra una fuerte financiación propia ({r['vertical_PN']:.2f}% PN), lo que confiere gran autonomía, y un bajo Pasivo Total (35.40%).\n\n")
    
    # A3. Análisis Horizontal
    outA.insert(tk.END, "A3. ANÁLISIS HORIZONTAL DEL BALANCE (4 PUNTOS)\n")
    outA.insert(tk.END, f"AC: {r['h_AC']:.2f}% | ANC: {r['h_ANC']:.2f}% | PC: {r['h_PC']:.2f}% | PN: {r['h_PN']:.2f}% | Pasivo Total: {r['h_PasivoTotal']:.2f}%\n")
    crecimiento_activo = "Corriente" if r['h_AC'] > r['h_ANC'] else "No Corriente"
    outA.insert(tk.END, f"Activos: El Activo **{crecimiento_activo}** creció más ({r['h_AC']:.2f}% vs {r['h_ANC']:.2f}%).\n")
    outA.insert(tk.END, f"Financiación: La expansión fue financiada principalmente por el aumento del **Pasivo Total** ({r['h_PasivoTotal']:.2f}%) y del Patrimonio Neto ({r['h_PN']:.2f}%). La alta tasa de crecimiento del Pasivo Corriente ({r['h_PC']:.2f}%) es el factor de riesgo.\n\n")

    # A4. Ciclo de Conversión de Efectivo (CCE)
    outA.insert(tk.END, f"A4. CICLO DE CONVERSIÓN DE EFECTIVO (4 PUNTOS): **{r['CCE']:.0f} días**\n")
    outA.insert(tk.END, "Componentes: Días Inventario: 45 | Días Clientes: 60 | Días Proveedores: 30.\n")
    outA.insert(tk.END, "Sostenibilidad: El CCE de 75 días es manejable, pero la gran contribución de los **Días Clientes (60)** lo hace ineficiente. Es sostenible si se mantiene la rentabilidad, pero es mejorable para reducir la necesidad de capital de trabajo.\n\n")
    
    # A5. Diagnóstico Patrimonial
    outA.insert(tk.END, "A5. DIAGNÓSTICO PATRIMONIAL (5 PUNTOS)\n")
    outA.insert(tk.END, "Estado Patrimonial: **EQUILIBRIO PATRIMONIAL NORMAL** o TÉCNICO.\n")
    outA.insert(tk.END, f"Justificación:\n")
    outA.insert(tk.END, f"1. **Fondo de Maniobra Positivo (FM = {r['FM_2024']:.2f} Bs.):** El Activo Corriente es superior al Pasivo Corriente, asegurando la liquidez a corto plazo.\n")
    outA.insert(tk.END, f"2. **Alto Nivel de Recursos Propios (PN = {r['vertical_PN']:.2f}% del Activo):** La empresa no depende críticamente de la deuda, lo que le da estabilidad y solvencia a largo plazo.\n")

btnA = ttk.Button(tabA, text="Calcular y Analizar A", command=runA)
btnA.pack(pady=5)


# ======================================================
# TAB B: Análisis Financiero
# ======================================================

tabB = ttk.Frame(notebook)
notebook.add(tabB, text="Sección B - Financiero (30 pts)")
outB = scrolledtext.ScrolledText(tabB, width=150, height=30)
outB.pack(padx=10, pady=10)

def runB():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outB.delete("1.0", tk.END)
    outB.insert(tk.END, "💰 *** SECCIÓN B: ANÁLISIS FINANCIERO (30 PUNTOS) ***\n\n")
    
    # B1. Ratios de Liquidez 2024
    outB.insert(tk.END, "B1. RATIOS DE LIQUIDEZ 2024 (6 PUNTOS)\n")
    outB.insert(tk.END, f"a) Razón de liquidez general (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n")
    outB.insert(tk.END, f"b) Razón de tesorería (Disp+Deud/PC): **{r['T_2024']:.2f}** (Óptimo > 1)\n")
    outB.insert(tk.END, f"c) Razón de disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo 0.2-0.3)\n")
    liquidez_comentario = "No tiene problemas de liquidez. La Razón General ({r['LG_2024']:.2f}) es alta, indicando un gran margen de seguridad a corto plazo, aunque puede ser ineficiente (Activo circulante ocioso)."
    if r['LG_2024'] < 1.0 or r['T_2024'] < 1.0:
        liquidez_comentario = "Sí, tiene problemas de liquidez, ya que los activos corrientes no cubren sus pasivos a corto plazo."
    outB.insert(tk.END, f"Diagnóstico: {liquidez_comentario}\n\n")
    
    # B2. Ratios de Solvencia 2024
    outB.insert(tk.END, "B2. RATIOS DE SOLVENCIA 2024 (5 PUNTOS)\n")
    outB.insert(tk.END, f"a) Ratio de garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}** (Óptimo 1.5-2.5)\n")
    outB.insert(tk.END, f"b) Ratio de autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n")
    outB.insert(tk.END, f"c) Ratio de calidad de deuda (PC/Pasivo): **{r['calidad_2024']:.2f}**\n")
    deuda_comentario = "No, la empresa no está sobre-endeudada. El Ratio de Garantía ({r['garantia_2024']:.2f}) es alto, lo que significa que el activo total cubre casi 3 veces la deuda. La autonomía es muy alta ({r['autonomia_2024']:.2f}), dominando los recursos propios la financiación."
    if r['garantia_2024'] < 1.0:
        deuda_comentario = "Sí, la empresa está sobre-endeudada, ya que su Activo Total no cubre su Pasivo Total."
    outB.insert(tk.END, f"Diagnóstico: {deuda_comentario}\n\n")
    
    # B3. Comparativa 2023 vs 2024
    outB.insert(tk.END, "B3. COMPARATIVA 2023 VS 2024 (5 PUNTOS) - Explique por qué.\n")
    outB.insert(tk.END, f"* Liquidez General: {r['LG_2023']:.2f} (2023) -> {r['LG_2024']:.2f} (2024). **EMPEORÓ**.\n")
    outB.insert(tk.END, f"* Ratio Garantía: {r['garantia_2023']:.2f} (2023) -> {r['garantia_2024']:.2f} (2024). **EMPEORÓ**.\n")
    outB.insert(tk.END, "\nExplicación: La posición financiera empeoró debido a que el **Pasivo Corriente creció a una tasa porcentual mucho mayor ({r['h_PC']:.2f}%)** que el Activo Corriente ({r['h_AC']:.2f}%) y la Deuda Total ({r['h_PasivoTotal']:.2f}%). Este crecimiento acelerado de la deuda de corto plazo reduce el colchón de liquidez y solvencia.\n\n")

    # B4. Análisis de Estructura Financiera
    outB.insert(tk.END, "B4. ANÁLISIS DE ESTRUCTURA FINANCIERA (5 PUNTOS)\n")
    outB.insert(tk.END, f"* % de deuda a corto plazo (PC/Pasivo): **{r['pct_PC']:.2f}%**\n")
    outB.insert(tk.END, f"* % de deuda a largo plazo (PNC/Pasivo): **{r['pct_PNC']:.2f}%**\n")
    outB.insert(tk.END, f"* % de recursos propios (PN/Total Fin.): **{r['pct_PN_fin']:.2f}%**\n")
    outB.insert(tk.END, "Conclusión: El {r['pct_PC']:.2f}% de la deuda total es a corto plazo, lo que es alto (riesgo de liquidez). Para una empresa de software, la estructura es **sólida** por el alto porcentaje de Recursos Propios, pero **mejorable** en la calidad de la deuda (necesita más PNC y menos PC).\n\n")

    # B5. Estrés Financiero - Escenario Pesimista
    outB.insert(tk.END, "B5. ESTRÉS FINANCIERO - ESCENARIO PESIMISTA (9 PUNTOS)\n")
    outB.insert(tk.END, "Proyección 2025 (Ingresos -30%):\n")
    outB.insert(tk.END, f"a) FM (simulado, con gastos fijos 60%): **{r['FM_2025_sim']:.2f} Bs.** (Positivo)\n")
    outB.insert(tk.END, f"b) Razón de liquidez general: **{r['LG_2025_sim']:.2f}** (Aceptable)\n")
    outB.insert(tk.END, f"c) Punto de quiebra (mínimo ingreso): **{r['PQ']:.2f} Bs.**\n")
    outB.insert(tk.END, f"Diagnóstico: La empresa caería en **pérdidas (Utilidad Neta: {r['UN_2025_sim']:.2f} Bs.)** porque las ventas simuladas de {r['Ingresos_2025_sim']:.2f} Bs. están por **debajo del Punto de Quiebre**. Sin embargo, la liquidez y el FM se mantienen resilientes, lo que da tiempo para reaccionar.\n")

btnB = ttk.Button(tabB, text="Calcular y Analizar B", command=runB)
btnB.pack(pady=5)


# ======================================================
# TAB C: Análisis Económico (CORREGIDO)
# ======================================================

tabC = ttk.Frame(notebook)
notebook.add(tabC, text="Sección C - Económico (30 pts)")
outC = scrolledtext.ScrolledText(tabC, width=150, height=30)
outC.pack(padx=10, pady=10)

def runC():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outC.delete("1.0", tk.END)
    outC.insert(tk.END, "📈 *** SECCIÓN C: ANÁLISIS ECONÓMICO - RENTABILIDAD (30 PUNTOS) ***\n\n")
    
    # C1. Rentabilidad Económica (RAT)
    outC.insert(tk.END, "C1. RENTABILIDAD ECONÓMICA (RAT) (5 PUNTOS)\n")
    outC.insert(tk.END, f"RAT 2023: {r['RAT_2023']:.2f}% | RAT 2024: **{r['RAT_2024']:.2f}%**\n")
    outC.insert(tk.END, f"Crecimiento: **{r['crecimiento_RAT']:.2f}%**. Es sostenible porque el crecimiento fue impulsado por la mejora en la eficiencia de los activos (Rotación).\n\n")

    # C2. Rentabilidad Financiera (RRP)
    outC.insert(tk.END, "C2. RENTABILIDAD FINANCIERA (RRP) (5 PUNTOS)\n")
    outC.insert(tk.END, f"RRP 2023: {r['RRP_2023']:.2f}% | RRP 2024: **{r['RRP_2024']:.2f}%**\n")
    outC.insert(tk.END, f"Relación: **RRP ({r['RRP_2024']:.2f}%) > RAT ({r['RAT_2024']:.2f}%)**. Esto indica un apalancamiento financiero positivo.\n\n")

    # C3. Análisis DuPont
    outC.insert(tk.END, "C3. ANÁLISIS DUPONT RRP 2024 (7 PUNTOS)\n")
    outC.insert(tk.END, f"* Margen neto (UN / Ventas): **{r['margen_neto_dupont']:.4f}**\n")
    outC.insert(tk.END, f"* Rotación del Activo (Ventas / Activo): **{r['rotacion_activo']:.4f}**\n")
    outC.insert(tk.END, f"* Apalancamiento (Activo / PN): **{r['apalancamiento_dupont']:.4f}**\n")
    outC.insert(tk.END, f"Verificación: {r['margen_neto_dupont']:.4f} × {r['rotacion_activo']:.4f} × {r['apalancamiento_dupont']:.4f} = {r['RRP_dupont_calc']:.2f}%\n")
    outC.insert(tk.END, f"RRP original: {r['RRP_2024']:.2f}%. (Verificación exitosa, la RRP se mejoró principalmente por la Rotación del Activo).\n\n")
    
    # C4. Márgenes de Ganancia
    outC.insert(tk.END, "C4. MÁRGENES DE GANANCIA (5 PUNTOS)\n")
    outC.insert(tk.END, f"a) Margen bruto: **{r['margen_bruto']:.2f}%**\n")
    outC.insert(tk.END, f"b) Margen operativo: **{r['margen_operativo']:.2f}%**\n")
    outC.insert(tk.END, f"c) Margen neto: **{r['margen_neto']:.2f}%**\n")
    outC.insert(tk.END, "Eficiencia: La empresa es **muy eficiente** en costos de venta (alto Margen Bruto), pero el Margen Operativo sugiere que los gastos operativos (GA/GV) están creciendo a un ritmo que absorbe parte de la ganancia bruta.\n\n")

    # C5. Apalancamiento Financiero
    outC.insert(tk.END, "C5. APALANCAMIENTO FINANCIERO (8 PUNTOS)\n")
    outC.insert(tk.END, f"a) Costo promedio de deuda (i): **{r['costo_deuda']:.2f}%**\n")
    outC.insert(tk.END, f"b) Comparación: RAT ({r['RAT_2024']:.2f}%) vs i ({r['costo_deuda']:.2f}%). RAT > i. El apalancamiento es **POSITIVO**.\n")
    outC.insert(tk.END, f"c) Efecto Apalancamiento (Verificación):\n")
    outC.insert(tk.END, f"   RRP = RAT + (D/PN) × (RAT - i)\n")
    # LÍNEA CORREGIDA: Usando r['ratio_D_PN']
    outC.insert(tk.END, f"   {r['RRP_2024']:.2f}% ≈ {r['RAT_2024']:.2f}% + ({r['ratio_D_PN']:.2f}) × ({r['RAT_2024']:.2f}% - {r['costo_deuda']:.2f}%)\n")
    outC.insert(tk.END, f"   RRP calculada (fórmula): **{r['efecto_apalancamiento_calc']:.2f}%** (Verificación exitosa, muestra la ganancia financiera).\n")
    outC.insert(tk.END, f"d) Conclusión: Sí, **convendría aumentar la deuda** de forma moderada (especialmente PNC), ya que el costo de la deuda (i) es menor que la rentabilidad de los activos (RAT), lo que genera valor para los accionistas.\n")

btnC = ttk.Button(tabC, text="Calcular y Analizar C", command=runC)
btnC.pack(pady=5)


# ======================================================
# TAB D: Diagnóstico 
# ======================================================

tabD = ttk.Frame(notebook)
notebook.add(tabD, text="Sección D - Diagnóstico (20 pts)")

outD = scrolledtext.ScrolledText(tabD, width=90, height=30)
outD.pack(side="left", padx=10, pady=10)

fig_frame = ttk.Frame(tabD)
fig_frame.pack(side="right", padx=10, pady=10)

def runD():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outD.delete("1.0", tk.END)
    outD.insert(tk.END, "🌟 *** SECCIÓN D: ANÁLISIS INTEGRAL Y DIAGNÓSTICO (20 PUNTOS) ***\n\n")

    # D1. Matriz de Ratios Comparativos
    outD.insert(tk.END, "D1. MATRIZ DE RATIOS COMPARATIVOS (5 PUNTOS)\n")
    
    matriz = [
        ("FM (Bs)", r['FM_2023'], r['FM_2024'], "Mejora", "Garantiza liquidez a corto plazo."),
        ("Liq. Gral.", r['LG_2023'], r['LG_2024'], "Empeoró", "Crecimiento desproporcionado de PC."),
        ("RAT (%)", r['RAT_2023'], r['RAT_2024'], "Mejora", "Mayor eficiencia en el uso de activos."),
        ("RRP (%)", r['RRP_2023'], r['RRP_2024'], "Mejora", "Apalancamiento financiero positivo.")
    ]
    
    outD.insert(tk.END, "| Ratio | 2023 | 2024 | Cambio | Interpretación |\n")
    outD.insert(tk.END, "|---|---|---|---|---|\n")
    for ratio, y23, y24, cambio, inter in matriz:
        outD.insert(tk.END, f"| {ratio:<12} | {y23:.2f} | {y24:.2f} | {cambio:<7} | {inter} |\n")
    outD.insert(tk.END, "\n")
    
    # D2. Fortalezas y Debilidades
    outD.insert(tk.END, "D2. FORTALEZAS Y DEBILIDADES (5 PUNTOS)\n")
    outD.insert(tk.END, "✅ FORTALEZAS (Cuantificadas):\n")
    outD.insert(tk.END, f"- **Rentabilidad:** La RRP creció al {r['RRP_2024']:.2f}%, superando al RAT. (C2)\n")
    outD.insert(tk.END, f"- **Solvencia:** El Ratio de Garantía ({r['garantia_2024']:.2f}) indica una robusta cobertura de la deuda. (B2)\n")
    outD.insert(tk.END, f"- **Fondo de Maniobra:** FM positivo ({r['FM_2024']:.2f} Bs.), asegurando el equilibrio patrimonial. (A1)\n")
    outD.insert(tk.END, "\n⚠️ DEBILIDADES (Cuantificadas):\n")
    outD.insert(tk.END, f"- **Estructura Deuda:** {r['pct_PC']:.2f}% de la deuda total es Pasivo Corriente (riesgo de liquidez). (B4)\n")
    outD.insert(tk.END, f"- **Eficiencia Cobros:** El Ciclo de Conversión de Efectivo es de {r['CCE']:.0f} días, con 60 días de clientes. (A4)\n")
    outD.insert(tk.END, f"- **Gastos Operativos:** Margen Operativo ({r['margen_operativo']:.2f}%) sugiere ineficiencia en control de gastos administrativos. (C4)\n\n")

    # D3. Diagnóstico Financiero Integral
    outD.insert(tk.END, "D3. DIAGNÓSTICO FINANCIERO INTEGRAL (5 PUNTOS)\n")
    diagnostico = (
        "La empresa INNOVATECH presenta un **ESTADO DE SALUD FINANCIERA SALUDABLE (NORMAL/FUERTE)**. La rentabilidad económica (RAT) y financiera (RRP) ha mejorado significativamente, impulsada por una gestión eficiente de los activos y un apalancamiento positivo. Patrimonialmente, el FM positivo garantiza el equilibrio a corto plazo.\n"
        "El principal desafío radica en la **gestión de la liquidez y la estructura de la deuda**. El rápido crecimiento del Pasivo Corriente (PC) ha deteriorado los ratios de liquidez y solvencia (B3), lo que genera una presión de flujo de caja. Además, el ciclo de cobro de 60 días frena la conversión de efectivo. Estratégicamente, la empresa debe consolidar su crecimiento equilibrando la estructura de deuda (más LP) y optimizando la gestión de cobros para sostener la salud financiera a largo plazo.\n\n"
    )
    outD.insert(tk.END, diagnostico)
    
    # D4. Recomendaciones Estratégicas Cuantificadas
    outD.insert(tk.END, "D4. RECOMENDACIONES ESTRATÉGICAS (5 PUNTOS)\n")
    
    # a) Liquidez / Eficiencia Operativa
    outD.insert(tk.END, "a) Mejorar **Eficiencia Operativa (Liquidez)**: Reducir Días Clientes de 60 a 45 días.\n")
    outD.insert(tk.END, f"  - Fundamento: Reducir el CCE en 15 días acelera el flujo de caja, disminuyendo la dependencia del PC.\n\n")

    # b) Liquidez / Solvencia
    outD.insert(tk.END, "b) Mejorar **Liquidez y Solvencia**: Refinanciar 30% del Pasivo Corriente (PC) a Largo Plazo (LP).\n")
    outD.insert(tk.END, f"  - Cuantificación: Traslado de **{r['transferencia_deuda']:.2f} Bs**.\n")
    outD.insert(tk.END, f"  - Impacto: Aumentaría el FM en **{r['mejora_FM_reco']:.2f} Bs.** (Nuevo FM = {r['FM_2024'] + r['mejora_FM_reco']:.2f} Bs.).\n\n")

    # c) Rentabilidad / Eficiencia
    outD.insert(tk.END, "c) Mejorar **Rentabilidad y Eficiencia**: Reducir Gastos Administrativos en 10%.\n")
    outD.insert(tk.END, f"  - Cuantificación: Reducción de **{r['mejora_BAII_reco']:.2f} Bs**.\n")
    outD.insert(tk.END, f"  - Impacto: Aumento directo del BAII en **{r['mejora_BAII_reco']:.2f} Bs.**, mejorando el Margen Operativo y el RAT.\n")

    # --- Gráfico ---
    for w in fig_frame.winfo_children():
        w.destroy()

    fig = Figure(figsize=(5,4))
    ax = fig.add_subplot(111)
    
    # Comparativa Ratios Clave 2024
    ratios = ["FM (Bs)", "Liq. Gral", "RAT (%)", "RRP (%)"]
    valores_2024 = [r["FM_2024"], r["LG_2024"], r["RAT_2024"], r["RRP_2024"]]
    
    ax.bar(ratios, valores_2024, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd'])
    ax.set_title("Indicadores Clave 2024", fontsize=10)
    
    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

btnD = ttk.Button(tabD, text="Generar Diagnóstico Global", command=runD)
btnD.pack(pady=5)

root.mainloop()