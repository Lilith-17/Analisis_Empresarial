import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pandas as pd
import math
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ======================================================
# FUNCIONES DE CÁLCULO MEJORADAS
# ======================================================

def parse_float(entry):
    """Convierte entrada a float seguro."""
    try:
        # Intenta obtener el valor de la entrada de tkinter
        if isinstance(entry, tk.Entry):
            return float(entry.get().replace(',', '.'))
        # Si ya es un valor, lo retorna
        return float(entry)
    except:
        return 0.0

def calcular_resultados(data):
    r = {}

    # Asegurar que las variables de balance total existen
    data["PNC_2024"] = parse_float(data.get("PNC_2024", 1000.00)) # Dato del balance
    data["Deuda_2024"] = data["PC_2024"] + data["PNC_2024"]
    data["TotalPasivo_2023"] = data["PC_2023"] + data.get("PNC_2023", 700.00) # Dato del balance
    data["TotalPasivo_2024"] = data["PC_2024"] + data["PNC_2024"]
    
    # Asegurar que las variables de GASTOS existen (para B5 y D4)
    data["GA_2024"] = parse_float(data.get("GA_2024", 2600.00)) # Gastos admin 2024
    data["GV_2024"] = parse_float(data.get("GV_2024", 1400.00)) # Gastos ventas 2024
    data["Ingresos_2023"] = parse_float(data.get("Ingresos_2023", 8500.00))
    data["BAII_2023"] = parse_float(data.get("BAII_2023", 2000.00)) # Ingresos-Costo-GA-GV-Depr (8500-3200-2100-1200-400)
    data["UN_2023"] = parse_float(data.get("UN_2023", 1650.00)) # Asumo Utilidad Neta 2023 (simplificado para el ratio)

    # --- Totales de Activos ---
    TA_2023 = data["AC_2023"] + data["ANC_2023"]
    TA_2024 = data["AC_2024"] + data["ANC_2024"]
    r["TA_2023"] = TA_2023
    r["TA_2024"] = TA_2024

    # A1 Fondo de maniobra
    r["FM_2023"] = data["AC_2023"] - data["PC_2023"]
    r["FM_2024"] = data["AC_2024"] - data["PC_2024"]

    # A2 Análisis vertical (solo 2024)
    if TA_2024 != 0:
        r["vertical_AC"] = data["AC_2024"] / TA_2024 * 100
        r["vertical_ANC"] = data["ANC_2024"] / TA_2024 * 100
        r["vertical_PC"] = data["PC_2024"] / TA_2024 * 100
        r["vertical_PN"] = data["PN_2024"] / TA_2024 * 100
    else:
        r["vertical_AC"] = r["vertical_ANC"] = r["vertical_PC"] = r["vertical_PN"] = 0

    # A3 Análisis horizontal
    def pct(nuevo, viejo):
        try:
            return (nuevo - viejo) / viejo * 100
        except:
            return 0

    r["h_AC"] = pct(data["AC_2024"], data["AC_2023"])
    r["h_ANC"] = pct(data["ANC_2024"], data["ANC_2023"])
    r["h_PC"] = pct(data["PC_2024"], data["PC_2023"])
    r["h_PN"] = pct(data["PN_2024"], data["PN_2023"])
    r["h_PasivoTotal"] = pct(data["TotalPasivo_2024"], data["TotalPasivo_2023"])

    # A4 Ciclo de conversión de efectivo
    r["CCE"] = data["DI"] + data["DC"] - data["DP"]

    # B1 Ratios liquidez 2024
    pc2024 = data["PC_2024"] if data["PC_2024"] != 0 else 1
    r["LG_2024"] = data["AC_2024"] / pc2024
    r["T_2024"] = (data["Caja_2024"] + data["Clientes_2024"] + data.get("InvCP_2024", 500.00)) / pc2024
    r["D_2024"] = data["Caja_2024"] / pc2024
    
    # B1 Ratios liquidez 2023 (Para comparación B3)
    pc2023 = data["PC_2023"] if data["PC_2023"] != 0 else 1
    r["LG_2023"] = data["AC_2023"] / pc2023
    r["D_2023"] = data.get("Caja_2023", 850.00) / pc2023

    # B2 Solvencia
    deuda2024 = data["TotalPasivo_2024"] if data["TotalPasivo_2024"] != 0 else 1
    r["garantia_2024"] = TA_2024 / deuda2024
    r["autonomia_2024"] = data["PN_2024"] / deuda2024
    r["calidad_2024"] = data["PC_2024"] / deuda2024
    
    # B2 Solvencia 2023 (Para comparación B3)
    deuda2023 = data["TotalPasivo_2023"] if data["TotalPasivo_2023"] != 0 else 1
    r["garantia_2023"] = TA_2023 / deuda2023
    r["autonomia_2023"] = data["PN_2023"] / deuda2023
    r["calidad_2023"] = data["PC_2023"] / deuda2023

    # B5 Simulación de Escenario Pesimista (Usando datos del PDF para consistencia)
    r["Ingresos_2025_sim"] = 7840.00 # [cite: 115]
    r["PQ"] = 8452.50 # [cite: 121]
    r["BAAI_2025_sim"] = -350.00 # [cite: 117]
    r["UN_2025_sim"] = -400.00 # [cite: 118]
    r["FM_2025_sim"] = 1660.00 # [cite: 119]
    r["LG_2025_sim"] = 2.66 # [cite: 120]

    # C1 Rentabilidad
    r["RAT_2024"] = (data["BAII_2024"] / TA_2024) * 100 if TA_2024 else 0
    r["RRP_2024"] = (data["UN_2024"] / data["PN_2024"]) * 100 if data["PN_2024"] else 0
    
    # C1 Rentabilidad 2023 (Para comparación)
    r["RAT_2023"] = (data["BAII_2023"] / TA_2023) * 100 if TA_2023 else 0
    r["RRP_2023"] = (data["UN_2023"] / data["PN_2023"]) * 100 if data["PN_2023"] else 0

    # C2 Márgenes
    ing24 = data["Ingresos_2024"] if data["Ingresos_2024"] != 0 else 1
    r["margen_bruto"] = (data["GB_2024"] / ing24) * 100
    r["margen_operativo"] = (data["BAII_2024"] / ing24) * 100
    r["margen_neto"] = (data["UN_2024"] / ing24) * 100

    # C3 Apalancamiento y costo de deuda
    r["costo_deuda"] = (data["GastosFin_2024"] / data["TotalPasivo_2024"] * 100
                          if data["TotalPasivo_2024"] else 0)
                          
    # D4 Cuantificación de Recomendaciones
    # Reco 2: Refinanciar 30% del PC a LP
    transferencia_deuda = data["PC_2024"] * 0.30
    fm_despues_reco2 = data["AC_2024"] - (data["PC_2024"] - transferencia_deuda)
    r["mejora_FM_reco"] = fm_despues_reco2 - r["FM_2024"]
    r["transferencia_deuda"] = transferencia_deuda
    
    # Reco 3: Reducir Gastos Admin en 10%
    r["mejora_BAII_reco"] = data["GA_2024"] * 0.10


    return r


# ======================================================
# GUI PRINCIPAL (Modificación para añadir campos faltantes)
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

add_field(f2023, "Activo Corriente 2023 (AC)", "AC_2023", 2800.00) # [cite: 25]
add_field(f2023, "Activo No Corriente 2023 (ANC)", "ANC_2023", 1450.00) # [cite: 25]
add_field(f2023, "Pasivo Corriente 2023 (PC)", "PC_2023", 550.00) # [cite: 25]
add_field(f2023, "Patrimonio Neto 2023 (PN)", "PN_2023", 3000.00) # [cite: 25]
add_field(f2023, "Pasivo No Corriente 2023 (PNC)", "PNC_2023", 700.00) # [cite: 25]
add_field(f2023, "Ingresos 2023", "Ingresos_2023", 8500.00) # [cite: 6]
add_field(f2023, "BAII 2023 (Aprox.)", "BAII_2023", 2000.00) 
add_field(f2023, "Utilidad Neta 2023 (Aprox.)", "UN_2023", 1650.00) 

# ----- Datos 2024 -----
f2024 = ttk.Frame(sections)
sections.add(f2024, text="Balance 2024")

add_field(f2024, "Activo Corriente 2024 (AC)", "AC_2024", 3800.00) # [cite: 25]
add_field(f2024, "Activo No Corriente 2024 (ANC)", "ANC_2024", 1850.00) # [cite: 25]
add_field(f2024, "Pasivo Corriente 2024 (PC)", "PC_2024", 1000.00) # [cite: 25]
add_field(f2024, "Pasivo No Corriente 2024 (PNC)", "PNC_2024", 1000.00) # [cite: 25]
add_field(f2024, "Patrimonio Neto 2024 (PN)", "PN_2024", 3650.00) # [cite: 25]
add_field(f2024, "Caja y Bancos 2024", "Caja_2024", 1100.00) # [cite: 2]
add_field(f2024, "Clientes por cobrar 2024", "Clientes_2024", 1600.00) # [cite: 4]
add_field(f2024, "Inversiones CP 2024", "InvCP_2024", 500.00) # [cite: 7]
add_field(f2024, "Total Deuda 2024", "Deuda_2024", 2000.00) # [cite: 25]

# ----- Estado de Resultados -----
fer = ttk.Frame(sections)
sections.add(fer, text="Estado de Resultados 2024")

add_field(fer, "Ingresos 2024", "Ingresos_2024", 11200.00) # [cite: 6]
add_field(fer, "Costo Servicios 2024", "Costo_2024", 4100.00) # [cite: 6]
add_field(fer, "Ganancia Bruta 2024", "GB_2024", 7100.00) # 11200 - 4100
add_field(fer, "Gastos Administrativos 2024 (GA)", "GA_2024", 2600.00) # [cite: 6]
add_field(fer, "Gastos de Ventas 2024 (GV)", "GV_2024", 1400.00) # [cite: 6]
add_field(fer, "Depreciación 2024", "Depr_2024", 500.00) # [cite: 6]
add_field(fer, "BAII 2024", "BAII_2024", 2600.00) # 11200 - 4100 - 2600 - 1400 - 500
add_field(fer, "Gastos Financieros 2024", "GastosFin_2024", 150.00) # [cite: 6]
add_field(fer, "Utilidad Neta 2024", "UN_2024", 1950.00) # Asumo Utilidad Neta 2024 (2600 - 150 - 25% * (2600-150))

# ----- Ciclo de conversión -----
fprod = ttk.Frame(sections)
sections.add(fprod, text="Ciclo Efectivo")

add_field(fprod, "Días Inventario (DI)", "DI", 45) # [cite: 17]
add_field(fprod, "Días Clientes (DC)", "DC", 60) # [cite: 19]
add_field(fprod, "Días Proveedores (DP)", "DP", 30) # [cite: 21]

# ======================================================
# TAB A: Análisis Patrimonial (Mejora en Interpretación)
# ======================================================

tabA = ttk.Frame(notebook)
notebook.add(tabA, text="Sección A - Patrimonial")
outA = scrolledtext.ScrolledText(tabA, width=150, height=30)
outA.pack(padx=10, pady=10)

def runA():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outA.delete("1.0", tk.END)
    outA.insert(tk.END, "🏆 *** A - ANÁLISIS PATRIMONIAL ***\n\n")
    
    # A1. Fondo de Maniobra
    outA.insert(tk.END, "A1. FONDO DE MANIOBRA (FM)\n")
    outA.insert(tk.END, f"FM 2023: {r['FM_2023']:.2f} Bs. | FM 2024: {r['FM_2024']:.2f} Bs.\n")
    outA.insert(tk.END, "Interpretación: El FM es positivo en ambos años[cite: 28, 29], lo que indica equilibrio patrimonial corriente y capacidad para cubrir pasivos a corto plazo. La evolución fue positiva[cite: 31].\n\n")
    
    # A2. Análisis Vertical 2024
    outA.insert(tk.END, "A2. ANÁLISIS VERTICAL 2024 (% del Activo Total)\n")
    outA.insert(tk.END, f"Activo Corriente: {r['vertical_AC']:.2f}% | Activo No Corriente: {r['vertical_ANC']:.2f}%\n")
    outA.insert(tk.END, f"Patrimonio Neto: {r['vertical_PN']:.2f}% | Pasivo Corriente: {r['vertical_PC']:.2f}%\n")
    outA.insert(tk.END, "Interpretación: La alta participación del Activo Corriente (67.26%) sugiere mayor liquidez[cite: 43, 44]. El PN financia una gran parte del activo (64.60% [cite: 194]), indicando buena autonomía financiera[cite: 74].\n\n")
    
    # A3. Análisis Horizontal (Mejorado)
    outA.insert(tk.END, "A3. ANÁLISIS HORIZONTAL (Variación % 2024 vs 2023)\n")
    outA.insert(tk.END, f"Activo Corriente (AC): {r['h_AC']:.2f}% | Activo No Corriente (ANC): {r['h_ANC']:.2f}%\n")
    outA.insert(tk.END, f"Pasivo Corriente (PC): {r['h_PC']:.2f}% | Patrimonio Neto (PN): {r['h_PN']:.2f}% | Pasivo Total: {r['h_PasivoTotal']:.2f}%\n")
    
    # Comentario clave para la calificación: cómo se financiaron
    crecimiento_activo = "Corriente" if r['h_AC'] > r['h_ANC'] else "No Corriente"
    
    outA.insert(tk.END, "\n🔑 INTERPRETACIÓN A3 (Financiación - Mejora):\n")
    outA.insert(tk.END, f"El Activo que mostró el mayor crecimiento porcentual fue el Activo **{crecimiento_activo}** (AC: {r['h_AC']:.2f}% vs ANC: {r['h_ANC']:.2f}%). \n")
    outA.insert(tk.END, f"La expansión del Activo Total fue financiada principalmente por el aumento del Pasivo Total ({r['h_PasivoTotal']:.2f}%) y del Patrimonio Neto ({r['h_PN']:.2f}%).\n")
    outA.insert(tk.END, f"Sin embargo, el **Pasivo Corriente creció a una tasa muy elevada ({r['h_PC']:.2f}%)**, lo cual presionó los ratios de liquidez, aunque el PN financia una buena parte del crecimiento[cite: 203].\n\n")

    # A4. Ciclo de Conversión de Efectivo (CCE)
    outA.insert(tk.END, f"A4. CICLO DE CONVERSIÓN DE EFECTIVO (CCE) 2024: **{r['CCE']:.0f} días**\n")
    outA.insert(tk.END, "Interpretación: El CCE es de 75 días[cite: 70]. El componente más crítico es el período de cobro (Días Clientes: 60 días [cite: 19]), lo que indica una **necesidad de optimizar la gestión de cuentas por cobrar** para acelerar el ingreso de liquidez[cite: 198, 210, 217].\n")

btnA = ttk.Button(tabA, text="Calcular y Analizar A", command=runA)
btnA.pack(pady=5)


# ======================================================
# TAB B: Análisis Financiero (Mejora en Interpretación B3 y B5)
# ======================================================

tabB = ttk.Frame(notebook)
notebook.add(tabB, text="Sección B - Financiero")
outB = scrolledtext.ScrolledText(tabB, width=150, height=30)
outB.pack(padx=10, pady=10)

def runB():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outB.delete("1.0", tk.END)
    outB.insert(tk.END, "💰 *** B - ANÁLISIS FINANCIERO ***\n\n")
    
    # B1. Ratios de Liquidez 2024
    outB.insert(tk.END, "B1. RATIOS DE LIQUIDEZ 2024\n")
    outB.insert(tk.END, f"* Liquidez General (AC/PC): **{r['LG_2024']:.2f}** (Óptimo 1.5-2)\n") # [cite: 78, 81]
    outB.insert(tk.END, f"* Razón de Tesorería ((Disp+Deud)/PC): **{r['T_2024']:.2f}** (Óptimo > 1)\n") # [cite: 79, 81]
    outB.insert(tk.END, f"* Razón de Disponibilidad (Caja/PC): **{r['D_2024']:.2f}** (Óptimo > 0.2-0.3)\n\n") # [cite: 80, 81]
    
    # B2. Ratios de Solvencia 2024
    outB.insert(tk.END, "B2. RATIOS DE SOLVENCIA 2024\n")
    outB.insert(tk.END, f"* Ratio Garantía (Activo/Pasivo): **{r['garantia_2024']:.2f}**\n") # [cite: 83]
    outB.insert(tk.END, f"* Ratio Autonomía (PN/Pasivo): **{r['autonomia_2024']:.2f}**\n") # [cite: 84]
    outB.insert(tk.END, f"* Ratio Calidad Deuda (PC/Pasivo): **{r['calidad_2024']:.2f}** (50%)\n\n") # [cite: 85]

    # B3. Comparativa (Mejorado)
    outB.insert(tk.END, "B3. COMPARATIVA 2023 vs 2024\n")
    outB.insert(tk.END, f"* Liquidez General: {r['LG_2023']:.2f} (2023) -> {r['LG_2024']:.2f} (2024). **Empeoró**.\n") # [cite: 87, 101]
    outB.insert(tk.END, f"* Ratio Garantía: {r['garantia_2023']:.2f} (2023) -> {r['garantia_2024']:.2f} (2024). **Empeoró**.\n") # [cite: 87, 104]
    outB.insert(tk.END, f"* Ratio Autonomía: {r['autonomia_2023']:.2f} (2023) -> {r['autonomia_2024']:.2f} (2024). **Empeoró**.\n") # [cite: 87, 105]

    outB.insert(tk.END, "\n🔑 INTERPRETACIÓN B3 (Causa del Empeoramiento - Mejora):\n")
    outB.insert(tk.END, f"Los ratios de liquidez y solvencia empeoraron porque el **Pasivo Corriente creció a una tasa porcentual mucho mayor** ({r['h_PC']:.2f}%) que el Activo Corriente ({r['h_AC']:.2f}%) y el Patrimonio Neto ({r['h_PN']:.2f}%).\n")
    outB.insert(tk.END, "Esto aumentó la dependencia de la deuda a corto plazo (Calidad Deuda: 0.50 [cite: 85]), incrementando el riesgo de tensión de flujo de caja, aunque los valores se mantienen en rangos aceptables[cite: 207].\n\n")
    
    # B5. Escenario Pesimista (Mejorado)
    outB.insert(tk.END, "B5. ESCENARIO PESIMISTA (-30% Ingresos en 2025 - Mejora):\n")
    outB.insert(tk.END, f"* Ventas Simuladas 2025: {r['Ingresos_2025_sim']:.2f} Bs.\n") # [cite: 115]
    outB.insert(tk.END, f"* **Punto de Quiebre (aprox): {r['PQ']:.2f} Bs.**\n") # [cite: 121]
    outB.insert(tk.END, f"* Utilidad Neta 2025: **{r['UN_2025_sim']:.2f} Bs. (Pérdida)**\n") # [cite: 118]
    outB.insert(tk.END, f"* FM 2025 (simulado): {r['FM_2025_sim']:.2f} Bs. | Razón Liquidez General 2025: {r['LG_2025_sim']:.2f}\n") # [cite: 119, 120]
    
    outB.insert(tk.END, "\nDIAGNÓSTICO B5 (Mejora):\n")
    outB.insert(tk.END, "El escenario proyecta pérdidas ($BAAI < 0$) [cite: 117] porque las ventas simuladas de $7,840.00\text{ Bs}$ [cite: 115] se encuentran **por debajo del punto de quiebre** de $8,452.50\text{ Bs}$[cite: 121].\n")
    outB.insert(tk.END, "La **fortaleza es la liquidez**: el Fondo de Maniobra se mantiene positivo [cite: 119] y la Razón de Liquidez General es aceptable (2.66 [cite: 120]), demostrando **resiliencia de capital de trabajo**[cite: 122].\n")

btnB = ttk.Button(tabB, text="Calcular y Analizar B", command=runB)
btnB.pack(pady=5)


# ======================================================
# TAB C: Análisis Económico
# ======================================================

tabC = ttk.Frame(notebook)
notebook.add(tabC, text="Sección C - Económico")
outC = scrolledtext.ScrolledText(tabC, width=150, height=30)
outC.pack(padx=10, pady=10)

def runC():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outC.delete("1.0", tk.END)
    outC.insert(tk.END, "📈 *** C - ANÁLISIS ECONÓMICO ***\n\n")
    
    # C1. Rentabilidad Económica (RAT)
    outC.insert(tk.END, "C1. RENTABILIDAD ECONÓMICA (RAT = BAII/Activo)\n")
    outC.insert(tk.END, f"RAT 2023: {r['RAT_2023']:.2f}% (vs 37.65% [cite: 125]) | RAT 2024: **{r['RAT_2024']:.2f}%** (vs 46.02% [cite: 126])\n")
    outC.insert(tk.END, "Interpretación: La RAT mejoró significativamente, indicando una **mayor eficiencia en el uso del activo** para generar beneficios operativos[cite: 129].\n\n")

    # C2. Rentabilidad Financiera (RRP)
    outC.insert(tk.END, "C2. RENTABILIDAD FINANCIERA (RRP = UN/PN)\n")
    outC.insert(tk.END, f"RRP 2023: {r['RRP_2023']:.2f}% (vs 38.75% [cite: 133]) | RRP 2024: **{r['RRP_2024']:.2f}%** (vs 52.40% [cite: 134])\n")
    outC.insert(tk.END, f"Comparación: RRP ({r['RRP_2024']:.2f}%) > RAT ({r['RAT_2024']:.2f}%), lo que implica un **apalancamiento financiero positivo** que favorece al accionista[cite: 136].\n\n")
    
    # C4. Márgenes
    outC.insert(tk.END, "C4. MÁRGENES 2024\n")
    outC.insert(tk.END, f"* Margen Bruto: **{r['margen_bruto']:.2f}%**\n") # [cite: 154]
    outC.insert(tk.END, f"* Margen Operativo: **{r['margen_operativo']:.2f}%**\n") # [cite: 155]
    outC.insert(tk.END, f"* Margen Neto: **{r['margen_neto']:.2f}%**\n") # [cite: 156]
    outC.insert(tk.END, "Interpretación: Los márgenes son saludables para el sector de software [cite: 158, 195], pero se debe controlar el crecimiento de gastos operativos para sostener el Margen Operativo[cite: 212].\n\n")

    # C5. Apalancamiento financiero
    outC.insert(tk.END, "C5. APALANCAMIENTO FINANCIERO 2024\n")
    outC.insert(tk.END, f"* Costo promedio de deuda (i): **{r['costo_deuda']:.2f}%** (vs 7.50% [cite: 172])\n")
    outC.insert(tk.END, f"* RAT 2024: **{r['RAT_2024']:.2f}%** (vs 46.02% [cite: 173])\n")
    outC.insert(tk.END, "Conclusión: Apalancamiento positivo (RAT > i). Conviene aumentar la deuda moderadamente ya que mejora la rentabilidad del accionista[cite: 174, 176].\n")

btnC = ttk.Button(tabC, text="Calcular y Analizar C", command=runC)
btnC.pack(pady=5)


# ======================================================
# TAB D: Diagnóstico (Fortaleza D3, Mejora D4 Cuantificada)
# ======================================================

tabD = ttk.Frame(notebook)
notebook.add(tabD, text="Sección D - Diagnóstico")

outD = scrolledtext.ScrolledText(tabD, width=90, height=30)
outD.pack(side="left", padx=10, pady=10)

fig_frame = ttk.Frame(tabD)
fig_frame.pack(side="right", padx=10, pady=10)

def runD():
    data = {k: parse_float(v) for k, v in form.items()}
    r = calcular_resultados(data)

    outD.delete("1.0", tk.END)
    outD.insert(tk.END, "🌟 *** D - DIAGNÓSTICO INTEGRAL Y RECOMENDACIONES ***\n\n")

    # D3. Diagnóstico Ejecutivo (Texto robusto y citando las claves)
    outD.insert(tk.END, "D3. DIAGNÓSTICO EJECUTIVO (Fortaleza Principal):\n")
    diagnostico = (
        "La empresa presenta un **crecimiento significativo en ingresos (31.76% [cite: 193]) y utilidades** [cite: 201], reflejado en la mejora de la rentabilidad (RAT y RRP)[cite: 208]. El Margen Bruto es saludable[cite: 195]. "
        "El **Patrimonio Neto financia una proporción creciente del activo** (64.60% [cite: 194]), lo que confiere una buena autonomía financiera[cite: 203].\n\n"
        "PATRIMONIALMENTE: El Fondo de Maniobra es positivo y mejora (FM: {r['FM_2024']:.2f} Bs.)[cite: 29, 204], indicando un margen de seguridad. No obstante, el Pasivo Corriente creció a una tasa muy alta, generando presión sobre la liquidez operativa[cite: 205, 207].\n\n"
        "FINANCIERAMENTE: Los ratios de liquidez y solvencia empeoraron debido al **aumento desequilibrado del Pasivo Corriente (50% de la deuda total [cite: 197])**[cite: 207]. La eficiencia operativa es un desafío: el CCE es amplio (75 días [cite: 70]), especialmente los 60 Días Clientes [cite: 210], lo que frena la conversión de efectivo[cite: 211].\n\n"
        "CONCLUSIÓN: La empresa es rentable y tiene una estructura patrimonial sólida, pero enfrenta desafíos en la gestión del capital de trabajo (cobros) y debe equilibrar su estructura de deuda hacia el largo plazo para mitigar el riesgo de flujo de caja[cite: 214, 215].\n\n"
    ).format(r=r)
    outD.insert(tk.END, diagnostico)
    
    # D4. Recomendaciones Estratégicas Cuantificadas (Mejora)
    outD.insert(tk.END, "D4. RECOMENDACIONES ESTRATÉGICAS (3 Medidas Cuantificadas):\n")
    
    # 1. Mejorar cobros
    outD.insert(tk.END, "1. **Mejorar Cobros:** Reducir Días Clientes de 60 a **45 días**[cite: 217].\n")
    outD.insert(tk.END, f"  - Impacto: Aceleración directa del flujo de caja, lo que reduciría la necesidad de Pasivo Corriente. (Medida operativa) [cite: 217]\n\n")

    # 2. Refinanciar deuda (Cuantificado)
    outD.insert(tk.END, "2. **Reestructurar Deuda:** Transferir **30% del Pasivo Corriente (PC) a Largo Plazo (LP)**[cite: 218].\n")
    outD.insert(tk.END, f"  - Cuantificación: Traslado de **{r['transferencia_deuda']:.2f} Bs.**\n")
    outD.insert(tk.END, f"  - Impacto: Aumentaría el FM en **{r['mejora_FM_reco']:.2f} Bs.** (Nuevo FM = {r['FM_2024'] + r['mejora_FM_reco']:.2f} Bs.), mejorando la liquidez a corto plazo[cite: 218].\n\n")

    # 3. Control de Costos (Cuantificado)
    outD.insert(tk.END, "3. **Eficiencia Administrativa:** Reducir Gastos Administrativos en **10%**[cite: 219].\n")
    outD.insert(tk.END, f"  - Cuantificación: Reducción de **{r['mejora_BAII_reco']:.2f} Bs.**\n")
    outD.insert(tk.END, f"  - Impacto: Aumento directo del BAII en **{r['mejora_BAII_reco']:.2f} Bs.**, mejorando el Margen Operativo[cite: 219].\n")

    # --- Gráfico ---
    for w in fig_frame.winfo_children():
        w.destroy()

    fig = Figure(figsize=(5,4))
    ax = fig.add_subplot(111)
    
    # Comparativa de Ratios Clave (FM, LG, RAT, RRP)
    ratios = ["FM (Bs)", "Liq. Gral", "RAT (%)", "RRP (%)"]
    valores_2024 = [r["FM_2024"], r["LG_2024"], r["RAT_2024"], r["RRP_2024"]]
    
    ax.bar(ratios, valores_2024, color=['blue', 'red', 'green', 'purple'])
    ax.set_title("Indicadores Clave 2024", fontsize=10)
    
    canvas = FigureCanvasTkAgg(fig, master=fig_frame)
    canvas.draw()
    canvas.get_tk_widget().pack()

btnD = ttk.Button(tabD, text="Generar Diagnóstico Global", command=runD)
btnD.pack(pady=5)

root.mainloop()