# ===============================
# ANALISIS INTEGRAL DE EMPRESA 2023-2024
# Secciones A, B, C y D
# ===============================

# 1️⃣ Importar librerías
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

sns.set(style="whitegrid")

# -------------------------------
# 2️⃣ Datos iniciales
# -------------------------------

# Activos
activos = pd.DataFrame({
    "Cuenta": ["Activo Corriente", "Activo No Corriente", "Total Activo"],
    "2023": [2800, 1450, 4250],
    "2024": [3800, 1850, 5650]
})

# Pasivos y Patrimonio
pasivos_pn = pd.DataFrame({
    "Cuenta": ["Pasivo Corriente", "Pasivo No Corriente", "Patrimonio Neto", "Total PN + Pasivo"],
    "2023": [550, 700, 3000, 4250],
    "2024": [1000, 1000, 3650, 5650]
})

# Estado de Resultados
er = pd.DataFrame({
    "Concepto": ["Ingresos", "Costo Servicios", "Ganancia Bruta", "Gastos Adm", "Gastos Ventas", 
                 "Depreciacion", "BAII", "Gastos Financieros", "Otros Ingresos", "UN"],
    "2023": [8500, 3200, 5300, 2100, 1200, 400, 1600, 100, 50, 1162],
    "2024": [11200, 4100, 7100, 2600, 1400, 500, 2600, 150, 100, 1912]
})

# Otros datos
caja = {"2023": 850, "2024": 1100}
clientes = {"2023": 1200, "2024": 1600}
inventarios = {"2023": 450, "2024": 600}
deuda_total = {"2023": 1250, "2024": 2000}
deuda_corriente = {"2023": 550, "2024": 1000}
pn = {"2023": 3000, "2024": 3650}

# -------------------------------
# 3️⃣ SECCION A: ANALISIS PATRIMONIAL
# -------------------------------

# Fondo de Maniobra
def fondo_maniobra(ac, pc):
    return ac - pc

FM_2023 = fondo_maniobra(activos.loc[0,"2023"], pasivos_pn.loc[0,"2023"])
FM_2024 = fondo_maniobra(activos.loc[0,"2024"], pasivos_pn.loc[0,"2024"])
print("Fondo de Maniobra 2023:", FM_2023)
print("Fondo de Maniobra 2024:", FM_2024)

# Analisis Vertical 2024
vertical_2024 = activos["2024"]/activos.loc[2,"2024"]*100
print("\nAnalisis Vertical 2024 (%)")
print(vertical_2024)

# Analisis Horizontal (2024 vs 2023)
horizontal_activos = (activos["2024"] - activos["2023"])/activos["2023"]*100
print("\nAnalisis Horizontal 2024 vs 2023 (%)")
print(horizontal_activos)

# Ciclo de Conversion de Efectivo
dias_inventario = 45
dias_clientes = 60
dias_proveedores = 30

CCE_2024 = dias_inventario + dias_clientes - dias_proveedores
print("\nCiclo de Conversion de Efectivo 2024:", CCE_2024, "días")

# -------------------------------
# 4️⃣ SECCION B: ANALISIS FINANCIERO
# -------------------------------

# Ratios de Liquidez
def liquidez_general(ac, pc): return ac/pc
def tesoreria(caja, clientes, pc): return (caja + clientes)/pc
def disponibilidad(caja, pc): return caja/pc

LG_2023 = liquidez_general(activos.loc[0,"2023"], pasivos_pn.loc[0,"2023"])
LG_2024 = liquidez_general(activos.loc[0,"2024"], pasivos_pn.loc[0,"2024"])
T_2024 = tesoreria(caja["2024"], clientes["2024"], pasivos_pn.loc[0,"2024"])
D_2024 = disponibilidad(caja["2024"], pasivos_pn.loc[0,"2024"])

print("\nLiquidez General 2023:", LG_2023)
print("Liquidez General 2024:", LG_2024)
print("Razón Tesorería 2024:", T_2024)
print("Razón Disponibilidad 2024:", D_2024)

# Ratios de Solvencia
def garantia(at, pasivo): return at/pasivo
def autonomia(pn, pasivo): return pn/pasivo
def calidad_deuda(pc, pasivo): return pc/pasivo

garantia_2024 = garantia(activos.loc[2,"2024"], deuda_total["2024"])
autonomia_2024 = autonomia(pn["2024"], deuda_total["2024"])
calidad_2024 = calidad_deuda(deuda_corriente["2024"], deuda_total["2024"])

print("\nRatio de Garantía 2024:", garantia_2024)
print("Ratio de Autonomía 2024:", autonomia_2024)
print("Ratio Calidad de Deuda 2024:", calidad_2024)

# Escenario pesimista 2025
ingresos_2025 = er.loc[0,"2024"]*0.7
gastos_fijos = ingresos_2025*0.6
ac_2025 = activos.loc[0,"2024"]*0.7
pc_2025 = pasivos_pn.loc[0,"2024"] + gastos_fijos
FM_2025 = ac_2025 - pc_2025
RL_2025 = ac_2025 / pc_2025
print("\nEscenario Pesimista 2025 -> FM:", FM_2025, "Liquidez:", RL_2025)

# -------------------------------
# 5️⃣ SECCION C: ANALISIS ECONOMICO / RENTABILIDAD
# -------------------------------

# RAT y RRP
def RAT(BAII, AT): return BAII/AT*100
def RRP(UN, PN): return UN/PN*100

RAT_2023 = RAT(er.loc[6,"2023"], activos.loc[2,"2023"])
RAT_2024 = RAT(er.loc[6,"2024"], activos.loc[2,"2024"])
RRP_2023 = RRP(er.loc[9,"2023"], pn["2023"])
RRP_2024 = RRP(er.loc[9,"2024"], pn["2024"])

print("\nRAT 2023:", RAT_2023, "%")
print("RAT 2024:", RAT_2024, "%")
print("RRP 2023:", RRP_2023, "%")
print("RRP 2024:", RRP_2024, "%")

# Analisis DuPont 2024
margen_neto = er.loc[9,"2024"]/er.loc[0,"2024"]
rotacion_activo = er.loc[0,"2024"]/activos.loc[2,"2024"]
apalancamiento = activos.loc[2,"2024"]/pn["2024"]
RRP_dupont = margen_neto * rotacion_activo * apalancamiento * 100
print("\nDuPont 2024 -> Margen Neto:", margen_neto*100, "%, Rotacion Activo:", rotacion_activo, 
      ", Apalancamiento:", apalancamiento, ", RRP Dupont:", RRP_dupont, "%")

# Margenes 2024
margen_bruto = er.loc[2,"2024"]/er.loc[0,"2024"]*100
margen_operativo = er.loc[6,"2024"]/er.loc[0,"2024"]*100
margen_neto_pct = er.loc[9,"2024"]/er.loc[0,"2024"]*100
print("\nMargen Bruto 2024:", margen_bruto, "%")
print("Margen Operativo 2024:", margen_operativo, "%")
print("Margen Neto 2024:", margen_neto_pct, "%")

# Apalancamiento financiero
costo_deuda = er.loc[7,"2024"]/deuda_total["2024"]*100
efecto_apalancamiento = RAT_2024 + (deuda_total["2024"]/pn["2024"])*(RAT_2024 - costo_deuda)
print("\nCosto deuda 2024:", costo_deuda, "%")
print("Efecto apalancamiento 2024:", efecto_apalancamiento, "%")

# -------------------------------
# 6️⃣ SECCION D: MATRIZ DE RATIOS Y DIAGNOSTICO
# -------------------------------

ratios_comparativos = pd.DataFrame({
    "Ratio": ["FM", "Liq. Gral.", "RAT", "RRP"],
    "2023": [FM_2023, LG_2023, RAT_2023, RRP_2023],
    "2024": [FM_2024, LG_2024, RAT_2024, RRP_2024],
    "Cambio": [FM_2024-FM_2023, LG_2024-LG_2023, RAT_2024-RAT_2023, RRP_2024-RRP_2023],
    "Interpretación": ["FM positivo y creciente",
                       "Muy buena, ligera disminución",
                       "Rentabilidad económica creciente",
                       "Rentabilidad financiera mayor por apalancamiento"]
})

print("\nMatriz Comparativa Ratios 2023 vs 2024")
print(ratios_comparativos)

# -------------------------------
# 7️⃣ Graficos
# -------------------------------

# FM 2023-2024
plt.figure(figsize=(8,5))
plt.bar(["FM 2023","FM 2024"], [FM_2023, FM_2024], color="skyblue")
plt.title("Fondo de Maniobra 2023-2024")
plt.ylabel("Bs")
plt.show()

# RAT y RRP 2023-2024
plt.figure(figsize=(8,5))
plt.bar(["RAT 2023","RAT 2024"], [RAT_2023, RAT_2024], color="green")
plt.bar(["RRP 2023","RRP 2024"], [RRP_2023, RRP_2024], color="orange", alpha=0.7)
plt.title("Rentabilidad Económica y Financiera 2023-2024")
plt.ylabel("%")
plt.show()
