import tkinter as tk
from tkinter import ttk, messagebox

# ---------------- INTERPRETACIONES AUTOMÁTICAS ----------------
def interpretar(categoria, ratio, valor):
    if categoria == "Liquidez":
        if ratio == "Razón corriente":
            if valor < 1:
                return "Insuficiente liquidez: la empresa no puede cubrir sus deudas a corto plazo."
            elif 1 <= valor <= 2:
                return "Liquidez aceptable: la empresa puede cumplir con sus obligaciones."
            else:
                return "Liquidez excesiva: puede estar desaprovechando recursos."
        elif ratio == "Prueba ácida":
            if valor < 1:
                return "Riesgo de iliquidez sin inventarios."
            else:
                return "Buena liquidez sin depender del inventario."
        elif ratio == "Capital de trabajo":
            if valor < 0:
                return "Capital de trabajo negativo: riesgo de insolvencia."
            else:
                return "Capital de trabajo positivo: la empresa puede operar normalmente."

    elif categoria == "Actividad":
        if ratio == "Rotación de inventarios":
            if valor < 3:
                return "Rotación baja: exceso de inventario."
            elif 3 <= valor <= 6:
                return "Rotación eficiente."
            else:
                return "Rotación muy alta: riesgo de falta de stock."
        elif ratio == "Período promedio de inventarios":
            if valor > 120:
                return "Demora en la venta de inventarios."
            else:
                return "Gestión de inventarios adecuada."
        elif ratio == "Rotación de cuentas por cobrar":
            if valor < 4:
                return "Cobranza lenta: riesgo de morosidad."
            else:
                return "Buena gestión de cobranza."
        elif ratio == "Período promedio de cobro":
            if valor > 90:
                return "Demasiado tiempo para cobrar las ventas."
            else:
                return "Período de cobro saludable."
        elif ratio == "Rotación de activos totales":
            if valor < 1:
                return "Baja eficiencia en uso de activos."
            else:
                return "Buena eficiencia operativa."

    elif categoria == "Endeudamiento":
        if ratio == "Razón de endeudamiento":
            if valor > 0.6:
                return "Alto endeudamiento: dependencia de financiamiento externo."
            elif 0.4 <= valor <= 0.6:
                return "Nivel de endeudamiento moderado."
            else:
                return "Bajo endeudamiento: empresa conservadora."
        elif ratio == "Razón de endeudamiento patrimonial":
            if valor > 1:
                return "Más deuda que capital propio: riesgo financiero alto."
            else:
                return "Estructura patrimonial equilibrada."
        elif ratio == "Cobertura de intereses":
            if valor < 2:
                return "Capacidad limitada para cubrir los gastos financieros."
            else:
                return "Buena cobertura de intereses."

    elif categoria == "Rentabilidad":
        if ratio == "Margen neto":
            if valor < 0.05:
                return "Rentabilidad baja: control de costos deficiente."
            elif 0.05 <= valor <= 0.15:
                return "Rentabilidad adecuada."
            else:
                return "Excelente margen de ganancia."
        elif ratio == "ROA (rendimiento sobre activos)":
            if valor < 0.05:
                return "Poca eficiencia en el uso de activos."
            else:
                return "Buen rendimiento sobre activos."
        elif ratio == "ROE (rendimiento sobre patrimonio)":
            if valor < 0.10:
                return "Rentabilidad sobre patrimonio baja."
            elif 0.10 <= valor <= 0.20:
                return "Rentabilidad adecuada para los accionistas."
            else:
                return "Excelente retorno sobre la inversión."

    return "Sin referencia estándar disponible."

# ---------------- CALCULADORA DE RATIOS ----------------
def calcular_ratio():
    try:
        categoria = combo_categoria.get()
        ratio = combo_ratio.get()
        datos = {nombre: float(entry.get() or 0) for nombre, entry in entradas.items()}

        resultado = 0
        explicacion = ""

        # ---------------- LIQUIDEZ ----------------
        if categoria == "Liquidez":
            if ratio == "Razón corriente":
                resultado = datos["Activo corriente"] / datos["Pasivo corriente"]
                explicacion = f"Por cada Bs.1 de deuda a corto plazo, la empresa tiene {resultado:.2f} Bs. en activos corrientes."
            elif ratio == "Prueba ácida":
                resultado = (datos["Activo corriente"] - datos["Inventarios"]) / datos["Pasivo corriente"]
                explicacion = f"Excluyendo inventarios, tiene {resultado:.2f} Bs. líquidos por cada Bs.1 de deuda."
            elif ratio == "Capital de trabajo":
                resultado = datos["Activo corriente"] - datos["Pasivo corriente"]
                explicacion = f"Capital de trabajo disponible: {resultado:.2f} Bs."

        # ---------------- ACTIVIDAD ----------------
        elif categoria == "Actividad":
            if ratio == "Rotación de inventarios":
                resultado = datos["Costo de ventas"] / datos["Inventario promedio"]
                explicacion = f"El inventario rota {resultado:.2f} veces al año."
            elif ratio == "Período promedio de inventarios":
                rot = datos["Costo de ventas"] / datos["Inventario promedio"]
                resultado = 360 / rot
                explicacion = f"El inventario permanece {resultado:.2f} días en promedio."
            elif ratio == "Rotación de cuentas por cobrar":
                resultado = datos["Ventas netas"] / datos["Cuentas por cobrar promedio"]
                explicacion = f"Las cuentas por cobrar rotan {resultado:.2f} veces al año."
            elif ratio == "Período promedio de cobro":
                rot = datos["Ventas netas"] / datos["Cuentas por cobrar promedio"]
                resultado = 360 / rot
                explicacion = f"El período promedio de cobro es de {resultado:.2f} días."
            elif ratio == "Rotación de activos totales":
                resultado = datos["Ventas netas"] / datos["Activo total"]
                explicacion = f"Por cada Bs.1 invertido en activos, se generan {resultado:.2f} Bs. en ventas."

        # ---------------- ENDEUDAMIENTO ----------------
        elif categoria == "Endeudamiento":
            if ratio == "Razón de endeudamiento":
                resultado = datos["Pasivo total"] / datos["Activo total"]
                explicacion = f"El {resultado*100:.2f}% de los activos está financiado con deuda."
            elif ratio == "Razón de endeudamiento patrimonial":
                resultado = datos["Pasivo total"] / datos["Patrimonio"]
                explicacion = f"Por cada Bs.1 de capital propio, la empresa debe {resultado:.2f} Bs."
            elif ratio == "Cobertura de intereses":
                resultado = datos["UAII"] / datos["Gastos por intereses"]
                explicacion = f"La empresa puede cubrir {resultado:.2f} veces sus intereses."

        # ---------------- RENTABILIDAD ----------------
        elif categoria == "Rentabilidad":
            if ratio == "Margen neto":
                resultado = datos["Utilidad neta"] / datos["Ventas netas"]
                explicacion = f"Margen neto del {resultado*100:.2f}%."
            elif ratio == "ROA (rendimiento sobre activos)":
                resultado = datos["Utilidad neta"] / datos["Activo total"]
                explicacion = f"ROA del {resultado*100:.2f}%."
            elif ratio == "ROE (rendimiento sobre patrimonio)":
                resultado = datos["Utilidad neta"] / datos["Patrimonio"]
                explicacion = f"ROE del {resultado*100:.2f}%."

        interpretacion = interpretar(categoria, ratio, resultado)
        messagebox.showinfo("Resultado",
                            f"{ratio} = {resultado:.2f}\n\n{explicacion}\n\n📊 Interpretación:\n{interpretacion}")

    except ZeroDivisionError:
        messagebox.showerror("Error", "No se puede dividir entre cero.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error: {e}")

# ---------------- INTERFAZ GRAFICA ----------------
ventana = tk.Tk()
ventana.title("Calculadora de Ratios Financieros")
ventana.geometry("550x700")

categorias = {
    "Liquidez": ["Razón corriente", "Prueba ácida", "Capital de trabajo"],
    "Actividad": ["Rotación de inventarios", "Período promedio de inventarios", "Rotación de cuentas por cobrar",
                  "Período promedio de cobro", "Rotación de activos totales"],
    "Endeudamiento": ["Razón de endeudamiento", "Razón de endeudamiento patrimonial", "Cobertura de intereses"],
    "Rentabilidad": ["Margen neto", "ROA (rendimiento sobre activos)", "ROE (rendimiento sobre patrimonio)"]
}

# Selección de categoría y ratio
ttk.Label(ventana, text="Seleccione la categoría:").pack(pady=5)
combo_categoria = ttk.Combobox(ventana, values=list(categorias.keys()))
combo_categoria.pack()

ttk.Label(ventana, text="Seleccione el ratio a calcular:").pack(pady=5)
combo_ratio = ttk.Combobox(ventana)
combo_ratio.pack()

def actualizar_ratios(event):
    categoria = combo_categoria.get()
    combo_ratio['values'] = categorias.get(categoria, [])
    combo_ratio.set("")

combo_categoria.bind("<<ComboboxSelected>>", actualizar_ratios)

# Campos de entrada
ttk.Label(ventana, text="\nIngrese los datos necesarios:").pack()
entradas = {}
campos = ["Activo corriente", "Pasivo corriente", "Inventarios", "Inventario promedio", "Costo de ventas",
          "Ventas netas", "Cuentas por cobrar promedio", "Activo total", "Pasivo total", "Patrimonio",
          "UAII", "Gastos por intereses", "Utilidad neta"]

for campo in campos:
    ttk.Label(ventana, text=campo + ":").pack()
    entry = ttk.Entry(ventana)
    entry.pack()
    entradas[campo] = entry

# Botón de cálculo
ttk.Button(ventana, text="Calcular", command=calcular_ratio).pack(pady=20)

ventana.mainloop()
