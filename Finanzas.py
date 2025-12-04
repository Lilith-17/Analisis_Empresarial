import math

# ---------------- FUNCION PARA LEER TIEMPO ----------------
def leer_tiempo():
    print("\n--- INGRESO DE TIEMPO ---")
    años = int(input("Años: "))
    #semestres = int(input("Semestres: "))
    #cuatrimestres = int(input("Cuatrimestres: "))
    #bimestres = int(input("Bimestres: "))
    #trimestres = int(input("Trimestres: "))
    meses = int(input("Meses: "))
    semanas = int(input("Semanas: "))
    dias = int(input("Días: "))

    # Conversión a días (año comercial: 360 días)
    total_dias = (
        (años * 360) +
        #(semestres * 180) +
        #(cuatrimestres * 120) +
        #(bimestres * 60) +
        #(trimestres * 90) +
        (meses * 30) +
        (semanas * 7) +
        dias
    )

    tiempo_en_años = total_dias / 360
    print(f"Tiempo total convertido: {tiempo_en_años:.4f} años ({total_dias} días comerciales)")
    return tiempo_en_años

# ----------------- INTERÉS SIMPLE -----------------
def interes_simple_general():
    print("\n--- INTERÉS SIMPLE GENERAL ---")
    print("1. Calcular Valor Futuro (VF o Monto)")
    print("2. Calcular Valor Presente (VP o Capital)")
    print("3. Calcular Interés (I)")
    print("4. Calcular Tasa de interés (i)")
    print("5. Calcular Tiempo (n)")
    opcion = int(input("Elija una opción: "))

    if opcion == 1:  # VF = VP + I = VP * (1 + i*t)
        VP = float(input("Capital (VP o Principal): "))
        i = float(input("Tasa de interés (ej. 0.24 para 24% anual): "))
        t = leer_tiempo()
        VF = VP * (1 + i * t)
        print(f"Valor Futuro (Monto acumulado): {VF:.2f}")
        print(f"Interpretación: Si inviertes {VP}, en {t:.2f} años tendrás {VF:.2f}.")

    elif opcion == 2:  # VP = VF / (1 + i*t)
        VF = float(input("Valor Futuro (VF o Monto): "))
        i = float(input("Tasa de interés: "))
        t = leer_tiempo()
        VP = VF / (1 + i * t)
        print(f"Valor Presente (Capital): {VP:.2f}")

    elif opcion == 3:  # I = VP * i * t
        VP = float(input("Capital (VP): "))
        i = float(input("Tasa de interés: "))
        t = leer_tiempo()
        I = VP * i * t
        print(f"Interés generado: {I:.2f}")

    elif opcion == 4:  # i = I / (VP*t)
        I = float(input("Interés generado: "))
        VP = float(input("Capital (VP): "))
        t = leer_tiempo()
        i = I / (VP * t)
        print(f"Tasa de interés: {i:.4f} ({i*100:.2f}%)")

    elif opcion == 5:  # n = I / (VP*i)
        I = float(input("Interés generado: "))
        VP = float(input("Capital: "))
        i = float(input("Tasa de interés: "))
        t = I / (VP * i)
        print(f"Tiempo: {t:.2f} años")

# ----------------- INTERÉS SIMPLE CON VARIACIONES -----------------
def interes_simple_variado():
    print("\n--- INTERÉS SIMPLE CON VARIACIONES DE TASA ---")
    capital = float(input("Capital inicial: "))
    tramos = int(input("¿Cuántos tramos de tasa hay?: "))
    interes_total = 0

    for t in range(tramos):
        tasa = float(input(f"Tasa del tramo {t+1} (ej. 0.03 para 3% mensual, 0.24 para 24% anual): "))
        tiempo = float(input(f"Tiempo del tramo {t+1}: "))
        unidad = input("Unidad de tiempo (años, meses, días): ")
        tiempo_conv = leer_tiempo
        interes = capital * tasa * tiempo_conv
        interes_total += interes
        print(f"→ Tramo {t+1}: interés generado = {interes:.2f}")

    print(f"Interés acumulado total = {interes_total:.2f}")
    print(f"Interpretación: Con {tramos} variaciones de tasa, el capital generó {interes_total:.2f}.")
    
# ----------------- INTERÉS COMPUESTO -----------------
def interes_compuesto_general():
    print("\n--- INTERÉS COMPUESTO ---")
    print("1. Valor Futuro (VF)")
    print("2. Valor Presente (VP)")
    print("3. Tasa de Interés (i)")
    print("4. Tiempo (n)")
    opcion = int(input("Elija una opción: "))

    if opcion == 1:
        VP = float(input("Capital (VP): "))
        i = float(input("Tasa de interés: "))
        t = leer_tiempo()
        VF = VP * (1 + i) ** t
        print(f"Valor Futuro: {VF:.2f}")
    elif opcion == 2:
        VF = float(input("Valor Futuro (VF): "))
        i = float(input("Tasa de interés: "))
        t = leer_tiempo()
        VP = VF / (1 + i) ** t
        print(f"Valor Presente: {VP:.2f}")
    elif opcion == 3:
        VF = float(input("Valor Futuro: "))
        VP = float(input("Valor Presente: "))
        t = leer_tiempo()
        i = (VF / VP) ** (1/t) - 1
        print(f"Tasa de interés: {i*100:.2f}%")
    elif opcion == 4:
        VF = float(input("Valor Futuro: "))
        VP = float(input("Valor Presente: "))
        i = float(input("Tasa de interés: "))
        n = math.log(VF / VP) / math.log(1 + i)
        print(f"Tiempo: {n:.2f} años")

# ----------------- MENÚ PRINCIPAL -----------------
def main():
    while True:
        print("\n===== CALCULADORA FINANCIERA AVANZADA =====")
        print("1. Interés Simple (básico)")
        print("2. Interés Simple con variaciones de tasa")
        print("3. Interés Compuesto")
        print("4. Salir")
        opcion = int(input("Seleccione una opción: "))

        if opcion == 1:
            interes_simple_general()
        elif opcion == 2:
            interes_simple_variado()
        elif opcion == 3:
            interes_compuesto_general()
        elif opcion == 4:
            print("Saliendo... ¡Hasta luego!")
            break
        else:
            print("Opción inválida, intenta de nuevo.")

if __name__ == "__main__":
    main()


