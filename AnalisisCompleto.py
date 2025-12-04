# Analisis Integral (A, B, C, D) - GUI Tkinter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import math
import matplotlib
matplotlib.use('TkAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -------------------------
# Funciones de cálculo
# -------------------------
def calc_all(df_row):
    """Recibe diccionario (fila) con los campos y devuelve un dict con resultados."""
    r = {}
    # Totales activos
    TA_2023 = df_row.get('Activo Corriente 2023', 0) + df_row.get('Activo No Corriente 2023', 0)
    TA_2024 = df_row.get('Activo Corriente 2024', 0) + df_row.get('Activo No Corriente 2024', 0)
    r['TA_2023'] = TA_2023
    r['TA_2024'] = TA_2024

    # A1 - Fondo de Maniobra
    r['FM_2023'] = df_row.get('Activo Corriente 2023', 0) - df_row.get('Pasivo Corriente 2023', 0)
    r['FM_2024'] = df_row.get('Activo Corriente 2024', 0) - df_row.get('Pasivo Corriente 2024', 0)

    # A2 - Vertical 2024 (porcentajes)
    if TA_2024 != 0:
        r['vertical_AC_2024']  = df_row.get('Activo Corriente 2024',0) / TA_2024 * 100
        r['vertical_ANC_2024'] = df_row.get('Activo No Corriente 2024',0) / TA_2024 * 100
        r['vertical_PC_2024']  = df_row.get('Pasivo Corriente 2024',0) / TA_2024 * 100
        r['vertical_PN_2024']  = df_row.get('Patrimonio Neto 2024',0) / TA_2024 * 100
    else:
        r['vertical_AC_2024'] = r['vertical_ANC_2024'] = r['vertical_PC_2024'] = r['vertical_PN_2024'] = 0.0

    # A3 - Horizontal (crecimientos %)
    def pct(new, old):
        try:
            return (new - old) / old * 100
        except:
            return float('nan')
    r['h_AC']  = pct(df_row.get('Activo Corriente 2024',0), df_row.get('Activo Corriente 2023',0))
    r['h_ANC'] = pct(df_row.get('Activo No Corriente 2024',0), df_row.get('Activo No Corriente 2023',0))
    r['h_PC']  = pct(df_row.get('Pasivo Corriente 2024',0), df_row.get('Pasivo Corriente 2023',0))
    r['h_PN']  = pct(df_row.get('Patrimonio Neto 2024',0), df_row.get('Patrimonio Neto 2023',0))

    # A4 - Ciclo de Conversión de Efectivo
    dias_inv = int(df_row.get('Dias Inventario', 45))
    dias_cli = int(df_row.get('Dias Clientes', 60))
    dias_prov= int(df_row.get('Dias Proveedores', 30))
    r['CCE_2024'] = dias_inv + dias_cli - dias_prov

    # B1 - Ratios de liquidez
    pc2023 = max(df_row.get('Pasivo Corriente 2023', 1), 1)
    pc2024 = max(df_row.get('Pasivo Corriente 2024', 1), 1)
    r['LG_2023'] = df_row.get('Activo Corriente 2023', 0) / pc2023
    r['LG_2024'] = df_row.get('Activo Corriente 2024', 0) / pc2024
    r['T_2024']  = (df_row.get('Caja 2024', 0) + df_row.get('Clientes 2024', 0)) / pc2024
    r['D_2024']  = df_row.get('Caja 2024', 0) / pc2024

    # B2 - Solvencia
    deuda2024 = max(df_row.get('Deuda Total 2024', 1), 1)
    r['garantia_2024'] = (TA_2024) / deuda2024
    r['autonomia_2024'] = df_row.get('Patrimonio Neto 2024', 0) / deuda2024
    r['calidad_2024'] = df_row.get('Pasivo Corriente 2024', 0) / deuda2024

    # C1/C2 - Rentabilidades
    r['RAT_2023'] = (df_row.get('BAII 2023', 0) / TA_2023 * 100) if TA_2023 else 0.0
    r['RAT_2024'] = (df_row.get('BAII 2024', 0) / TA_2024 * 100) if TA_2024 else 0.0
    r['RRP_2023'] = df_row.get('UN 2023', 0) / max(df_row.get('Patrimonio Neto 2023', 1),1) * 100
    r['RRP_2024'] = df_row.get('UN 2024', 0) / max(df_row.get('Patrimonio Neto 2024', 1),1) * 100

    # C3 - DuPont components (2024)
    ingresos2024 = max(df_row.get('Ingresos 2024', 1), 1)
    r['margen_bruto_2024'] = df_row.get('Ganancia Bruta 2024', df_row.get('BAII 2024', 0) + df_row.get('Costo Servicios 2024', 0)) / ingresos2024 * 100
    r['margen_operativo_2024'] = df_row.get('BAII 2024', 0) / ingresos2024 * 100
    r['margen_neto_2024'] = df_row.get('UN 2024', 0) / ingresos2024 * 100
    r['rotacion_activo_2024'] = df_row.get('Ingresos 2024', 0) / TA_2024 if TA_2024 else 0.0
    r['apalancamiento_2024'] = TA_2024 / max(df_row.get('Patrimonio Neto 2024', 1), 1)

    # C5 - apalancamiento y costo de deuda
    r['costo_deuda_2024'] = df_row.get('Gastos Financieros 2024', 0) / max(df_row.get('Deuda Total 2024',1),1) * 100
    r['efecto_apalancamiento'] = r['RAT_2024'] + (df_row.get('Deuda Total 2024',0) / max(df_row.get('Patrimonio Neto 2024',1),1)) * (r['RAT_2024'] - r['costo_deuda_2024'])

    return r

# -------------------------
# GUI: pestañas A, B, C, D
# -------------------------
def launch_app():
    root = tk.Tk()
    root.title("Análisis Integral (A,B,C,D) - Interactivo")
    root.geometry("1100x750")
    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    shared = {'df': None, 'row': None, 'results': None}

    # Tab: carga
    tab_load = ttk.Frame(notebook)
    notebook.add(tab_load, text="Carga de datos")
    lbl = ttk.Label(tab_load, text="Cargar archivo Excel o CSV (se usará la PRIMERA FILA). Columnas deben tener nombres esperados.")
    lbl.pack(pady=6)
    btn_load = ttk.Button(tab_load, text="Cargar archivo (Excel/CSV)", width=30)
    btn_load.pack(pady=4)
    display = scrolledtext.ScrolledText(tab_load, width=130, height=28)
    display.pack(padx=8, pady=8)

    def do_load():
        fp = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls"), ("CSV files","*.csv")])
        if not fp:
            return
        try:
            if fp.lower().endswith('.csv'):
                df = pd.read_csv(fp)
            else:
                df = pd.read_excel(fp)
            shared['df'] = df
            shared['row'] = df.iloc[0].to_dict()
            display.delete('1.0', tk.END)
            display.insert(tk.END, "Preview (head):\n")
            display.insert(tk.END, df.head().to_string())
            messagebox.showinfo("Carga OK", "Archivo cargado correctamente. Se usará la primer fila para los cálculos.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo: {e}")

    btn_load.config(command=do_load)

    # Tab A
    tabA = ttk.Frame(notebook)
    notebook.add(tabA, text="A - Patrimonial")
    textA = scrolledtext.ScrolledText(tabA, width=130, height=25)
    textA.pack(padx=6, pady=6)
    def runA():
        if not shared['row']:
            messagebox.showwarning("Atencion","Carga un archivo primero.")
            return
        r = calc_all(shared['row'])
        shared['results'] = r
        textA.delete('1.0', tk.END)
        textA.insert(tk.END, f"Fondo de Maniobra 2023: {r['FM_2023']:.2f}\n")
        textA.insert(tk.END, f"Fondo de Maniobra 2024: {r['FM_2024']:.2f}\n\n")
        textA.insert(tk.END, "Análisis vertical 2024 (% del Activo Total):\n")
        textA.insert(tk.END, f"  Activo Corriente: {r['vertical_AC_2024']:.2f}%\n")
        textA.insert(tk.END, f"  Activo No Corriente: {r['vertical_ANC_2024']:.2f}%\n")
        textA.insert(tk.END, f"  Pasivo Corriente: {r['vertical_PC_2024']:.2f}%\n")
        textA.insert(tk.END, f"  Patrimonio Neto: {r['vertical_PN_2024']:.2f}%\n\n")
        textA.insert(tk.END, "Análisis horizontal (variación % 2024 vs 2023):\n")
        textA.insert(tk.END, f"  Activo Corriente: {r['h_AC']:.2f}%\n")
        textA.insert(tk.END, f"  Activo No Corriente: {r['h_ANC']:.2f}%\n")
        textA.insert(tk.END, f"  Pasivo Corriente: {r['h_PC']:.2f}%\n")
        textA.insert(tk.END, f"  Patrimonio Neto: {r['h_PN']:.2f}%\n\n")
        textA.insert(tk.END, f"Ciclo de Conversión de Efectivo (CCE) 2024: {r['CCE_2024']} días\n")

    btnA = ttk.Button(tabA, text="Ejecutar A", command=runA)
    btnA.pack(pady=4)

    # Tab B
    tabB = ttk.Frame(notebook)
    notebook.add(tabB, text="B - Financiero")
    textB = scrolledtext.ScrolledText(tabB, width=130, height=25)
    textB.pack(padx=6, pady=6)
    def runB():
        if not shared['row']:
            messagebox.showwarning("Atencion","Carga un archivo primero.")
            return
        r = calc_all(shared['row'])
        shared['results'] = r
        textB.delete('1.0', tk.END)
        textB.insert(tk.END, f"Liquidez General 2023: {r['LG_2023']:.2f}\n")
        textB.insert(tk.END, f"Liquidez General 2024: {r['LG_2024']:.2f}\n")
        textB.insert(tk.END, f"Razón Tesorería 2024: {r['T_2024']:.2f}\n")
        textB.insert(tk.END, f"Disponibilidad 2024: {r['D_2024']:.2f}\n\n")
        textB.insert(tk.END, f"Ratio Garantía 2024: {r['garantia_2024']:.2f}\n")
        textB.insert(tk.END, f"Ratio Autonomía 2024: {r['autonomia_2024']:.2f}\n")
        textB.insert(tk.END, f"Calidad Deuda 2024: {r['calidad_2024']:.2f}\n\n")
        # escenario pesimista
        ingresos24 = shared['row'].get('Ingresos 2024', shared['row'].get('Ingresos', 0))
        ingresos_25 = ingresos24 * 0.7
        gastos_fijos = ingresos_25 * 0.6
        ac25 = shared['row'].get('Activo Corriente 2024',0) * 0.7
        pc25 = shared['row'].get('Pasivo Corriente 2024',0) + gastos_fijos
        fm25 = ac25 - pc25
        rl25 = ac25 / pc25 if pc25!=0 else float('inf')
        textB.insert(tk.END, "Escenario pesimista 2025 (ingresos -30%):\n")
        textB.insert(tk.END, f"  FM estimado 2025: {fm25:.2f}\n")
        textB.insert(tk.END, f"  Liquidez estimada 2025: {rl25:.2f}\n")
    btnB = ttk.Button(tabB, text="Ejecutar B", command=runB)
    btnB.pack(pady=4)

    # Tab C
    tabC = ttk.Frame(notebook)
    notebook.add(tabC, text="C - Económico")
    textC = scrolledtext.ScrolledText(tabC, width=130, height=25)
    textC.pack(padx=6, pady=6)
    def runC():
        if not shared['row']:
            messagebox.showwarning("Atencion","Carga un archivo primero.")
            return
        r = calc_all(shared['row'])
        shared['results'] = r
        textC.delete('1.0', tk.END)
        textC.insert(tk.END, f"RAT 2023: {r['RAT_2023']:.2f}%\n")
        textC.insert(tk.END, f"RAT 2024: {r['RAT_2024']:.2f}%\n")
        textC.insert(tk.END, f"RRP 2023: {r['RRP_2023']:.2f}%\n")
        textC.insert(tk.END, f"RRP 2024: {r['RRP_2024']:.2f}%\n\n")
        textC.insert(tk.END, "Márgenes 2024:\n")
        textC.insert(tk.END, f"  Margen bruto: {r['margen_bruto_2024']:.2f}%\n")
        textC.insert(tk.END, f"  Margen operativo: {r['margen_operativo_2024']:.2f}%\n")
        textC.insert(tk.END, f"  Margen neto: {r['margen_neto_2024']:.2f}%\n\n")
        textC.insert(tk.END, f"Costo deuda 2024: {r['costo_deuda_2024']:.2f}%\n")
        textC.insert(tk.END, f"Efecto apalancamiento: {r['efecto_apalancamiento']:.2f}%\n")
    btnC = ttk.Button(tabC, text="Ejecutar C", command=runC)
    btnC.pack(pady=4)

    # Tab D
    tabD = ttk.Frame(notebook)
    notebook.add(tabD, text="D - Diagnóstico")
    textD = scrolledtext.ScrolledText(tabD, width=90, height=20)
    textD.pack(side='left', padx=6, pady=6)
    fig_frame = ttk.Frame(tabD)
    fig_frame.pack(side='right', padx=6, pady=6)
    def runD():
        if not shared['results']:
            messagebox.showwarning("Atencion","Ejecuta A/B/C primero o carga datos y luego ejecuta alguna sección.")
            return
        r = shared['results']
        # DataFrame resumen
        df_ratios = pd.DataFrame({
            'Ratio': ['FM','Liquidez Gen','RAT','RRP'],
            '2023': [r['FM_2023'], r['LG_2023'], r['RAT_2023'], r['RRP_2023']],
            '2024': [r['FM_2024'], r['LG_2024'], r['RAT_2024'], r['RRP_2024']]
        })
        df_ratios['Cambio'] = df_ratios['2024'] - df_ratios['2023']
        textD.delete('1.0', tk.END)
        textD.insert(tk.END, df_ratios.to_string(index=False))
        # plot FM
        for widget in fig_frame.winfo_children():
            widget.destroy()
        fig = Figure(figsize=(5,3))
        ax = fig.add_subplot(111)
        ax.bar(['FM 2023','FM 2024'], [r['FM_2023'], r['FM_2024']], color='skyblue')
        ax.set_title('Fondo de Maniobra')
        canvas = FigureCanvasTkAgg(fig, master=fig_frame)
        canvas.draw()
        canvas.get_tk_widget().pack()
    btnD = ttk.Button(tabD, text="Ejecutar D", command=runD)
    btnD.pack(pady=4)

    root.mainloop()

# Lanzar la aplicación
if __name__ == "__main__":
    launch_app()

