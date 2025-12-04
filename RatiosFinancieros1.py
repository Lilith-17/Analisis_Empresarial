import tkinter as tk
from tkinter import ttk, scrolledtext
import math

# ---------- UTILIDADES ----------
def safe_float(s):
    try: return float(s)
    except: return 0.0

def safe_div(a, b): return a / b if b else float('nan')
def pct(v): return v*100 if v==v else float('nan')
def fmt_num(v): return "N/A" if v!=v else f"{v:,.2f}"
def fmt_pct(v): return "N/A" if v!=v else f"{pct(v):.2f}%"

# ---------- INTERPRETACIONES ----------
def interpretar(ratio, v):
    if ratio=="Razón corriente":
        if v<1: return "Insuficiente: riesgo de iliquidez."
        if v<=2: return "Aceptable: cubre deudas de corto plazo."
        return "Exceso de liquidez."
    if ratio=="Prueba ácida":
        return "Buena liquidez sin inventario." if v>=1 else "Riesgo: depende de inventarios."
    if ratio=="Capital de trabajo":
        return "Positivo: operación estable." if v>=0 else "Negativo: riesgo de insolvencia."
    if ratio=="Rotación inventarios":
        return "Adecuada." if 3<=v<=6 else ("Baja rotación." if v<3 else "Muy alta.")
    if ratio=="Período promedio de inventarios":
        return "Razonable." if v<=120 else "Inventarios lentos."
    if ratio=="Rotación CxC":
        return "Buena cobranza." if v>=4 else "Cobranza lenta."
    if ratio=="Período promedio de cobro":
        return "Saludable." if v<=90 else "Cobro tardío."
    if ratio=="Rotación activos":
        return "Activos eficientes." if v>=1 else "Uso ineficiente de activos."
    if ratio=="Pasivo/Activo":
        if v>0.6: return "Alto endeudamiento."
        if v>=0.4: return "Moderado."
        return "Bajo endeudamiento."
    if ratio=="Deuda/Patrimonio":
        return "Riesgo alto." if v>1 else "Estructura equilibrada."
    if ratio=="Cobertura intereses":
        return "Adecuada." if v>=2 else "Débil cobertura."
    if ratio=="Margen neto":
        return "Excelente." if v>0.15 else ("Adecuado." if v>=0.05 else "Bajo.")
    if ratio=="ROA":
        return "Aceptable." if v>=0.05 else "Bajo."
    if ratio=="ROE":
        if v>0.2: return "Excelente retorno."
        if v>=0.1: return "Adecuado."
        return "Bajo."
    return "Sin interpretación."

# ---------- CAMPOS Y RATIOS ----------
CATEGORIES={
"Liquidez":["Razón corriente","Prueba ácida","Capital de trabajo"],
"Actividad":["Rotación inventarios","Período promedio de inventarios","Rotación CxC","Período promedio de cobro","Rotación activos"],
"Endeudamiento":["Pasivo/Activo","Deuda/Patrimonio","Cobertura intereses"],
"Rentabilidad":["Margen neto","ROA","ROE"]
}
FIELDS={
"Razón corriente":["Activo corriente","Pasivo corriente"],
"Prueba ácida":["Activo corriente","Inventarios","Pasivo corriente"],
"Capital de trabajo":["Activo corriente","Pasivo corriente"],
"Rotación inventarios":["Costo de ventas","Inventario promedio"],
"Período promedio de inventarios":["Costo de ventas","Inventario promedio"],
"Rotación CxC":["Ventas netas","Cuentas por cobrar"],
"Período promedio de cobro":["Ventas netas","Cuentas por cobrar"],
"Rotación activos":["Ventas netas","Activo total"],
"Pasivo/Activo":["Pasivo total","Activo total"],
"Deuda/Patrimonio":["Pasivo total","Patrimonio"],
"Cobertura intereses":["UAII","Gastos por intereses"],
"Margen neto":["Utilidad neta","Ventas netas"],
"ROA":["Utilidad neta","Activo total"],
"ROE":["Utilidad neta","Patrimonio"]
}

# ---------- CÁLCULO ----------
def calcular(r,vals):
    if r=="Razón corriente": return safe_div(vals["Activo corriente"],vals["Pasivo corriente"])
    if r=="Prueba ácida": return safe_div(vals["Activo corriente"]-vals["Inventarios"],vals["Pasivo corriente"])
    if r=="Capital de trabajo": return vals["Activo corriente"]-vals["Pasivo corriente"]
    if r=="Rotación inventarios": return safe_div(vals["Costo de ventas"],vals["Inventario promedio"])
    if r=="Período promedio de inventarios": return safe_div(360,safe_div(vals["Costo de ventas"],vals["Inventario promedio"]))
    if r=="Rotación CxC": return safe_div(vals["Ventas netas"],vals["Cuentas por cobrar"])
    if r=="Período promedio de cobro": return safe_div(360,safe_div(vals["Ventas netas"],vals["Cuentas por cobrar"]))
    if r=="Rotación activos": return safe_div(vals["Ventas netas"],vals["Activo total"])
    if r=="Pasivo/Activo": return safe_div(vals["Pasivo total"],vals["Activo total"])
    if r=="Deuda/Patrimonio": return safe_div(vals["Pasivo total"],vals["Patrimonio"])
    if r=="Cobertura intereses": return safe_div(vals["UAII"],vals["Gastos por intereses"])
    if r=="Margen neto": return safe_div(vals["Utilidad neta"],vals["Ventas netas"])
    if r=="ROA": return safe_div(vals["Utilidad neta"],vals["Activo total"])
    if r=="ROE": return safe_div(vals["Utilidad neta"],vals["Patrimonio"])
    return float('nan')

# ---------- SCROLLABLE ----------
class Scroll(ttk.Frame):
    def __init__(self,root):
        super().__init__(root)
        c=tk.Canvas(self)
        sb=ttk.Scrollbar(self,orient="vertical",command=c.yview)
        self.frm=ttk.Frame(c)
        self.frm.bind("<Configure>",lambda e:c.configure(scrollregion=c.bbox("all")))
        c.create_window((0,0),window=self.frm,anchor="nw")
        c.configure(yscrollcommand=sb.set)
        c.pack(side="left",fill="both",expand=True)
        sb.pack(side="right",fill="y")
        self.frm.bind("<Enter>",lambda e:c.bind_all("<MouseWheel>",lambda ev:c.yview_scroll(int(-1*(ev.delta/120)),"units")))
        self.frm.bind("<Leave>",lambda e:c.unbind_all("<MouseWheel>"))

# ---------- APP ----------
class App:
    def __init__(self,root):
        root.title("Calculadora de Ratios Financieros")
        root.geometry("950x720")
        main=Scroll(root)
        main.pack(fill="both",expand=True)
        self.frm=main.frm

        ttk.Label(self.frm,text="Calculadora de Ratios Financieros",font=("Segoe UI",14,"bold")).grid(row=0,column=0,columnspan=4,pady=8)
        ttk.Label(self.frm,text="Categoría:").grid(row=1,column=0,sticky="w",padx=5)
        self.cat=ttk.Combobox(self.frm,values=list(CATEGORIES.keys()),state="readonly",width=25)
        self.cat.grid(row=1,column=1,sticky="w")
        self.cat.bind("<<ComboboxSelected>>",self.sel_cat)
        ttk.Label(self.frm,text="Ratio:").grid(row=2,column=0,sticky="w",padx=5)
        self.ratio=ttk.Combobox(self.frm,state="readonly",width=40)
        self.ratio.grid(row=2,column=1,sticky="w")
        self.ratio.bind("<<ComboboxSelected>>",self.sel_ratio)

        # frame entradas
        self.entries={}
        ttk.Label(self.frm,text="Datos requeridos:",font=("Segoe UI",10,"bold")).grid(row=3,column=0,columnspan=4,pady=6,sticky="w")
        all=set(sum(FIELDS.values(),[]))
        row=4
        for f in sorted(all):
            ttk.Label(self.frm,text=f+":").grid(row=row,column=0,sticky="w",padx=6)
            e=ttk.Entry(self.frm,width=20,state="disabled")
            e.grid(row=row,column=1,sticky="w",padx=6)
            self.entries[f]=e
            row+=1
        self.row_calc=row+1
        ttk.Button(self.frm,text="Calcular",command=self.calc).grid(row=self.row_calc,column=0,padx=8,pady=10,sticky="w")
        ttk.Button(self.frm,text="Limpiar",command=self.reset).grid(row=self.row_calc,column=1,padx=8,pady=10,sticky="w")

        ttk.Label(self.frm,text="Reporte:",font=("Segoe UI",10,"bold")).grid(row=self.row_calc+1,column=0,columnspan=3,sticky="w")
        self.rep=scrolledtext.ScrolledText(self.frm,width=110,height=18)
        self.rep.grid(row=self.row_calc+2,column=0,columnspan=4,padx=8,pady=4)
        self.rep.configure(state="disabled")

    def sel_cat(self,e=None):
        cat=self.cat.get()
        self.ratio["values"]=CATEGORIES.get(cat,[])
        self.ratio.set("")
        self.disable_all()

    def sel_ratio(self,e=None):
        r=self.ratio.get()
        self.disable_all()
        for f in FIELDS.get(r,[]):
            self.entries[f].configure(state="normal")

    def disable_all(self):
        for e in self.entries.values():
            e.configure(state="disabled")
            e.delete(0,"end")
        self.clear()

    def calc(self):
        r=self.ratio.get()
        if not r:
            self.show("Seleccione categoría y ratio.")
            return
        vals={f:safe_float(self.entries[f].get()) for f in FIELDS[r]}
        val=calcular(r,vals)
        inter=interpretar(r,val)
        lines=[f"RATIO: {r}","-"*40]
        lines.append(f"Resultado: {fmt_pct(val) if r in ['Pasivo/Activo','Margen neto','ROA','ROE'] else fmt_num(val)}")
        lines.append(f"Interpretación: {inter}")
        self.show("\n".join(lines))

    def show(self,txt):
        self.rep.configure(state="normal")
        self.rep.delete("1.0","end")
        self.rep.insert("end",txt)
        self.rep.configure(state="disabled")

    def clear(self):
        self.rep.configure(state="normal")
        self.rep.delete("1.0","end")
        self.rep.configure(state="disabled")

    def reset(self):
        self.cat.set("")
        self.ratio.set("")
        self.disable_all()

if __name__=="__main__":
    root=tk.Tk()
    App(root)
    root.mainloop()

