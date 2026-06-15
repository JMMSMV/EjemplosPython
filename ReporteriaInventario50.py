import pandas as pd
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox

# --- Funciones de reporte ---
def reporte_compacto(df, tc, guardar=True):
    # Solicitar rango de fechas
    fecha_ini = simpledialog.askstring("Fecha inicial", "Ingrese fecha inicial (dd/mm/yyyy):")
    fecha_fin = simpledialog.askstring("Fecha final", "Ingrese fecha final (dd/mm/yyyy):")
    if not fecha_ini or not fecha_fin:
        messagebox.showwarning("Atención", "Debe ingresar ambas fechas.")
        return None, None, None

    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True)
    mask = (df["Fecha"] >= pd.to_datetime(fecha_ini, dayfirst=True)) & (df["Fecha"] <= pd.to_datetime(fecha_fin, dayfirst=True))
    df = df.loc[mask]

    df["Importe"] = df["Cantidad"] * df["Precio"]

    # Desagregar importes por tipo de pago
    df["Importe_Efectivo"] = df.apply(lambda x: x["Importe"] if x["TipoPago"]=="Efectivo" else 0, axis=1)
    df["Importe_Transferencia"] = df.apply(lambda x: x["Importe"] if x["TipoPago"]=="Transferencia" else 0, axis=1)

    # Agrupar
    tabla = df.groupby(["Codigo","Descripcion"]).agg(
        Cantidad=("Cantidad","sum"),
        Importe_Efectivo=("Importe_Efectivo","sum"),
        Importe_Transferencia=("Importe_Transferencia","sum"),
        Importe_Total=("Importe","sum")
    ).reset_index()

    # Totales
    total_efectivo = tabla["Importe_Efectivo"].sum()
    total_transfer = tabla["Importe_Transferencia"].sum()
    total_usd = tabla["Importe_Total"].sum()
    total_cup = total_usd * tc
    efectivo_cup = total_efectivo * tc
    transfer_cup = total_transfer * tc

    # Construir DataFrame igual al Treeview (incluyendo resumen)
    filas = tabla.copy()
    filas = filas.astype({"Cantidad":"int"})  # cantidad entera
    filas["Importe_Efectivo"] = filas["Importe_Efectivo"].map(lambda x: f"{x:.2f}")
    filas["Importe_Transferencia"] = filas["Importe_Transferencia"].map(lambda x: f"{x:.2f}")
    filas["Importe_Total"] = filas["Importe_Total"].map(lambda x: f"{x:.2f}")

    # Añadir resumen como filas extra
    resumen_filas = pd.DataFrame([
        {"Codigo":"", "Descripcion":"TOTAL USD", "Cantidad":"", 
         "Importe_Efectivo":f"{total_efectivo:.2f}", 
         "Importe_Transferencia":f"{total_transfer:.2f}", 
         "Importe_Total":f"{total_usd:.2f}"},
        {"Codigo":"", "Descripcion":"TOTAL CUP", "Cantidad":"", 
         "Importe_Efectivo":f"{efectivo_cup:.2f}", 
         "Importe_Transferencia":f"{transfer_cup:.2f}", 
         "Importe_Total":f"{total_cup:.2f}"},
        {"Codigo":"", "Descripcion":f"Fecha de corte {fecha_ini} - {fecha_fin}", "Cantidad":"", 
         "Importe_Efectivo":"", "Importe_Transferencia":"", 
         "Importe_Total":f"Tasa: {tc:.2f}"}
    ])

    exportar = pd.concat([filas, resumen_filas], ignore_index=True)

    # Guardar
    if guardar:
        ruta2= "/storage/emulated/0/Reporte_compacto.xlsx"
        exportar.to_excel(ruta2, index=False)
    else:
        ruta2 = None

    return tabla, resumen_filas, ruta2

# --- Interfaz gráfica ---
def generar_reporte():
    tc = float(entry_tc.get())
    ruta3= "/storage/emulated/0/Seleccionados.xlsx"
    df = pd.read_excel(ruta3)

    guardar = bool(save_var.get())

    tabla, resumen, archivo = reporte_compacto(df, tc, guardar)

    if tabla is not None:
        lbl_result.config(text=f"Reporte generado: {archivo if archivo else 'Solo visualizado'}")

        # Limpiar Treeview
        for item in tree.get_children():
            tree.delete(item)

        # Insertar filas
        for _, fila in tabla.iterrows():
            tree.insert("", tk.END, values=(fila["Codigo"], fila["Descripcion"], int(fila["Cantidad"]),
                                            f"{fila['Importe_Efectivo']:.2f}", f"{fila['Importe_Transferencia']:.2f}", f"{fila['Importe_Total']:.2f}"))

        # Insertar resumen
        for _, fila in resumen.iterrows():
            tree.insert("", tk.END, values=(fila["Codigo"], fila["Descripcion"], fila["Cantidad"],
                                            fila["Importe_Efectivo"], fila["Importe_Transferencia"], fila["Importe_Total"]))

# Ventana principal
root = tk.Tk()
root.title("Generador de Reportes")

btn_salir = tk.Button(root, text="Salir aplicación", command=root.quit, bg="red", fg="white")
btn_salir.pack(anchor="ne", padx=10, pady=5)

tk.Label(root, text="Tasa de cambio (tc):").pack(pady=5)
entry_tc = tk.Entry(root)
entry_tc.insert(0,"545.00")
entry_tc.pack(pady=5)

# Checkbox guardar
save_var = tk.IntVar(value=1)
chk_save = tk.Checkbutton(root, text="Guardar en Excel", variable=save_var)
chk_save.pack(pady=5)

btn = tk.Button(root, text="Generar", command=generar_reporte)
btn.pack(pady=10)

lbl_result = tk.Label(root, text="")
lbl_result.pack(pady=5)

# --- Treeview con scrollbars ---
frame_tabla = ttk.Frame(root)
frame_tabla.pack(expand=True, fill="both", pady=10)

columnas = ("Codigo","Producto","Cantidad","Importe_Efectivo","Importe_Transferencia","Importe_Total")
tree = ttk.Treeview(frame_tabla, columns=columnas, show="headings")

for col in columnas:
    tree.heading(col, text=col)
    tree.column(col, anchor="center", width=150)

scroll_y = ttk.Scrollbar(frame_tabla, orient="vertical", command=tree.yview)
scroll_x = ttk.Scrollbar(frame_tabla, orient="horizontal", command=tree.xview)
tree.configure(yscroll=scroll_y.set, xscroll=scroll_x.set)

tree.grid(row=0, column=0, sticky="nsew")
scroll_y.grid(row=0, column=1, sticky="ns")
scroll_x.grid(row=1, column=0, sticky="ew")

frame_tabla.rowconfigure(0, weight=1)
frame_tabla.columnconfigure(0, weight=1)

root.mainloop()