import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import random


# -------------------------------
# Clase principal de la ventana
# -------------------------------
class MainWindow(tk.Tk):
    """Ventana principal de la aplicación"""
    def __init__(self):
        super().__init__()
        self.title("Catálogo Relacional")
        self.geometry("1200x700")

        # Botón salir
        btn_salir = tk.Button(self, text="Salir aplicación", command=self.quit,
                              bg="red", fg="white")
        btn_salir.pack(anchor="ne", padx=10, pady=5)


# -------------------------------
# Clase para los combos de selección
# -------------------------------
class ComboPanel(tk.Frame):
    """Panel con los combobox de selección de grupos"""
    def __init__(self, parent, grupos, grupos1):
        super().__init__(parent)
        self.pack(pady=10)

        self.grupo_var = tk.StringVar()
        self.combo = ttk.Combobox(self, textvariable=self.grupo_var, state="readonly")
        self.combo["values"] = [f"{k} - {v}" for k, v in grupos.items()]
        self.combo.pack(pady=5)

        self.grupo_var1 = tk.StringVar()
        self.combo1 = ttk.Combobox(self, textvariable=self.grupo_var1, state="readonly")
        self.combo1["values"] = [f"{k} - {v}" for k, v in grupos1.items()]
        self.combo1.pack(pady=5)


# -------------------------------
# Clase para el cuadro de búsqueda
# -------------------------------
class SearchBox(tk.Entry):
    """Caja de texto para búsqueda"""
    def __init__(self, parent):
        super().__init__(parent, width=40)
        self.pack(pady=5)


# -------------------------------
# Clase genérica para Treeview
# -------------------------------
class TreeViewFrame(tk.Frame):
    """Frame que contiene un Treeview con scrollbars"""
    def __init__(self, parent, columnas, ancho_desc=150):
        super().__init__(parent)
        self.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(self, columns=columnas, show="headings")
        for col in columnas:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=ancho_desc if col=="Descripcion" else 150, anchor="center")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)


# -------------------------------
# Clase para el panel de botones
# -------------------------------
class ControlPanel(tk.Frame):
    """Panel con botones de acción"""
    def __init__(self, parent):
        super().__init__(parent)
        self.pack(pady=10)

        tk.Button(self, text="Borrar ítem").pack(side="left", padx=5)
        tk.Button(self, text="Borrar todo Treeview").pack(side="left", padx=5)
        tk.Button(self, text="Exportar a Excel").pack(side="left", padx=5)


# -------------------------------
# Clase para mostrar el total
# -------------------------------
class TotalLabel(tk.Label):
    """Etiqueta para mostrar el importe total"""
    def __init__(self, parent):
        super().__init__(parent, text="Importe Total: 0.00", font=("Arial", 12, "bold"))
        self.pack(pady=5)


#import tkinter as tk

# -------------------------------
# Clase controladora de la GUI
# -------------------------------
class CatalogoController:
    def __init__(self, root, grupos, grupos1, df):
        # Instanciar componentes gráficos
        self.combo_panel = ComboPanel(root, grupos, grupos1)
        self.search_box = SearchBox(root)
        self.tree1_frame = TreeViewFrame(root, ["Codigo","Norden","Descripcion","Precio","PrecioCUP"], ancho_desc=800)
        self.tree2_frame = TreeViewFrame(root, ["Codigo","Norden","Descripcion","Cantidad","Precio","Importe","TipoPago"], ancho_desc=150)
        self.total_label = TotalLabel(root)
        self.control_panel = ControlPanel(root)

        # Guardar referencia al DataFrame
        self.df = df

        # Conectar eventos
        self.combo_panel.combo.bind("<<ComboboxSelected>>", self.mostrar_datos)
        self.search_box.bind("<KeyRelease>", self.mostrar_datos)
        self.tree1_frame.tree.bind("<Double-1>", self.agregar_item)
        self.tree2_frame.tree.bind("<Double-1>", self.editar_celda)

        # Conectar botones
        for btn in self.control_panel.winfo_children():
            if "Borrar ítem" in btn.cget("text"):
                btn.config(command=self.borrar_item)
            elif "Borrar todo" in btn.cget("text"):
                btn.config(command=self.borrar_todo)
            elif "Exportar" in btn.cget("text"):
                btn.config(command=self.exportar_excel)

        # Mostrar datos iniciales
        self.mostrar_datos()

    # -------------------------------
    # Lógica de eventos
    # -------------------------------
    def mostrar_datos(self, event=None):
        """Filtra y muestra datos en el Treeview de catálogo"""
        self.tree1_frame.tree.delete(*self.tree1_frame.tree.get_children())
        if self.combo_panel.grupo_var.get():
            seleccionado = int(self.combo_panel.grupo_var.get().split(" - ")[0])
            filtrado = self.df[self.df["GrupoID"] == seleccionado]
        else:
            filtrado = self.df.copy()
        criterio = self.search_box.get().lower()
        if criterio:
            filtrado = filtrado[
                filtrado["DESCRIPCION"].str.lower().str.contains(criterio, na=False) |
                filtrado["CODIGO"].astype(str).str.lower().str.contains(criterio, na=False)
            ]
        for _, row in filtrado.iterrows():
            self.tree1_frame.tree.insert("", "end", values=(row["CODIGO"], row["Norden"], row["DESCRIPCION"],
                                                            row["PRECIO_T"], row["PRECIOF_CUP"]))

    def agregar_item(self, event):
        """Agrega un ítem del catálogo al Treeview de seleccionados"""
        sel = self.tree1_frame.tree.selection()
        if not sel: return
        codigo, no, desc, precio, precio_cu = self.tree1_frame.tree.item(sel[0], "values")
        pago_input = simpledialog.askstring("Forma de pago", "Ingrese T para Transferencia o E para Efectivo:")
        if not pago_input: return
        pago_input = pago_input.strip().upper()
        tipo_pago = "Transferencia" if pago_input=="T" else "Efectivo" if pago_input=="E" else None
        if not tipo_pago:
            messagebox.showwarning("Atención", "Debe ingresar T o E.")
            return
        self.tree2_frame.tree.insert("", "end", values=(codigo, no, desc, 1, precio, precio, tipo_pago))
        self.actualizar_total()

    def editar_celda(self, event):
        """Permite editar cantidad o descripción en el Treeview de seleccionados"""
        sel = self.tree2_frame.tree.selection()
        if not sel: return
        col = self.tree2_frame.tree.identify_column(event.x)
        col_index = int(col.replace("#","")) - 1
        if col_index not in [2,3]: return
        item = sel[0]
        old_vals = list(self.tree2_frame.tree.item(item,"values"))
        new_val = simpledialog.askfloat("Editar", f"Ingrese nuevo valor para {self.tree2_frame.tree.heading(col)['text']}:")
        if new_val is None: return
        old_vals[col_index] = new_val
        cantidad = float(old_vals[3])
        precio = float(old_vals[4])
        old_vals[5] = round(cantidad*precio,2)
        self.tree2_frame.tree.item(item, values=old_vals)
        self.actualizar_total()

    def borrar_item(self):
        """Elimina ítems seleccionados del Treeview de seleccionados"""
        for s in self.tree2_frame.tree.selection():
            self.tree2_frame.tree.delete(s)
        self.actualizar_total()

    def borrar_todo(self):
        """Elimina todos los ítems del Treeview de seleccionados"""
        self.tree2_frame.tree.delete(*self.tree2_frame.tree.get_children())
        self.actualizar_total()

    def actualizar_total(self):
        """Recalcula el importe total"""
        total = sum(float(self.tree2_frame.tree.item(item,"values")[5]) for item in self.tree2_frame.tree.get_children())
        self.total_label.config(text=f"Importe Total: {round(total,2)}")

    def exportar_excel(self):
        """Exporta los datos seleccionados a Excel"""
        if not self.tree2_frame.tree.get_children():
            messagebox.showwarning("Atención", "No hay datos para exportar.")
            return
        # Aquí luego conectaremos con la lógica de exportación
        messagebox.showinfo("Exportar", "Función de exportación pendiente de implementación.")
        




# -------------------------------
# Clase para manejar los datos
# -------------------------------
class CatalogoData:
    def __init__(self, ruta_excel):
        self.ruta_excel = ruta_excel
        self.df = pd.read_excel(ruta_excel)

        # Definir grupos
        self.grupos = {
            1: "Protección personal",
            2: "Químicos y adhesivos",
            3: "Fijaciones",
            4: "Anclajes y soportes",
            5: "Herrajes",
            6: "Herramientas",
            7: "Accesorios varios"
        }
        self.grupos1 = {
            1: "Ventas",
            2: "Cuenta Casa",
            3: "Traspasos",
        }

        # Añadir columnas de grupo
        self.df["GrupoID"] = [random.choice(list(self.grupos.keys())) for _ in range(len(self.df))]
        self.df["GrupoNombre"] = self.df["GrupoID"].map(self.grupos)

    def filtrar_por_grupo(self, grupo_id, criterio=""):
        filtrado = self.df[self.df["GrupoID"] == grupo_id] if grupo_id else self.df.copy()
        if criterio:
            criterio = criterio.lower()
            filtrado = filtrado[
                filtrado["DESCRIPCION"].str.lower().str.contains(criterio, na=False) |
                filtrado["CODIGO"].astype(str).str.lower().str.contains(criterio, na=False)
            ]
        return filtrado


# -------------------------------
# Controlador que conecta GUI + Datos
# -------------------------------
class CatalogoController:
    def __init__(self, root, data: CatalogoData):
        self.data = data

        # Combobox panel
        self.combo_panel = ComboPanel(root, data.grupos, data.grupos1)
        self.search_box = SearchBox(root)

        # Treeviews
        self.tree1_frame = TreeViewFrame(root, ["Codigo","Norden","Descripcion","Precio","PrecioCUP"], ancho_desc=800)
        self.tree2_frame = TreeViewFrame(root, ["Codigo","Norden","Descripcion","Cantidad","Precio","Importe","TipoPago"], ancho_desc=150)

        # Etiqueta total
        self.total_label = TotalLabel(root)

        # Panel de botones
        self.control_panel = ControlPanel(root)

        # Conectar eventos
        self.combo_panel.combo.bind("<<ComboboxSelected>>", self.mostrar_datos)
        self.search_box.bind("<KeyRelease>", self.mostrar_datos)
        self.tree1_frame.tree.bind("<Double-1>", self.agregar_item)
        self.tree2_frame.tree.bind("<Double-1>", self.editar_celda)

        # Botones
        for btn in self.control_panel.winfo_children():
            if "Borrar ítem" in btn.cget("text"):
                btn.config(command=self.borrar_item)
            elif "Borrar todo" in btn.cget("text"):
                btn.config(command=self.borrar_todo)
            elif "Exportar" in btn.cget("text"):
                btn.config(command=self.exportar_excel)

        # Mostrar datos iniciales
        self.mostrar_datos()

    # -------------------------------
    # Funciones de lógica conectadas a Excel
    # -------------------------------
    def mostrar_datos(self, event=None):
        self.tree1_frame.tree.delete(*self.tree1_frame.tree.get_children())
        grupo_id = int(self.combo_panel.grupo_var.get().split(" - ")[0]) if self.combo_panel.grupo_var.get() else None
        criterio = self.search_box.get()
        filtrado = self.data.filtrar_por_grupo(grupo_id, criterio)
        for _, row in filtrado.iterrows():
            self.tree1_frame.tree.insert("", "end", values=(row["CODIGO"], row["Norden"], row["DESCRIPCION"],
                                                            row["PRECIO_T"], row["PRECIOF_CUP"]))

    def agregar_item(self, event):
        sel = self.tree1_frame.tree.selection()
        if not sel: return
        codigo, no, desc, precio, precio_cu = self.tree1_frame.tree.item(sel[0], "values")
        pago_input = simpledialog.askstring("Forma de pago", "Ingrese T para Transferencia o E para Efectivo:")
        if not pago_input: return
        pago_input = pago_input.strip().upper()
        tipo_pago = "Transferencia" if pago_input=="T" else "Efectivo" if pago_input=="E" else None
        if not tipo_pago:
            messagebox.showwarning("Atención", "Debe ingresar T o E.")
            return
        self.tree2_frame.tree.insert("", "end", values=(codigo, no, desc, 1, precio, precio, tipo_pago))
        self.actualizar_total()

    def editar_celda(self, event):
        sel = self.tree2_frame.tree.selection()
        if not sel: return
        col = self.tree2_frame.tree.identify_column(event.x)
        col_index = int(col.replace("#","")) - 1
        if col_index not in [2,3]: return
        item = sel[0]
        old_vals = list(self.tree2_frame.tree.item(item,"values"))
        new_val = simpledialog.askfloat("Editar", f"Ingrese nuevo valor para {self.tree2_frame.tree.heading(col)['text']}:")
        if new_val is None: return
        old_vals[col_index] = new_val
        cantidad = float(old_vals[3])
        precio = float(old_vals[4])
        old_vals[5] = round(cantidad*precio,2)
        self.tree2_frame.tree.item(item, values=old_vals)
        self.actualizar_total()

    def borrar_item(self):
        for s in self.tree2_frame.tree.selection():
            self.tree2_frame.tree.delete(s)
        self.actualizar_total()

    def borrar_todo(self):
        self.tree2_frame.tree.delete(*self.tree2_frame.tree.get_children())
        self.actualizar_total()

    def actualizar_total(self):
        total = sum(float(self.tree2_frame.tree.item(item,"values")[5]) for item in self.tree2_frame.tree.get_children())
        self.total_label.config(text=f"Importe Total: {round(total,2)}")

    def exportar_excel(self):
        if not self.tree2_frame.tree.get_children():
            messagebox.showwarning("Atención", "No hay datos para exportar.")
            return
        messagebox.showinfo("Exportar", "Aquí se implementará la exportación a Excel.")
        


# -------------------------------
# Ejemplo de instanciación de la GUI
# -------------------------------
if __name__ == "__main__":
 
    # Crear ventana principal
    root = MainWindow()

    # Ruta del archivo Excel
    ruta_excel = "/storage/emulated/0/DetalleZuluetaF23526.xlsx"

    # Crear objeto de datos (carga el Excel y prepara los grupos)
    data = CatalogoData(ruta_excel)

    # Conectar GUI con los datos
    app = CatalogoController(root, data)
    root.mainloop()