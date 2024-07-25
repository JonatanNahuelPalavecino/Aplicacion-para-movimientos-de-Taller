from customtkinter import *
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from datetime import datetime
import customtkinter
import sqlite3
from PIL import Image, ImageTk
import openpyxl
from dotenv import load_dotenv
import os

color_uno = "#000000"
color_dos = "#010101"
color_tres = "#00ffff"
color_cuatro = "#ff7b00"
color_cinco = "#ff00aa"
color_seis = "#ffffff"
gris_oscuro = "#272626"
gris_claro = "#4b4a4a"

nuevo_movimiento = None

class Movimiento:

    def __init__(self, fecha, solicitud, bo, mov_ax):
        self.fecha = fecha
        self.solicitud = solicitud
        self.bo = bo
        self.mov_ax = mov_ax
        self.items = []

    def agregar_items(self, item):
        self.items.append(item)

    def modificar_items(self, item_anterior, item_nuevo):
        self.items.remove(item_anterior)
        self.items.append(item_nuevo)

    def eliminar_item(self, item):
        self.items.remove(item)

    def total_items(self):
        return len(self.items)

class App:

    db_name = "database.db"

    def __init__(self, window) -> None:

        self.root = window
        self.root.columnconfigure(0, weight = 1)
        self.root.rowconfigure(0, weight = 1)
        self.root.title("App de movimientos")
        self.center_window(self.root, 500, 600)
        self.root.minsize(480,500)
        self.root.config(bg = color_uno)
        self.root.after(200, self.crear_icono)
        self.crear_icono()

        # Creacion del Frame Container

        frame = CTkFrame(self.root, bg_color = color_dos)
        frame.grid(column = 0, row = 0, sticky = "nsew", padx = 50, pady = 50)

        frame.columnconfigure([0, 1], weight = 1)
        frame.rowconfigure([0, 1, 2, 3], weight = 1)

        # Inclusión del logo en el Frame

        logo_img = Image.open("images/logo.png")
        logo = CTkImage(light_image=logo_img, size = (148, 65))
        CTkLabel(frame, image = logo, text = "").grid(column= 0, row = 0, columnspan = 2)

        # Botonera de Inicio

        CTkButton(frame, font = ("sans serif", 17), text = "Registrar un Ingreso", border_color = color_tres, fg_color = color_cuatro, hover_color = color_cinco, corner_radius = 25, height = 40, command = lambda: self.open_window("Ingreso")).grid(column = 0, row = 1, columnspan = 2)

        CTkButton(frame, font = ("sans serif", 17), text = "Registrar un Egreso", border_color = color_tres, fg_color = color_cuatro, hover_color = color_cinco, corner_radius = 25, height = 40, command = lambda: self.open_window("Egreso")).grid(column = 0, row = 2, columnspan = 2)

        CTkButton(frame, font = ("sans serif", 17), text = "Ver Movimientos", border_color = color_tres, fg_color = color_cuatro, hover_color = color_cinco, corner_radius = 25, height = 40, command = self.open_window).grid(column = 0, row = 3, columnspan = 2)

    def crear_icono(self):
        icono_img = Image.open("images/icono.png")
        icono = CTkImage(light_image=icono_img)
        self.root.call("wm", "iconphoto", self.root._w, ImageTk.PhotoImage(icono_img))

    def center_window(self, win, width, height):
        screen_width = win.winfo_screenwidth()
        screen_height = win.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        win.geometry(f'{width}x{height}+{x}+{y}')

    def show_loading_popup(self):
        loading_popup = CTkToplevel(self.ventana, fg_color = color_uno)
        loading_popup.title("Cargando...")
        loading_label = CTkLabel(loading_popup, text="Cargando, por favor espere...")
        loading_label.pack(padx=20, pady=20)
        self.center_window(loading_popup, 300, 100)
        loading_popup.transient(self.ventana)
        loading_popup.grab_set()
        loading_popup.focus()
        return loading_popup

    def run_query(self, query, parameters = ()):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query, parameters)
            conn.commit()
        return result
    
    def bases_operativas(self):
        query = "SELECT proveedor, base FROM bases_operativas"
        db_rows = self.run_query(query)
        bo = []
        for row in db_rows:
            bo.append(f"{row[0]} - {row[1]}")
        return bo
    
    def actualizar_excel(self, fecha, serie, descripcion, mov_ax, base_operativa):
        # Cargar variables del archivo .env
        load_dotenv()

        # Obtener la ruta desde las variables de entorno
        excel_path = os.getenv('EXCEL_PATH')
        try:
            # Cargar el archivo Excel existente
            wb = openpyxl.load_workbook(excel_path)

        except FileNotFoundError:
            # Si el archivo no existe, o su ruta, dar aviso
            return "El archivo no existe."
            return False
        except PermissionError:
            return "No se puede acceder al archivo. Asegúrate de que no esté abierto en otra aplicación."
        except Exception as e:
            return f"Hubo un error inesperado al acceder al archivo: {e}"
        else:
            ws = wb.active
        
        # Agregar el nuevo movimiento
        ws.append([fecha, serie, descripcion, mov_ax, base_operativa])
        
        try:
            # Guardar el archivo Excel
            wb.save(excel_path)
        except PermissionError:
            return "No se puede guardar el archivo. Asegúrate de que no esté abierto en otra aplicación."
        except Exception as e:
            return f"Hubo un error inesperado al guardar el archivo: {e}"
        
        return None  # Si todo fue bien, no retorna nada
    
    def buscar_articulo(self, sn):
        query = "SELECT descripcion FROM equipos WHERE serial_number = ?"
        descripcion = None
        db_rows = self.run_query(query, sn)
        for row in db_rows:
            descripcion = row[0]
        return descripcion
    
    def crear_movimiento(self, mov):
        query = "INSERT INTO movimientos (fecha, tipo_solicitud, serial_number, mov_ax, base_operativa) VALUES (?, ?, ?, ?, ?)"
        # Mostrar la ventana de "Cargando..."
        loading_popup = self.show_loading_popup()
        self.root.update()  # Asegurarse de que la ventana se muestre inmediatamente

        error_message = None
        for item in mov.items:
            fecha_str = mov.fecha.strftime('%d/%m/%Y')
            params = (fecha_str, mov.solicitud, item, mov.mov_ax, mov.bo)
            result = self.actualizar_excel(fecha_str, item, "", mov.mov_ax, mov.bo)
            if result:
                error_message = result  # Guardar el mensaje de error si lo hay
            try:
                self.run_query(query, params)
            except Exception as e:
                print(f"Hubo un error: {e}")

        # Cerrar la ventana de "Cargando..." una vez completado el proceso
        loading_popup.destroy()

        # Mostrar el mensaje de error si hubo alguno
        if error_message:
            messagebox.showerror("Error", error_message, parent = self.ventana)
        
        messagebox.showinfo("Éxito", "Los movimientos se guardaron correctamente.", parent = self.ventana)

    def buscar_movimientos(self, fecha, solicitud):

        fecha = fecha.strftime('%d/%m/%Y')

        if solicitud == "Ingreso" or solicitud == "Egreso":
            query = "SELECT fecha, tipo_solicitud, movimientos.serial_number,descripcion, mov_ax, base_operativa FROM movimientos LEFT JOIN equipos ON movimientos.serial_number = equipos.serial_number WHERE fecha = ? AND tipo_solicitud = ?"
            params = (fecha, solicitud)
        else:
            query = "SELECT fecha, tipo_solicitud, movimientos.serial_number,descripcion, mov_ax, base_operativa FROM movimientos LEFT JOIN equipos ON movimientos.serial_number = equipos.serial_number WHERE fecha = ?"
            params = (fecha, )

        movimientos = []
        db_rows = self.run_query(query, params)
        for row in db_rows:
            movimientos.append(row)
        return movimientos
     
    def open_window(self, movimiento = None):

        global nuevo_movimiento

        # Crear la nueva ventana
        self.ventana = CTkToplevel(self.root, fg_color = color_uno)
        self.ventana.minsize(480,500)
        self.center_window(self.ventana, 500,800)
        self.ventana.grab_set()  # Para asegurarse de que la ventana principal no se pueda usar hasta que se cierre la nueva ventana
        self.ventana.focus()  # Enfocar la nueva ventana

        # Configurar columnas y filas para centrado
        self.ventana.columnconfigure(0, weight=1)
        self.ventana.rowconfigure(0, weight=1)

        # Crear un frame para contener los widgets
        container_frame = CTkFrame(self.ventana, fg_color=color_uno)
        container_frame.grid(column=0, row=0, sticky="nsew", padx=20, pady=20)
        
        container_frame.columnconfigure(0, weight=1)
        container_frame.rowconfigure(0, weight=1)

        # Obtenemos la fecha actual
        fecha_actual = datetime.now()

        # Crear y centrar el titulo del input
        CTkLabel(container_frame, text="Fecha").grid(column=0, row=0, columnspan=2)

        # Estilizar el input de tipo date
        style = ttk.Style(self.ventana)
        style.theme_use('default')
        style.configure('custom.DateEntry',
                        fieldbackground='#272626',
                        background='#4b4a4a',
                        foreground='#ffffff',
                        selectbackground='#00ffff',
                        selectforeground='#ff00aa',
                        arrowcolor='#000000',
                        borderradius=50,
                        borderwidth=0,
                        padding=5)

        # Crear y centrar el input de tipo date
        fecha = DateEntry(container_frame, date_pattern="dd/mm/yyyy", selectmode="day", style='custom.DateEntry')
        fecha.set_date(fecha_actual)
        fecha.grid(column=0, row=1, columnspan=2)

        # Crear y centrar el titulo del input
        CTkLabel(container_frame, text="Tipo de Movimiento").grid(column=0, row=2, columnspan=2)

        # Crear y centrar el input
        
        solicitud = CTkComboBox(container_frame, values=["Ingreso", "Egreso"], variable = StringVar(value="Elige una opcion"))
        solicitud.grid(column=0, row=3, columnspan=2)

        # Estilos para la tabla

        style.configure("Treeview",
            background=gris_claro,
            foreground="#ffffff",
            fieldbackground="#4b4a4a",
            rowheight=25,
            bordercolor="#000000",
            borderwidth=0,
            font = ("Arial", 12))
        style.configure("Treeview.Heading",
                        background=gris_oscuro,
                        foreground=gris_claro,
                        bordercolor="#000000",
                        borderwidth=0,
                        font=('Arial', 13, 'bold'))
        
        # Creamos una fn lambda para ubicar la tabla y botones que aparecen debajo dependiendo que opcion se haya elegido
        
        row_value = lambda: 12 if movimiento else 5



        if movimiento:

            self.ventana.title(f"{movimiento} de Equipos")

            # Si es un movimiento, se setea con el valor correspondiente y se frizza para que no se pueda modificar

            solicitud.set(movimiento)
            solicitud.configure(state="disabled")

            # Crear y centrar el titulo del input
            CTkLabel(container_frame, text="Selecciona la Base Operativa").grid(column=0, row=4, columnspan=2)

            # Crear y centrar el input

            option_bo = StringVar(value = "Elige una opcion")
            bo = CTkComboBox(container_frame, values=self.bases_operativas(), variable=option_bo)
            bo.grid(column=0, row=5, columnspan=2)

            # Crear y centrar el titulo del input
            CTkLabel(container_frame, text="PDT / DDT").grid(column=0, row=6, columnspan=2)

            # Crear y centrar el input
            mov_ax = CTkEntry(container_frame)
            mov_ax.grid(column=0, row=7, columnspan=2)

            def crear_movimiento():

                global nuevo_movimiento

                # Verificar si los campos están completos
                if not fecha.get_date() or not solicitud.get() or bo.get() == "Elige una opcion" or not bo.get() or not mov_ax.get():
                    messagebox.showerror("Error", "Todos los campos deben estar completos", parent = self.ventana)
                    return

                # Crear el objeto Movimiento
                nuevo_movimiento = Movimiento(
                    fecha=fecha.get_date(),
                    solicitud=solicitud.get(),
                    bo=bo.get(),
                    mov_ax=mov_ax.get()
                )

                fecha.config(state = "disabled")
                solicitud.configure(state = "disabled")
                bo.configure(state = "disabled")
                mov_ax.configure(state = "disabled")
                btn_crear_movimiento.configure(state = "disabled")

                serial.configure(state= "normal")
                btn_editar_articulo.configure(state= "normal")
                btn_eliminar_articulo.configure(state= "normal")
                btn_finalizar_mov.configure(state= "normal")

                serial.focus()
                
                messagebox.showinfo("Éxito", "Movimiento creado exitosamente", parent = self.ventana)

            # Crear y centrar el boton de aceptar y crear el movimiento
            btn_crear_movimiento = CTkButton(container_frame, text="Aceptar", command=crear_movimiento)
            btn_crear_movimiento.grid(column=0, row=8, columnspan=2)

            # Crear y centrar el input para la serie
            CTkLabel(container_frame, text="Coloque aqui el serial number").grid(column=0, row=9)

            serial = CTkEntry(container_frame)
            serial.grid(column=0, row=10)
            serial.configure(state= "disabled")

            # Creamos un texto para mostrar el total de items agregados al movimiento

            total_items = CTkLabel(container_frame, text = "")
            total_items.grid(column = 0, row = 11)

            # Manejar el evento de entrada de datos en el campo serial

            def on_serial_enter(event):
                serial_num = serial.get().strip()
                if serial_num:
                    if serial_num in nuevo_movimiento.items:
                        messagebox.showerror("Error", "No podes ingresar dos serial number iguales en un mismo movimiento", parent = self.ventana)
                        return
                    desc = self.buscar_articulo((serial_num,))
                    if desc is None:
                        desc = "Equipo no encontrado en la tabla Equipos"
                    nuevo_movimiento.agregar_items(serial_num)
                    tab.insert("", "end", text=serial_num, values=(desc,))
                    serial.delete(0, END)
                    serial.focus()
                    total_items.configure(text = f"Total de items: {nuevo_movimiento.total_items()}")
                else:
                    messagebox.showerror("Error", "Debes colocar el serial number", parent = self.ventana)


            # Asociar la función on_serial_enter al evento Return (Enter)
            
            serial.bind("<Return>", on_serial_enter)

            # Crear y centrar la tabla para movimientos ingreso / egreso

            tab = ttk.Treeview(container_frame, height=10, columns=2)
            tab.grid(column=0, row= row_value(), columnspan=2, sticky="nsew")
            tab.heading("#0", text="N° de Serie", anchor=CENTER)
            tab.heading("#1", text="Descripción", anchor=CENTER)
            tab.column("#0", anchor=CENTER)
            tab.column("#1", anchor=CENTER)

            def crear_movimiento():

                if nuevo_movimiento.total_items() > 0:
                    self.crear_movimiento(nuevo_movimiento)
                    self.ventana.destroy()
                else:
                    messagebox.showerror("Error", "Debe agregar al menos un item en el movimiento", parent = self.ventana)

                    # Creamos la fn para eliminar un articulo de la tabla

            def borrar_articulo():
                global nuevo_movimiento

                try:
                    selected_item = tab.selection()[0]  # Obtén el primer ítem seleccionado
                    serie = tab.item(selected_item)["text"]
                except IndexError:
                    messagebox.showerror("Error", "Debes seleccionar un ítem para eliminar", parent = self.ventana)
                    return

                if serie in nuevo_movimiento.items:
                    nuevo_movimiento.eliminar_item(serie)
                    tab.delete(selected_item)  # Elimina el ítem del Treeview
                    total_items.configure(text=f"Total de items: {nuevo_movimiento.total_items()}")
                else:
                    messagebox.showerror("Error", "El artículo seleccionado no está en la lista de movimiento", parent = self.ventana)

            # Creamos la fn para modificar un articulo de la tabla

            def modificar_articulo():
                global nuevo_movimiento

                try:
                    selected_item = tab.selection()[0]  # Obtén el primer ítem seleccionado
                    serie_antigua = tab.item(selected_item)["text"]
                except IndexError:
                    messagebox.showerror("Error", "Debes seleccionar un ítem para editar", parent = self.ventana)
                    return
                
                ventana_mod = CTkToplevel(self.ventana, fg_color = color_uno)
                ventana_mod.geometry("300x350")
                ventana_mod.minsize(200,350)
                ventana_mod.transient(self.ventana)  # Para que la nueva ventana se abra sobre la ventana principal
                ventana_mod.grab_set()  # Para asegurarse de que la ventana principal no se pueda usar hasta que se cierre la nueva ventana

                CTkLabel(ventana_mod, text = "Serie a modificar: ").grid(column = 0, row = 0)

                serie_anterior = CTkEntry(ventana_mod, textvariable = StringVar(value = serie_antigua))
                serie_anterior.grid(column = 1, row = 0)
                serie_anterior.configure(state = "disabled")

                CTkLabel(ventana_mod, text = "Nueva Serie: ").grid(column = 0, row = 1)

                serie_nueva = CTkEntry(ventana_mod)
                serie_nueva.grid(column = 1, row = 1)

                def modificar_items():

                    valor_serie_anterior = serie_anterior.get().strip()
                    valor_serie_nueva = serie_nueva.get().strip()

                    if not valor_serie_nueva:
                        messagebox.showerror("Error", "Debe seleccionar un nuevo serial number", parent = ventana_mod)
                        return

                    if valor_serie_anterior in nuevo_movimiento.items:

                        if valor_serie_nueva in nuevo_movimiento.items:
                            messagebox.showerror("Error", "No podes ingresar dos serial number iguales en un mismo movimiento", parent = ventana_mod)
                            return
                        
                        nuevo_movimiento.modificar_items(valor_serie_anterior, valor_serie_nueva)

                        desc = self.buscar_articulo((valor_serie_nueva,))

                        if desc is None:
                            desc = "Equipo no encontrado en la tabla Equipos"

                        tab.delete(selected_item)  # Elimina el ítem del Treeview
                        tab.insert("", "end", text = valor_serie_nueva, values=(desc,))

                        total_items.configure(text=f"Total de items: {nuevo_movimiento.total_items()}")
                        messagebox.showinfo("Success", "El articulo a sido modificado exitosamente", parent = ventana_mod)
                        ventana_mod.destroy()
                    else:
                        messagebox.showerror("Error", "El artículo seleccionado no está en la lista de movimiento", parent = ventana_mod)

                btn_modificar_item = CTkButton(ventana_mod, text = "Confirmar modificacion", command = modificar_items)
                btn_modificar_item.grid(column = 0, row = 3, columnspan= 2, sticky = "nsew")

            # Creamos el boton para modificar articulo

            btn_editar_articulo = CTkButton(container_frame, text = "Editar Articulo", fg_color = "#9b5f1a", command = modificar_articulo)
            btn_editar_articulo.grid(column = 0, row = row_value() + 1, columnspan = 2)
            btn_editar_articulo.configure(state = "disabled")

            # Creamos el boton para eliminar articulo

            btn_eliminar_articulo = CTkButton(container_frame, text = "Eliminar Articulo", fg_color = "#d32a2a", command = borrar_articulo)
            btn_eliminar_articulo.grid(column = 0, row = row_value() + 2, columnspan = 2)
            btn_eliminar_articulo.configure(state = "disabled")

            # Creamos el boton para finalizar movimiento

            btn_finalizar_mov = CTkButton(container_frame, text = "Finalizar movimiento", fg_color = "#39863f", command = crear_movimiento)
            btn_finalizar_mov.grid(column = 0, row = row_value() + 3, columnspan = 2)
            btn_finalizar_mov.configure(state = "disabled")

        else:
                        
            # Seteamos el titulo de la ventana ya que se selecciono la ventana de "Ver Movimientos"
            
            self.ventana.title(f"Buscar movimiento de Equipos")

            def buscar_movimientos():

                date = fecha.get_date()
                soli = solicitud.get()

                movimientos = self.buscar_movimientos(date, soli)

                tab.delete(*tab.get_children())
                for movimiento in movimientos:
                    tab.insert("", "end", values = movimiento)

            btn_buscar_movimiento = CTkButton(container_frame, text="Buscar", command = buscar_movimientos)
            btn_buscar_movimiento.grid(column=0, row=4, columnspan=2)

            # Crear y centrar la tabla para movimientos ingreso / egreso

            tab = ttk.Treeview(container_frame, height=10, columns=("fecha", "solicitud", "serie", "desc", "bo"), show="headings")
            tab.grid(column=0, row= row_value(), columnspan=2, sticky="nsew")
            tab.heading("fecha", text="Fecha", anchor=CENTER)
            tab.heading("solicitud", text="Tipo de Solicitud", anchor=CENTER)
            tab.heading("serie", text="N° de Serie", anchor=CENTER)
            tab.heading("desc", text="Descripcion", anchor=CENTER)
            tab.heading("bo", text="Base Operativa", anchor=CENTER)
            tab.column("fecha", width = 50, anchor=CENTER)
            tab.column("solicitud", width = 50, anchor=CENTER)
            tab.column("serie", width = 100, anchor=CENTER)
            tab.column("desc", width = 100, anchor=CENTER)
            tab.column("bo", width = 100, anchor=CENTER)

        # Configurar para expandir la tabla
        
        container_frame.rowconfigure(6, weight=1)


if __name__ == "__main__":

    window = customtkinter.CTk()
    application = App(window)
    window.mainloop()
