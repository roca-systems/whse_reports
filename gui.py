import os
import sys
import logging
import customtkinter
import re
import json
import pandas as pd
from pathlib import Path
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox, simpledialog
from datetime import datetime
from comparacion_poo import TxtProcessor, ExcelGenerator, EmailSender, JetAccessBot # Asumiendo que el script ComparacionSinSolicitudDeFechasPOO.py se guarda como comparacion_poo.py

# =================================================================================
# Funciones Auxiliares
# =================================================================================

def resource_path(relative_path):
    """Get the absolute path to the resource, works for PyInstaller."""
    if hasattr(sys, '_MEIPASS'):  # PyInstaller creates a temporary folder for bundled files
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def load_destinatarios():
    """Carga la lista de correos de destinatarios desde el archivo JSON."""
    try:
        config_path = resource_path("destinatarios.json")
        with open(config_path, "r", encoding="utf-8") as file:
            return json.load(file).get("correos", [])
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"No se pudo cargar destinatarios.json: {e}")
        return []

def save_destinatarios(correos):
    """Guarda la lista de correos de destinatarios en el archivo JSON."""
    try:
        config_path = resource_path("destinatarios.json")
        with open(config_path, "w", encoding="utf-8") as file:
            json.dump({"correos": correos}, file, indent=4, ensure_ascii=False)
        messagebox.showinfo("Información", "Destinatarios Guardados con Éxito.")
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"No se pudo guardar destinatarios.json: {e}")

def load_prioridades():
    """Carga el diccionario de prioridades desde el archivo JSON."""
    try:
        config_path = resource_path("promedio_folios_con_importancia_MACDERMID_DE_MEXICO.json")
        with open(config_path, "r", encoding="utf-8") as file:
            return json.load(file)
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"No se pudo cargar el archivo de prioridades:\n{e}")
        return {}

def save_prioridades(data):
    """Guarda el diccionario de prioridades en el archivo JSON."""
    try:
        config_path = resource_path("promedio_folios_con_importancia_MACDERMID_DE_MEXICO.json")
        with open(config_path, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
        messagebox.showinfo("Información", "Prioridades Guardadas con Éxito.")
    except Exception as e:
        messagebox.showerror("Error de Configuración", f"No se pudo guardar el archivo de prioridades:\n{e}")

# =================================================================================
# Clase Principal de la Aplicación
# =================================================================================

class App(customtkinter.CTk):

    def __init__(self):
        # --- Configuración de Usuarios ---
        self.valid_users = {
            'admin': 'admin',
            'emorgan':'sls1',
            'ehernandez':'sls1',
            'rsolis':'sls1',
            'lvalera':'sls1',
            'earteaga':'sls1',
        }

        super().__init__()
        self.title("ROCA - ROBOT PROCESS AUTOMATION")
        try:
            self.wm_iconbitmap(resource_path("icons/roca_icon-2.ico"))
        except Exception as e:
            logging.error(f"Failed to load icon: {e}")

        self.geometry("1050x550")
        self.resizable(False, False)

        # Initialize frames
        self.login_frame = customtkinter.CTkFrame(self, width=550, height=550)
        self.login_frame.grid(row=0, column=0, sticky="nsew")
        self.login_frame.grid_propagate(True)

        self.main_frame = customtkinter.CTkFrame(self, width=550, height=550)
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.main_frame.grid_propagate(True)

        self.destinatarios_frame = customtkinter.CTkFrame(self, width=550, height=550)
        self.destinatarios_frame.grid(row=0, column=0, sticky="nsew")
        self.destinatarios_frame.grid_propagate(True)

        self.prioridad_frame = customtkinter.CTkFrame(self, width=550, height=550)
        self.prioridad_frame.grid(row=0, column=0, sticky="nsew")
        self.prioridad_frame.grid_propagate(True)

        # Frame para la imagen lateral
        self.side_image_frame= customtkinter.CTkFrame(self, width=500, height=550)
        self.side_image_frame.grid(row=0, column=2, sticky="nsew")
        self.side_image_frame.grid_propagate(True)

        # Configuración de columnas y filas
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(2, weight=1)

        self.show_login_frame()


    def hide_all_frames(self):
        self.login_frame.grid_forget()
        self.main_frame.grid_forget()
        self.destinatarios_frame.grid_forget()
        self.prioridad_frame.grid_forget()
        for image in self.side_image_frame.winfo_children():
            image.destroy()


    # ==============================
    # 1. Login Frame
    # ==============================
    def show_login_frame(self):
        self.hide_all_frames()
        self.login_frame.grid(row=0, column=0, sticky="nsew")
        self.side_image_frame.grid(row=0, column=2)

        # Welcome Image
        welcome_image = customtkinter.CTkImage(light_image=Image.open(resource_path('icons/login.PNG')), size=(550 , 550))
        welcome_label = customtkinter.CTkLabel(self.side_image_frame, image=welcome_image, text='')
        welcome_label.pack(side="right", fill='both', expand=True)

        # Widgets
        for widget in self.login_frame.winfo_children():
            widget.destroy()

        title = customtkinter.CTkLabel(self.login_frame, text="\n GENERACIÓN DE REPORTES \n DRP \n", font=("Arial", 25))
        title.grid(row=0, column=0, padx=50, pady=20, columnspan=2)

        login_username_label = customtkinter.CTkLabel(self.login_frame, text="Usuario:", font=("Arial", 20))
        login_username_label.grid(row=1, column=0, padx=50, pady=5)
        self.login_username_entry = customtkinter.CTkEntry(self.login_frame, width=200)
        self.login_username_entry.grid(row=1, column=1, padx=50, pady=20, sticky="w")

        login_password_label = customtkinter.CTkLabel(self.login_frame, text="Password:", font=("Arial", 20))
        login_password_label.grid(row=2, column=0, padx=50, pady=5)
        self.login_password_entry = customtkinter.CTkEntry(self.login_frame, width=200, show="*")
        self.login_password_entry.grid(row=2, column=1, padx=50, pady=20, sticky="w")

        login_button = customtkinter.CTkButton(self.login_frame, text="Login", command=self.validate_login, font=("Arial", 15))
        login_button.grid(row=3, column=0, pady=30, columnspan=2)

        instructions = "Bienvenido al Generador de Reportes de DRP."
        title = customtkinter.CTkLabel(self.login_frame, text=instructions, font=("Arial", 15))
        title.grid(row=4, column=0, padx=50, pady=20, columnspan=2)

    def validate_login(self):
        username = self.login_username_entry.get()
        password = self.login_password_entry.get()

        if username in self.valid_users and self.valid_users[username] == password:
            self.show_main_frame()
        else:
            messagebox.showerror("Error: Acceso Incorrecto", "Usuario o Contraseña Incorrectos")
            logging.error(f"Intento de Login Fallido para el Usuario: {username}")


    # ==============================
    # 2. Main Frame (Ejecución del Bot)
    # ==============================
    def show_main_frame(self):
        self.hide_all_frames()
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.side_image_frame.grid(row=0, column=1)

        # Instructions Image
        instructions_image = customtkinter.CTkImage(light_image=Image.open(resource_path('icons/DRP.png')), size=(400 , 550)) # Usando como placeholder
        instructions_label = customtkinter.CTkLabel(self.side_image_frame, image=instructions_image, text='')
        instructions_label.pack(side="right", fill='both', expand=True)

        # Widgets
        for widget in self.main_frame.winfo_children():
            widget.destroy()

        task_title_label = customtkinter.CTkLabel(self.main_frame, text="\n GENERACIÓN DE REPORTES DRP \n COMPARATIVO DE FOLIOS DE EMBARQUES \n", font=("Arial", 25))
        task_title_label.grid(row=0, column=0, padx=50, pady=10, columnspan=2)

        # 1. Usuario JetAccess
        user_label = customtkinter.CTkLabel(self.main_frame, text="Usuario JetAccess:", font=("Arial", 18))
        user_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.jetaccess_user_entry = customtkinter.CTkEntry(self.main_frame, width=250)
        self.jetaccess_user_entry.grid(row=1, column=1, padx=20, pady=10, sticky="w")

        # 2. Contraseña JetAccess
        pwd_label = customtkinter.CTkLabel(self.main_frame, text="Contraseña JetAccess:", font=("Arial", 18))
        pwd_label.grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.jetaccess_pwd_entry = customtkinter.CTkEntry(self.main_frame, width=250, show="*")
        self.jetaccess_pwd_entry.grid(row=2, column=1, padx=20, pady=10, sticky="w")

        # Botón de correr el bot
        self.run_button = customtkinter.CTkButton(self.main_frame, text="INICIAR BOT DRP", command=self.submit_run_bot)
        self.run_button.grid(row=3, column=0, columnspan=2, pady=30)

        # Botón de Editar Destinatarios
        self.destinatarios_button = customtkinter.CTkButton(self.main_frame, text="EDITAR DESTINATARIOS", command=self.show_destinatarios_frame)
        self.destinatarios_button.grid(row=4, column=0, columnspan=2, pady=10)

        # Botón de Editar Prioridad
        self.prioridad_button = customtkinter.CTkButton(self.main_frame, text="EDITAR PRIORIDAD", command=self.show_prioridad_frame)
        self.prioridad_button.grid(row=5, column=0, columnspan=2, pady=10)


    def submit_run_bot(self):
        user = self.jetaccess_user_entry.get().strip()
        pwd = self.jetaccess_pwd_entry.get().strip()

        if not user or not pwd:
            messagebox.showerror("Error: Validación", "Por favor ingrese Usuario y Contraseña de JetAccess.")
            return

        # Deshabilitar el botón
        self.run_button.configure(state="disabled", text="EJECUTANDO...")

        # Iniciar el bot en un hilo separado o usando after para no bloquear la GUI
        self.after(100, self.run_bot, user, pwd)


    def run_bot(self, user, pwd):
        # Deshabilitar el botón para evitar ejecuciones múltiples
        self.run_button.configure(state="disabled", text="EJECUTANDO...")

        try:
            # 1. Ejecutar el bot para descargar el archivo TXT
            jet_bot = JetAccessBot()
            archivo_txt = jet_bot.ejecutar_descarga(user, pwd)

            if not archivo_txt:
                messagebox.showerror("Error", "No se pudo descargar el archivo de reporte TXT.")
                return

            # 2. Procesar el TXT para obtener el conteo y la lista de clientes
            txt_processor = TxtProcessor(archivo_txt)
            conteo_por_shipto, clientes_unicos = txt_processor.contar_folios_por_shipto_clientes() # Modificación a la clase para devolver lista de clientes

            # 3. Mostrar la ventana emergente para la selección de cliente
            self.show_client_selection_popup(conteo_por_shipto, clientes_unicos)

        except Exception as e:
            messagebox.showerror("Error: Ejecución del Bot", f"Ocurrió un error en la ejecución principal: {e}")
        finally:
            # Re-habilitar el botón
            self.run_button.configure(state="normal", text="INICIAR BOT DE REPORTE")


    def show_client_selection_popup(self, conteo_actual, clientes_unicos):
        """Muestra una ventana emergente para que el usuario seleccione un cliente."""
        self.clientes_unicos = clientes_unicos
        self.conteo_actual = conteo_actual

        popup = customtkinter.CTkToplevel(self)
        popup.title("Selección de Cliente")
        popup.geometry("400x300")
        popup.grab_set() # Modal behavior

        customtkinter.CTkLabel(popup, text="Clientes encontrados:", font=("Arial", 18)).pack(pady=10)

        # Crear lista desplegable de clientes
        opciones = ["* - Todos los Clientes"] + [f"{i}. {c}" for i, c in enumerate(clientes_unicos, 1)]
        self.cliente_seleccionado = customtkinter.StringVar(value=opciones[0])
        client_menu = customtkinter.CTkOptionMenu(popup, values=opciones, variable=self.cliente_seleccionado, width=350)
        client_menu.pack(pady=5)

        # Botón de Continuar
        customtkinter.CTkButton(popup, text="Continuar", command=lambda: self.process_client_selection(popup)).pack(pady=20)


    def process_client_selection(self, popup):
        """Procesa la selección del cliente y continúa con la generación del reporte y el envío de correo."""
        popup.destroy()
        selection_text = self.cliente_seleccionado.get()

        if selection_text.startswith("*"):
            seleccion = "*"
        else:
            # Extraer el nombre del cliente de la opción seleccionada (ej: "1. MACDERMID DE MEXICO SA DE CV" -> "MACDERMID DE MEXICO SA DE CV")
            match = re.search(r'\. (.*)', selection_text)
            seleccion = match.group(1).strip() if match else None

        if seleccion is None:
            messagebox.showerror("Error", "Selección de cliente inválida. Cancelando el proceso.")
            self.run_button.configure(state="normal", text="INICIAR BOT DRP")
            return

        try:
            # Obtener el reporte final
            self.generar_y_enviar_reporte(self.conteo_actual, seleccion)

        except Exception as e:
            messagebox.showerror("Error: Generación de Reporte", f"Ocurrió un error al generar/enviar el reporte: {e}")
        finally:
            self.run_button.configure(state="normal", text="INICIAR BOT DRP")
            messagebox.showinfo("Proceso Terminado", "El flujo principal de generación de reportes ha terminado.")


    def generar_y_enviar_reporte(self, conteo_actual, seleccion):
        """Lógica para generar el Excel y enviar el correo."""
        archivos_generados = []
        RUTA_REPORTES = Path(__file__).parent / "reportes_finales"
        RUTA_REPORTES.mkdir(exist_ok=True)

        if seleccion == "*":
            # Si se selecciona '*', el original procesaba un solo JSON de ejemplo (MACDERMID_DE_MEXICO)
            # Para simplificar y seguir el flujo del original:
            json_path = resource_path("promedio_folios_con_importancia_MACDERMID_DE_MEXICO.json")
            excel_output = RUTA_REPORTES / f"comparacion_folios_TODOS_{datetime.now().strftime('%Y%m%d')}.xlsx"
            archivos_generados.append(ExcelGenerator(json_path, conteo_actual, excel_output).generar())
        else:
            # Lógica para un cliente específico (ej. si existiera un JSON por cliente)
            # Siguiendo el ejemplo del script original, se usa la plantilla con el nombre del cliente
            cliente_sanitizado = re.sub(r'[^A-Za-z0-9_-]+', '_', seleccion)
            json_path = resource_path("promedio_folios_con_importancia_MACDERMID_DE_MEXICO.json") # Usamos el único JSON disponible como plantilla
            excel_output = RUTA_REPORTES / f"comparacion_folios_{cliente_sanitizado}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            archivos_generados.append(ExcelGenerator(json_path, conteo_actual, excel_output).generar())


        # Enviar correo
        destinatarios = load_destinatarios()
        if destinatarios:
            emailer = EmailSender("analytics@slservices.com.mx", "eyxf abfk bvct dclj", destinatarios)
            emailer.enviar(archivos_generados)
        else:
            messagebox.showwarning("Advertencia de Correo", "No se encontraron destinatarios. No se enviará el correo.")


    # ==============================
    # 3. Destinatarios Frame
    # ==============================
    def show_destinatarios_frame(self):
        self.hide_all_frames()
        self.destinatarios_frame.grid(row=0, column=0, sticky="nsew")
        self.side_image_frame.grid(row=0, column=2)

        # Instructions Image (Placeholder)
        instructions_image = customtkinter.CTkImage(light_image=Image.open(resource_path('icons/DRP.png')), size=(500 , 550))
        instructions_label = customtkinter.CTkLabel(self.side_image_frame, image=instructions_image, text='')
        instructions_label.pack(side="right", fill='both', expand=True)

        # Widgets
        for widget in self.destinatarios_frame.winfo_children():
            widget.destroy()

        task_title_label = customtkinter.CTkLabel(self.destinatarios_frame, text="\n EDITAR DESTINATARIOS \n", font=("Arial", 25))
        task_title_label.grid(row=0, column=0, padx=50, pady=10, columnspan=2)

        current_destinatarios = "\n".join(load_destinatarios())

        customtkinter.CTkLabel(self.destinatarios_frame, text="Correos (uno por línea):", font=("Arial", 18)).grid(row=1, column=0, padx=50, pady=5, sticky="nw")
        self.destinatarios_textbox = customtkinter.CTkTextbox(self.destinatarios_frame, width=400, height=200)
        self.destinatarios_textbox.insert("0.0", current_destinatarios)
        self.destinatarios_textbox.grid(row=2, column=0, padx=50, pady=20, columnspan=2)

        # Botones
        self.save_button = customtkinter.CTkButton(self.destinatarios_frame, text="GUARDAR", command=self.save_destinatarios_gui)
        self.save_button.grid(row=3, column=0, pady=20, padx=20)
        self.cancel_button = customtkinter.CTkButton(self.destinatarios_frame, text="CANCELAR", command=self.show_main_frame)
        self.cancel_button.grid(row=3, column=1, pady=20, padx=20)

    def save_destinatarios_gui(self):
        text = self.destinatarios_textbox.get("1.0", "end-1c").strip()
        correos = [c.strip() for c in text.split('\n') if c.strip()]
        save_destinatarios(correos)
        self.show_main_frame()


    # ==============================
    # 4. Prioridad Frame (Simplificada para editar un solo registro)
    # ==============================
    def show_prioridad_frame(self):
        self.hide_all_frames()
        self.prioridad_frame.grid(row=0, column=0, sticky="nsew")
        self.side_image_frame.grid(row=0, column=2)

        # Instructions Image (Placeholder)
        instructions_image = customtkinter.CTkImage(light_image=Image.open(resource_path('icons/DRP.png')), size=(500 , 550))
        instructions_label = customtkinter.CTkLabel(self.side_image_frame, image=instructions_image, text='')
        instructions_label.pack(side="right", fill='both', expand=True)

        # Widgets
        for widget in self.prioridad_frame.winfo_children():
            widget.destroy()

        task_title_label = customtkinter.CTkLabel(self.prioridad_frame, text="\n EDITOR DE PRIORIDADES \n(Promedio Folios e Importancia) \n", font=("Arial", 25))
        task_title_label.grid(row=0, column=0, padx=50, pady=10, columnspan=2)

        self.prioridad_data = load_prioridades()
        clientes = sorted(self.prioridad_data.keys())

        customtkinter.CTkLabel(self.prioridad_frame, text="Seleccione Cliente:", font=("Arial", 18)).grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.cliente_prioridad_var = customtkinter.StringVar(value=clientes[0] if clientes else "")
        self.cliente_prioridad_menu = customtkinter.CTkOptionMenu(self.prioridad_frame, values=clientes, variable=self.cliente_prioridad_var, command=self.load_prioridad_details, width=350)
        self.cliente_prioridad_menu.grid(row=1, column=1, padx=20, pady=10, sticky="w")

        # Controles para Promedio y Prioridad
        customtkinter.CTkLabel(self.prioridad_frame, text="Promedio Folios:", font=("Arial", 15)).grid(row=2, column=0, padx=20, pady=5, sticky="w")
        self.promedio_entry = customtkinter.CTkEntry(self.prioridad_frame, width=100)
        self.promedio_entry.grid(row=2, column=1, padx=20, pady=5, sticky="w")

        customtkinter.CTkLabel(self.prioridad_frame, text="Importancia:", font=("Arial", 15)).grid(row=3, column=0, padx=20, pady=5, sticky="w")
        self.importancia_var = customtkinter.StringVar()
        importancia_options = ["ALTA PRIORIDAD - Destinos Críticos", "MEDIA PRIORIDAD - Destinos Regulares", "BAJA PRIORIDAD - Destinos Ocasionales"]
        self.importancia_menu = customtkinter.CTkOptionMenu(self.prioridad_frame, values=importancia_options, variable=self.importancia_var, width=350)
        self.importancia_menu.grid(row=3, column=1, padx=20, pady=5, sticky="w")

        self.load_prioridad_details(self.cliente_prioridad_var.get())

        # Botones
        self.save_button = customtkinter.CTkButton(self.prioridad_frame, text="GUARDAR CAMBIOS", command=self.save_prioridad_gui)
        self.save_button.grid(row=4, column=0, pady=20, padx=20)
        self.cancel_button = customtkinter.CTkButton(self.prioridad_frame, text="CANCELAR", command=self.show_main_frame)
        self.cancel_button.grid(row=4, column=1, pady=20, padx=20)


    def load_prioridad_details(self, cliente):
        """Carga los detalles de promedio y prioridad para el cliente seleccionado."""
        if cliente and cliente in self.prioridad_data:
            data = self.prioridad_data[cliente]
            self.promedio_entry.delete(0, customtkinter.END)
            self.promedio_entry.insert(0, str(data.get("Promedio_Folios", 0)))
            self.importancia_var.set(data.get("Importancia", "BAJA PRIORIDAD - Destinos Ocasionales"))

    def save_prioridad_gui(self):
        """Guarda los cambios de prioridad en la estructura de datos y en el JSON."""
        cliente = self.cliente_prioridad_var.get()
        promedio_str = self.promedio_entry.get().strip()

        try:
            promedio = float(promedio_str)
        except ValueError:
            messagebox.showerror("Error de Validación", "El Promedio de Folios debe ser un número.")
            return

        importancia = self.importancia_var.get()

        if cliente and cliente in self.prioridad_data:
            self.prioridad_data[cliente]["Promedio_Folios"] = promedio
            self.prioridad_data[cliente]["Importancia"] = importancia
            save_prioridades(self.prioridad_data)
            self.show_main_frame()
        else:
            messagebox.showerror("Error", "No se encontró el cliente en los datos.")

# =================================================================================
# Ejecución Principal
# =================================================================================

if __name__ == "__main__":
    # La clase MainWorkflow del archivo ComparacionSinSolicitudDeFechasPOO.py necesita ser modificada
    # para no solicitar input() y, en su lugar, aceptar los parámetros user, pwd, y el archivo_txt
    # para que la GUI pueda controlar el flujo.
    # Se asume una modificación a las clases de comparacion.py (aquí importadas como comparacion_poo)
    # y en particular a TxtProcessor para que devuelva la lista de clientes.

    # Es CRUCIAL que las clases del script ComparacionSinSolicitudDeFechasPOO.py estén disponibles
    # en un archivo importable (ej. 'comparacion_poo.py') y que se realicen las adaptaciones
    # necesarias para que no utilicen input() directamente, sino que reciban los valores
    # desde la GUI (como se ejemplifica en run_bot y show_client_selection_popup).
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        messagebox.showerror("Error Fatal", f"Ocurrió un error inesperado al iniciar la aplicación:\n{e}")
        logging.error(f"Error fatal al iniciar: {e}")