import os
import re
import time
import glob
import tempfile
from pathlib import Path

# Importaciones de Selenium WebDriver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import Select

# ==============================
# Configuración de Rutas (Auxiliar)
# ==============================
RUTA_DESCARGA = os.path.join(tempfile.gettempdir(), "descargas_rpa")
os.makedirs(RUTA_DESCARGA, exist_ok=True)


# ==============================
# Clase para automatizar JetAccess (CON SELENIUM PURO)
# ... (Todo el código de la clase JetAccessBot se mantiene igual)
# ==============================
class JetAccessBot:
    def __init__(self):
        # 1. Configurar opciones y preferencias de descarga para Edge
        edge_options = webdriver.EdgeOptions()
        
        # Preferencias para descarga y seguridad
        prefs = {
            "download.default_directory": RUTA_DESCARGA,
            "download.prompt_for_download": False,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1, 
        }
        edge_options.add_experimental_option("prefs", prefs)
        
        # Argumentos para el driver
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--start-maximized')
        edge_options.add_argument('--window-size=1920,1080')
        
        # 2. Inicializar el Edge driver y el objeto de espera
        self.driver = webdriver.Edge(options=edge_options)
        self.wait = WebDriverWait(self.driver, 30)
        self.archivo_xls = None
        
    def ejecutar_descarga(self, usuario, contrasena):
        print("Iniciando automatización de JetAccess...")
        try:
            # Abrir navegador y entrar
            self.driver.get("https://www.jetaccess.com")
            
            # Login inicial
            self.driver.find_element(By.ID, "idlogin").send_keys("scc")
            self.driver.find_element(By.ID, "idlogin").send_keys(webdriver.Keys.ENTER)
            time.sleep(2)
            
            # Cambiar a la nueva ventana de login
            ventanas = self.driver.window_handles
            if len(ventanas) > 1:
                self.driver.switch_to.window(ventanas[-1])
            else:
                 print("Advertencia: No se detectó una ventana de login secundaria.")

            # Login con credenciales
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.1.3.5')))
            
            self.driver.find_element(By.NAME, '3.1.3.5').send_keys(usuario)
            self.driver.find_element(By.NAME, '3.1.3.9').send_keys(contrasena)
            self.driver.find_element(By.NAME, '3.1.3.11').click()
            print("Login completado.")

            # Navegación para llegar al reporte
            self.wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Inventario"))).click()
            
            # Click en Reportes
            xpath_reportes = "//a[font[contains(text(), 'Reportes')]]"
            self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_reportes))).click()

            # Seleccionar opción 3 y descargar directamente <select name="3.0.1.3.7.InventoryReportsGrouped.1.3.7" size="15"><option selected="selected" value="0">Nivel de Inventario</option>
            select_element = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.3.7')))
            selector = Select(select_element)
            selector.select_by_index(0)

            # Seleccionar siguiente
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.3.9'))).click()

            # Seleccionar Queretaro y descargar directamente
            select_element2 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector2 = Select(select_element2)
            selector2.select_by_index(2)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.13.3'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.13.7'))).click()

            # Seleccionar Monterrey y descargar directamente
            select_element3 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector3 = Select(select_element3)
            selector3.select_by_index(3)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Guadalajara y descargar directamente
            select_element4 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector4 = Select(select_element4)
            selector4.select_by_index(4)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Hermosillo y descargar directamente
            select_element5 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector5 = Select(select_element5)
            selector5.select_by_index(5)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Tijuana y descargar directamente
            select_element6 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector6 = Select(select_element6)
            selector6.select_by_index(6)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Nogales y descargar directamente
            select_element7 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector7 = Select(select_element7)
            selector7.select_by_index(7)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Monterrey_GC y descargar directamente
            select_element8 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector8 = Select(select_element8)
            selector8.select_by_index(8)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Tultitlan y descargar directamente
            select_element9 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector9 = Select(select_element9)
            selector9.select_by_index(9)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()

            # Seleccionar Ciudad Juarez y descargar directamente
            select_element10 = self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19')))
            selector10 = Select(select_element10)
            selector10.select_by_index(10)
            # Click en el botón "Siguiente" antes de exportar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'))).click()
            # Descargar XLS
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29'))).click()
            time.sleep(3)
            # Clic en regresar
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'))).click()


        except TimeoutException as e:
            print(f"❌ Error de tiempo de espera: El elemento no apareció. {e}")
        except Exception as e:
            print(f"❌ Error inesperado durante la ejecución: {e}")
        #finally:
             #self.cerrar()

        return self.archivo_xls


    def _wait_for_txt_download(self):
        """Espera a que el archivo TXT se descargue en la carpeta temporal."""
        archivo_xls = None
        inicio = time.time()
        print(f"Esperando descarga en: {RUTA_DESCARGA}")
        while time.time() - inicio < 60:
            xls = glob.glob(os.path.join(RUTA_DESCARGA, "*.xls"))
            if xls:
                archivo_xls = max(xls, key=os.path.getctime)
                break
            time.sleep(1)
        return archivo_xls

    # def cerrar(self):
    #     """Cierra el driver si está activo."""
    #     if hasattr(self, 'driver') and self.driver:
    #         print("Cerrando navegador.")
    #         self.driver.quit()


# ==============================
# BLOQUE DE PRUEBA INDEPENDIENTE
# ==============================
if __name__ == "__main__":
    
    # 🔑 CAMBIO CLAVE: Solicitar credenciales en la terminal
    print("--- PRUEBA DE AUTOMATIZACIÓN JETACCESS (SELENIUM) ---")
    USER = input("Ingrese su usuario de JetAccess: ")
    PASS = input("Ingrese su contraseña de JetAccess: ")
    
    bot = JetAccessBot()
    try:
        ruta_descargada = bot.ejecutar_descarga(USER, PASS)
        if ruta_descargada:
            print(f"\nProceso finalizado. Archivo procesable: {ruta_descargada}")
        else:
            print("\nProceso finalizado sin archivo descargado.")
    except Exception as e:
        print(f"\nError durante la prueba: {e}")