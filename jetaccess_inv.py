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
# Configuración de Rutas
# ==============================
RUTA_DESCARGA = os.path.join(tempfile.gettempdir(), "descargas_rpa")
os.makedirs(RUTA_DESCARGA, exist_ok=True)


# ==============================
# Clase para automatizar JetAccess (CON SELENIUM PURO)
# ==============================
class JetAccessBot:
    def __init__(self):
        # 1. Configurar opciones y preferencias de descarga para Edge
        edge_options = webdriver.EdgeOptions()
        
        prefs = {
            "download.default_directory": RUTA_DESCARGA,
            "download.prompt_for_download": False,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
            "profile.default_content_setting_values.automatic_downloads": 1, 
        }
        edge_options.add_experimental_option("prefs", prefs)
        
        # Argumentos
        edge_options.add_argument('--no-sandbox')
        edge_options.add_argument('--disable-dev-shm-usage')
        edge_options.add_argument('--disable-gpu')
        edge_options.add_argument('--start-maximized')
        edge_options.add_argument('--window-size=1920,1080')
        
        # 2. Inicializar el Edge driver
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
            inventario_element = self.wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Inventario")))
            # Clic en Inventario (forzado para evitar intercepción)
            self.driver.execute_script("arguments[0].click();", inventario_element)
            
            # Clic en Consignación a la Venta (uso de XPath y clic forzado)
            xpath_consignacion = "//a[font[contains(text(), 'Consignación a la Venta')]]"
            try:
                consignacion_element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_consignacion)))
                # Clic forzado
                self.driver.execute_script("arguments[0].click();", consignacion_element)
            except TimeoutException:
                # Fallback con script general si el elemento no se encuentra o es difícil
                print("Elemento 'Consignación a la Venta' no encontrado, usando script de fallback.")
                self.driver.execute_script("""
                    const links = document.querySelectorAll('a');
                    for (let link of links) {
                        if (link.textContent.trim() === 'Consignación a la Venta') {
                            link.click();
                            break;
                        }
                    }
                """)
            
            # Click en Reportes
            xpath_reportes = "//a[font[contains(text(), 'Reportes')]]"
            self.wait.until(EC.presence_of_element_located((By.XPATH, xpath_reportes))).click()

            # Seleccionar la opcion de "Nivel de Inventario" (Opción 0)
            select_name = '3.0.1.3.7.InventoryReportsGrouped.1.3.7'
            select_element = self.wait.until(EC.presence_of_element_located((By.NAME, select_name)))
            selector = Select(select_element)
            selector.select_by_index(0)

            # Click en el botón "Siguiente"
            self.wait.until(EC.presence_of_element_located((By.NAME, '3.0.1.3.7.InventoryReportsGrouped.1.3.9'))).click()
            
            # --------------------------------------------------
            # CICLO DE DESCARGA POR CIUDAD (Asegurando la continuidad)
            # --------------------------------------------------
            
            ciudades = {
                2: {'nombre': 'Queretaro', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.13.3', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.13.7'},
                3: {'nombre': 'Monterrey', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                4: {'nombre': 'Guadalajara', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                5: {'nombre': 'Hermosillo', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                6: {'nombre': 'Tijuana', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                7: {'nombre': 'Nogales', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                8: {'nombre': 'Monterrey_GC', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                9: {'nombre': 'Tultitlan', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'},
                10: {'nombre': 'Ciudad Juarez', 'btn_descarga': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.29', 'btn_regresar': '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.11.61'}
            }
            
            select_city_name = '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.19'
            btn_next_name = '3.0.1.3.7.InventoryReportsGrouped.1.5.0.1.9.43'

            for index, data in ciudades.items():
                ciudad = data['nombre']
                btn_descarga = data['btn_descarga']
                btn_regresar = data['btn_regresar']
                
                print(f"\n--- Procesando: {ciudad} (Índice {index}) ---")

                try:
                    # 1. Seleccionar la ciudad por índice
                    city_select_element = self.wait.until(EC.presence_of_element_located((By.NAME, select_city_name)))
                    city_selector = Select(city_select_element)
                    city_selector.select_by_index(index)
                    
                    # 2. Click en Siguiente
                    self.wait.until(EC.presence_of_element_located((By.NAME, btn_next_name))).click()
                    
                    # 3. Intentar descargar el XLS (Punto crítico)
                    try:
                        # Descargar XLS
                        self.wait.until(EC.presence_of_element_located((By.NAME, btn_descarga))).click()
                        time.sleep(3) # Espera breve para iniciar la descarga
                        
                        # Clic en regresar (asumiendo que el botón lleva de vuelta a la selección)
                        self.wait.until(EC.presence_of_element_located((By.NAME, btn_regresar))).click() 
                        print(f"Reporte de {ciudad} solicitado con éxito.")

                    except TimeoutException:
                        # Si falla la descarga o el regreso, el flujo CONTINÚA al siguiente ciclo (ciudad)
                        print(f"Aviso: No se encontró el botón de descarga/regreso para {ciudad}. Puede que no haya datos. Continuamos...")
                    except NoSuchElementException:
                        # Si el elemento no existe en la página
                        print(f"Aviso: El botón de descarga/regreso para {ciudad} no existe. Continuamos...")

                except TimeoutException as e:
                    # Si falla la selección de la ciudad o el botón "Siguiente", registra y CONTINÚA
                    print(f"Error fatal en la selección/navegación de {ciudad}. El flujo continua al siguiente ciclo.")
                except Exception as e:
                    print(f"Error desconocido al procesar {ciudad}: {e}. El flujo continua al siguiente ciclo.")


            # --------------------------------------------------
            # FIN DEL CICLO: Esperar la última descarga
            # --------------------------------------------------
            self.archivo_xls = self._wait_for_xls_download() 

        except TimeoutException as e:
            print(f"Error de tiempo de espera FATAL durante el inicio de la automatización: {e}")
        except Exception as e:
            print(f"Error inesperado FATAL durante la ejecución: {e}")
        finally:
            self.cerrar() # <--- Activado: Cierra el driver para liberar los archivos descargados.
            
        return self.archivo_xls


    def _wait_for_xls_download(self): 
        """Espera a que el archivo XLS se descargue en la carpeta temporal."""
        archivo_xls = None
        inicio = time.time()
        print(f"Esperando descarga en: {RUTA_DESCARGA}")
        while time.time() - inicio < 60:
            xls = glob.glob(os.path.join(RUTA_DESCARGA, "*.xls"))
            if xls:
                # Tomar el más reciente
                archivo_xls = max(xls, key=os.path.getctime)
                break
            time.sleep(1)
        return archivo_xls

    def cerrar(self):
        """Cierra el driver si está activo."""
        if hasattr(self, 'driver') and self.driver:
            print("Cerrando navegador.")
            self.driver.quit()


# ==============================
# BLOQUE DE PRUEBA INDEPENDIENTE
# ==============================
if __name__ == "__main__":
    
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