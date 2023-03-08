from selenium.webdriver.common.by import By
import time
from GIC_Path import Archivo_GIC
import tkinter as tk
from msedge.selenium_tools import Edge, EdgeOptions, EdgeService

def Descargar_GIC():
    # Inicializar la sesión de Edge
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument("user-data-dir:\\Users\\migumart\\AppData\\Local\\Microsoft\\Edge\\User Data")
    service = EdgeService(executable_path="C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedgedriver.exe")
    driver = Edge(options=options, service=service)

    # Direccionamiento a la Pagina de GIC y descarga el archivo de GIC
    driver.get('http://dataq-prod.int.net.nokia.com:7780/')
    driver.set_window_size(1296, 696)
    driver.switch_to.frame(0)
    driver.find_element(By.CSS_SELECTOR, "td:nth-child(2) tr:nth-child(3) span").click()
    driver.find_element(By.LINK_TEXT, "Export Excel").click()
    
    # Esperar la descarga
    time.sleep(2)

    # Ejecutar buscar el Path del Archivo del GIC (Class)
    root = tk.Tk()
    root.iconbitmap(r"C:\\Users\\migumart\\OneDrive - Nokia\\Archivos personales\\Automatizacion Python\\Descarga YRA2 (P20)\\nokia.ico")
    input_form = Archivo_GIC(root)
    root.mainloop()

    # Cerrar la sesión de Edge
    driver.quit()

# Ejecutar Descargar_GIC junto a Buscar Archivo del GIC
if __name__ == '__main__':
    Descargar_GIC()
