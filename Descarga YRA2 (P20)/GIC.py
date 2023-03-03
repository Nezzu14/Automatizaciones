#PRIMERO EN CORRER DE LOS ARCHIVOS GIC

from selenium import webdriver
from selenium.webdriver.common.by import By
#import shutil    
#import os
import time
from GIC_Path import Archivo_GIC
import tkinter as tk

def Descargar_GIC():   

    #    Inicializa el navegador Edge con las opciones configuradas
    #edge_driver_path="C:\Program Files (x86)\Microsoft\Edge\Application\msedgedriver.exe"
    #driver = webdriver.Edge(executable_path=edge_driver_path)

    options = webdriver.EdgeOptions()
    options.add_argument('--no-sandbox')
    driver = webdriver.Edge(options=options)

    #driver = webdriver.Edge()

    # Direccionamiento a la Pagina de GIC y descarga el archivo de GIC
    driver.get('http://dataq-prod.int.net.nokia.com:7780/')
    driver.set_window_size(1296, 696)
    driver.switch_to.frame(0)
    driver.find_element(By.CSS_SELECTOR, "td:nth-child(2) tr:nth-child(3) span").click()
    driver.find_element(By.LINK_TEXT, "Export Excel").click()
    
    # wait for the download
    #time.sleep(30)
    time.sleep(2)

    #   Ejecutar buscar el Path del Archivo del GIC (Class)
    root = tk.Tk()
    root.iconbitmap(r"C:\Users\migumart\OneDrive - Nokia\Archivos personales\Automatizacion Python\Descarga YRA2 (P20)\nokia.ico")
    input_form = Archivo_GIC(root)
    root.mainloop()

    # close the browser
    driver.quit()

#  Ejecutar Descarga_GIC junto a Buscar Archivo del GIC
if __name__ == '__main__':
    Descargar_GIC()
    
