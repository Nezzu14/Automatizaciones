from selenium import webdriver
from selenium.webdriver.common.by import By
import shutil    
import os
import time

def Descargar_GIC():
    
    # Definir la ruta del archivo original y el directorio de destino
    ruta_archivo = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Download')+"\\F&C GIC - SIG PC List.slk"

    directorio_destino = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + "\\YRA2"
   
    # Inicializa el navegador Edge con las opciones configuradas
    edge_path="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

    #driver = webdriver.Edge(edge_path)
    #driver = webdriver.Edge()
    driver = webdriver.Edge(executable_path=r'C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\msedge.exe')

    #Interacciones con Edge
    driver.get("http://dataq-prod.int.net.nokia.com:7780/")
    driver.set_window_size(1296, 696)
    driver.switch_to.frame(0)
    driver.find_element(By.CSS_SELECTOR, "td:nth-child(2) tr:nth-child(3) span").click()
    driver.find_element(By.LINK_TEXT, "Export Excel").click()
    
    #wait for the download
    time.sleep(30)
    
    # close the browser
    driver.quit()

    # Cortar el archivo al directorio de destino
    print(ruta_archivo)
    shutil.move(ruta_archivo, directorio_destino)

Descargar_GIC()

