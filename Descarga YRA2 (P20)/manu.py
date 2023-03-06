import requests

url = "http://dataq-prod.int.net.nokia.com:7780/pls/apex/f?p=115:8:4397957325663150:PRINT_SIG:NO:"
filename = "archivo_descargado.slk"
response = requests.get(url)
with open(filename, "wb") as f:
    f.write(response.content)