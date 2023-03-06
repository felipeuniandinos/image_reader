import pytesseract
from PIL import Image
import pandas as pd
from datetime import date
import os

def write(ruta,folder,numb,contador):
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'    
        # especificar la ruta de la imagen
    imagen_path = os.path.join(ruta, folder, numb, contador)
    img = Image.open(imagen_path)

        # convertir la imagen a texto utilizando PyTesseract, se deja en global para cargar la variable texto en la funcion cargarExcel
    texto = pytesseract.image_to_string(img)
    print(texto)


