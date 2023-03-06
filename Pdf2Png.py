from PyPDF2 import PdfReader
import re
from PIL import Image
import os

def ordenar_numeros(nombre_archivo): #funcion para obtener imagenes en orden
    # Extraer el número del nombre de archivo utilizando una expresión regular
    numero = re.search(r'\d+', nombre_archivo)
    if numero is not None:
        # Convertir el número de cadena a entero
        return int(numero.group())
    else:
        # Si no se encuentra un número, devolver 0 para que se ordene al principio
        return 0

# Convierte cada página a una imagen
def Pdf2Png1(ruta_input, folder_img, archivos_pdf, doc, borrar):
    reader = PdfReader(os.path.join(ruta_input, doc, archivos_pdf))
    for page in reader.pages:
        for image in page.images:
            numb = str(image.name[2:])
            filename = os.path.join(ruta_input, folder_img, borrar, numb)
            with open(filename, "wb") as fp:
                fp.write(image.data)
            # Verificar si la imagen es un archivo JPG y cambiar la extensión a TIFF
            if os.path.splitext(filename)[1] == ".jpg":
                new_filename = os.path.splitext(filename)[0] + ".tiff"
                # Verificar si el archivo ya existe antes de intentar renombrar
                if not os.path.exists(new_filename):
                    os.rename(filename, new_filename)
                