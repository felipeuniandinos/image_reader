from PyPDF2 import PdfReader
import re
import os
from pdf2image import convert_from_path
from PyPDF2 import PdfReader

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
def Pdf2Png1(ruta_input, folder_img, doc, borrar):
    folder_path = os.path.join(ruta_input,doc)

    # Obtener la lista de archivos en la carpeta
    files = os.listdir(folder_path)

    # Filtrar solo los archivos con extensión .pdf
    pdf_files = [f for f in files if f.endswith('.pdf')]
    for pdf_file in pdf_files:
        pdf_path = os.path.join(ruta_input, doc, pdf_file)
        # Continúa con el resto del código para convertir el archivo PDF a PNG
        with open(pdf_path, 'rb') as f:
            pdf = PdfReader(f)
            for page_num in range(len(pdf.pages)):
                page = pdf.pages[page_num]
                images = convert_from_path(pdf_path, 500, first_page=page_num+1, last_page=page_num+1, poppler_path=r'C:\Program Files\poppler-23.01.0\Library\bin')
                gray_image = images[0].convert('L') # convierte la imagen a escala de grises
                gray_image.save(f'{ruta_input}/{folder_img}/{borrar}/{page_num+1}.jpg', 'JPEG')



