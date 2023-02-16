from PyPDF2 import PdfReader

# Convierte cada página a una imagen
def Pdf2Png1(ruta_input,folder_img,archivos_pdf):

    reader = PdfReader(ruta_input+'\\'+archivos_pdf)

    # Extrae las páginas como imágenes y las guarda en una lista
    for page in reader.pages:
        for image in page.images:
            numb=str(image.name[2:])
            with open(ruta_input+'\\'+folder_img+'\\'+numb, "wb") as fp:
                fp.write(image.data)