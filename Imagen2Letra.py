import numpy as np
from PIL import Image
import cv2
import pytesseract

def Img2Str(ruta_input,name,folder_img):
    #se debe instalar tesseract-ocr-w64-setup-5.3.0.20221222, es un instalador
    #ubicado en este folder. luego se podrá ejecutar la siguiente linea de codigo.
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
    # Cargar imagen
    img = cv2.imread(ruta_input+'\\'+folder_img+'\\'+name)

    # Aplicar OCR
    texto = pytesseract.image_to_string(img)

    # Imprimir resultado
    return(texto)

def ImgCut(ruta_input,name,folder_img,etiqueta,x_etiqueta_ini,x_etiqueta_fin,y_etiqueta_ini,y_etiqueta_fin):
    img = cv2.imread(ruta_input+'\\'+folder_img+'\\'+name)
    crop_img = img[x_etiqueta_ini:x_etiqueta_fin,y_etiqueta_ini:y_etiqueta_fin]
    cv2.imwrite(ruta_input+'\\'+folder_img+'\\'+etiqueta+name, crop_img)
    return(etiqueta+name)

def ImgCutFirm(im,cc_str,ruta_output,folder_out1,ruta_input,name,folder_img,etiqueta,x_etiqueta_ini,x_etiqueta_fin,y_etiqueta_ini,y_etiqueta_fin):
    img = cv2.imread(ruta_input+'\\'+folder_img+'\\'+name)
    crop_img = img[x_etiqueta_ini:x_etiqueta_fin,y_etiqueta_ini:y_etiqueta_fin]
    cv2.imwrite(ruta_output+'\\'+folder_out1+'\\'+im[:-5]+'_'+cc_str+'.png', crop_img)
    img = Image.open(ruta_output+'\\'+folder_out1+'\\'+im[:-5]+'_'+cc_str+'.png')
    return(im[:-5]+'_'+cc_str+'.png')

def PrePross(ruta_input,name,folder_img):
    # Cargar la imagen y convertirla a escala de grises
    img = Image.open(ruta_input+'\\'+folder_img+'\\'+name)
    gray = img.convert('L')

    # Aplicar un umbral a la imagen
    threshold = 127
    binary = gray.point(lambda p: p > threshold and 255)
    binary = np.array(binary)

    # Encontrar los contornos de la imagen binaria
    contours = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)[0]

    # Encontrar el contorno con el área máxima
    max_area = 0
    max_contour = None
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > max_area:
            max_area = area
            max_contour = contour

    second_max_area = 0
    second_max_contour = None
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > second_max_area and area < max_area:
            second_max_area = area
            second_max_contour = contour

    # Obtener las coordenadas del rectángulo que encierra el contorno más grande
    x, y, w, h = cv2.boundingRect(second_max_contour)


    # Recortar la imagen original con las coordenadas del rectángulo
    cropped_img = img.crop((x, y, x+w, y+h))
    # nuevo tamaño a guardar
    width = 1817
    height = 2889
    resized_img = cropped_img.resize((width, height))
    # Guardar la nueva imagen
    resized_img.save(ruta_input+'\\'+folder_img+'\\proc_'+name)
    return('proc_'+name)