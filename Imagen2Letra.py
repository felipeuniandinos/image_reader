from PIL import Image
import pytesseract
import cv2
import numpy as np
import os


#################################################3
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


def PrePross(input,pros,borrar):
    
    # Configuración de tesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    custom_config = r'--oem 3 --psm 6'

    folder_path = "input/zborrar/"+borrar  
    files = os.listdir(folder_path)
    for file in files:
    # Comprobar si el archivo es una imagen jpg
        if file.endswith(".jpg"):
            # Ruta completa del archivo
            imagen = os.path.join(input,'zborrar',borrar,file)
            img=Image.open(imagen) 
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
            resized_img.save(input+'\\zborrar\\'+borrar+'\\'+pros+'\\'+file)
            #CAMBIO ROTACIÓN en PROD
            if borrar == 'borrar6':
                imagen = os.path.join(input+'\\zborrar\\'+borrar+'\\'+pros+'\\'+file)
                img = Image.open(imagen)
                exif_info = img._getexif()
                # Obtener la información de Exif de la imagen
                exif_info = img._getexif()

                if exif_info:
                    # Obtener la etiqueta Exif Orientation (274)
                    exif_orientation = exif_info.get(274)

                    # Determinar la rotación requerida según la orientación actual
                    if img.width > img.height:
                        img = img.rotate(90, expand=True)
                    text = pytesseract.image_to_string(img.crop((760, 18, 818, 428)), config=custom_config)
                    if 'mineria' not in text.lower():
                        img = img.rotate(180, expand=True)
                    # Guardar la imagen modificada
                img.save(input+'\\zborrar\\'+borrar+'\\'+pros+'\\'+file)
            #OTROS
            else :

                
                print("Reprocesamiento ", borrar,file)

    #################################################

def doc(ruta,folder,numb,rutadoc,borrar):
    # abrir la imagen
    imagen = Image.open(ruta+'//'+folder+'//'+borrar+'//'+numb)   
    

    if rutadoc=='doc_eq':
        # especificar el área que deseas cortar en  ES IM 004 DOCUMENTO EQUIVALENTE
        dimQr, dimCc, dimnum, dimDate, dimGrBru, dimLey = ((1321,333,1549,561), (503,652,857,704), (1700,319,2221,409), (1683,499,2223,565), (807,891,1183,959), (827,1225,1189,1295))

        dimensiones = [dimQr, dimCc, dimnum, dimDate, dimGrBru, dimLey]
        nombres = ['CODIGO_QR', 'Cc', 'Num', 'Date', 'GrBru', 'Ley']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

        
    elif rutadoc=='doc_rut':
        # especificar el área que deseas cortar en  RUT
        dimQr, dimCc, dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 = ((993,457,1248,699), (1323,934,1687,970),(343,794,775,838),(169,1868,323,1924),(705,1872,859,1920),(1397,1880,1553,1928),(1571,1880,1719,1928))
        dimensiones = [dimQr, dimCc,  dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 ]
        nombres = ['CODIGO_QR', 'Cc', 'Nit', 'Act_ppa', 'Act_sria', 'Otras_act','Otras_act1']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_ced':
        # especificar el área que deseas cortar en  CEDULA
        ancho, alto = 5,5
        dimensiones = (327, 168, 532 + ancho, 198 + alto)
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775+ ancho,838+alto),(212,211,528+ancho,243+alto), (327, 168, 532 + ancho, 198 + alto),(203,432,674+ancho ,472+alto),(745,123,1091+ancho,512+alto),(187,1085,466+ancho,1397+alto),(521,1331,668+ancho,1357+alto))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_alc':
        # especificar el área que deseas cortar en  ALCALDIA
        dimfull = ((112,466,2569,2401))

        dimensiones = [dimfull]
        nombres = ['full']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_carta':
        # especificar el área que deseas cortar en  CARTA REPRESENTANTE LEGAL
        dimfull = ((112,466,2569,2401))

        dimensiones = [dimfull]
        nombres = ['full']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_decProd':

        # especificar el área que deseas cortar en  ES IM 004 DECLARACION DE PRODUCCION
        dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini = ((1467, 184, 1564, 815), (1462, 1110, 1585, 1617), (1469, 1793, 1585, 2076), (1471, 2360, 1548, 2800), (963, 328, 1037, 593), (969, 960, 1038, 1294), (1325, 1993, 1380, 2320), (2160, 95, 205, 2824), (33, 821, 223, 1371), (39, 1369, 249, 1577))
        dimensiones= [dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Act', 'Cantidad','Unidad','Municipio','Fecha', 'Firma', 'Huella']
        
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_vin':
        # especificar el área que deseas cortar en  ALCALDIA
        dimfull = ((112,466,2569,2401))

        dimensiones = [dimfull]
        nombres = ['full']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_sisben':
        # especificar el área que deseas cortar en  SISBEN
        dimfull = ((112,466,2569,2401))

        dimensiones = [dimfull]
        nombres = ['full']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_trdatos':
        # especificar el área que deseas cortar en  ALCALDIA
        dimfull = ((112,466,2569,2401))

        dimensiones = [dimfull]
        nombres = ['full']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")