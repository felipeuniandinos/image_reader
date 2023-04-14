from PIL import Image
import pytesseract
import cv2
import numpy as np
import os

#################################################
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
    
    # Convertir la imagen en escala de grises
    imagen = imagen.convert('L')

    # Binarizar la imagen
    threshold = 200 # Este es el valor umbral. Los valores mayores o iguales a este se convierten en negro y los valores menores se convierten en blanco
    imagen = imagen.point(lambda x: 0 if x < threshold else 255, '1')


    if not os.path.exists('output/' + rutadoc):
        os.makedirs('output/' + rutadoc)

    if rutadoc=='1. DECLARACION DE PRODUCCION':
        # especificar el área que deseas cortar en  ES IM 004 DECLARACION DE PRODUCCION
        dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini = ((1389, 184, 1564, 815), (1389, 1110, 1564, 1617), (1390, 1793, 1564, 2076), (1388, 2360, 1500, 2800), ( 818,325,910,600), (818, 960, 1038, 1294), (1218, 1993, 1300, 2432), (93, 2152, 218, 2824), (33, 821, 223, 1371), (39, 1369, 249, 1577))
        dimensiones= [dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini]
        nombres = ['1. Nombre', '2. Apellido', '3. Cc', '8. Act', '4. Cantidad','5. Unidad','6. Municipio','7. Fecha', '9. Firma', '10. Huella']
    
        
    elif rutadoc=='2. RUT':
        # especificar el área que deseas cortar en  RUT
        dimQr, dimCc, dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2, dimName, LName = ((676,270,895,523), (975,752,1214,789),(214,611,508,654),(33,1671,159,1715),(457,1675,580,1711),(1002,1667,1120,1710),(1133,1667,1245,1718),(815,933,964,977),(11,933,800,981))
        dimensiones = [dimQr, dimCc,  dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2,dimName,LName]
        nombres = ['9. CODIGO_QR', '3. Cc', '4. Nit', '5. Act_ppa', '6. Act_sria', '7. Otras_act','8. Otras_act1','1. Nombre','2. Apellido']
        
    elif rutadoc=='2.1 SISBEN COOR':
        # especificar el área que deseas cortar en  SISBEN
        sisben_name, sisben_lname, sisben_cc,sisbe_date = ((2191,1497,2735,1549), (2139,1565,2791,1633),(2199,1419,2763,1475),(2579,2247,3067,2371))
        dimensiones = [sisben_name, sisben_lname, sisben_cc,sisbe_date]
        nombres = ['1. Nombre', '2. Apellido', '3. Cc', '7. Fecha']
        
    
    else:
        # especificar el área que deseas cortar en  CEDULA
        dimfull = ((112,466,4200,5000))
        dimensiones = [dimfull]
        nombres = ['full']
    
    
    if len(nombres) < len(dimensiones):
        nombres += [f"subdir_{i+1}" for i in range(len(dimensiones)-len(nombres))]
    for i, dim in enumerate(dimensiones):
        if not os.path.exists(f"output/{rutadoc}/{nombres[i]}"):
            os.makedirs(f"output/{rutadoc}/{nombres[i]}")
        imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    print("Paso a salida recorte: ", borrar,rutadoc, nombres, numb)
