from PIL import Image
import pytesseract
import cv2
import numpy as np
import os
import unicodedata
import shutil
#################################################

def PrePross(folder,pros,borrar):

    # Configuración de tesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    custom_config = r'--oem 3 --psm 6'

    folder_path = "input/zborrar/"+borrar  
    files = [f for f in os.listdir(folder_path) if f.endswith('.jpg')]
    for file in files:
    # Comprobar si el archivo es una imagen jpg
             
        if file.endswith(".jpg"):
            # Ruta completa del archivo
            imagen = os.path.join('input','zborrar',borrar,file)
            img=Image.open(imagen) 
            seachtype= os.path.join('input',folder,'GIRO.txt')
            try:
                with open(seachtype, 'r') as f:
                    first_word = f.readline().split()[0]
            except FileNotFoundError:
                # Si el archivo no se encuentra, asigna una cadena vacía a first_word
                first_word = ""
            if "si" in first_word:
                if img.width > img.height:
                    img = img.rotate(90, expand=True)
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
            if "si" in first_word:
                corr_resized_img=resized_img.crop((1671, 116, 1782, 869))
                corr_resized_img=corr_resized_img.rotate(90, expand=True)
                text = pytesseract.image_to_string(corr_resized_img, config=custom_config, lang='spa')
                text = unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore').decode('utf-8')
                if 'mineria' not in text.lower():   
                    resized_img = resized_img. rotate(180, expand=True)
                # Guardar la imagen modificad
            resized_img.save('input'+'\\zborrar\\'+borrar+'\\'+pros+'\\'+file)
            print("PROCESADO: ", borrar,", ",file)


    #################################################

def doc(ruta,folder,numb,rutadoc,borrar):
    # abrir la imagen
    imagen = Image.open(ruta+'//'+folder+'//'+borrar+'//'+numb)  


    if not os.path.exists('output/' + rutadoc):
        os.makedirs('output/' + rutadoc)

    if rutadoc=='1. DECLARACION DE PRODUCCION':
        # especificar el área que deseas cortar en  ES IM 004 DECLARACION DE PRODUCCION
        dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini = ((1389, 184, 1564, 815), (1389, 1110, 1564, 1617), (1390, 1793, 1564, 2076), (1388, 2360, 1500, 2800), ( 818,325,910,600), (818, 960, 1038, 1294), (1218, 1993, 1300, 2432), (93, 2152, 218, 2824), (33, 821, 223, 1371), (39, 1369, 249, 1577))
        dimensiones= [dim_nomini, dim_apell_ini, dim_cc_ini, dim_act_ini, dim_cant_ini, dim_uni_ini, dim_mun_ini, dim_dia_ini, dim_firma_ini, dim_huella_ini]
        nombres = ['1. Nombre', '2. Apellido', '3. Cc', '8. Act', '4. Cantidad','5. Unidad','6. Municipio','7. Fecha', '9. Firma', '10. Huella']
    
        
    elif rutadoc=='2. RUT':
        # especificar el área que deseas cortar en  RUT
        dimQr, dimCc, dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2, dimName, LName, rut_fecha = ((676,270,895,523), (975,752,1214,789),(214,611,508,654),(33,1671,159,1715),(457,1675,580,1711),(1002,1667,1120,1710),(1133,1667,1245,1718),(815,933,964,977),(11,933,800,981),(1356,2554,1504,2597))
        dimensiones = [dimQr, dimCc,  dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2,dimName,LName, rut_fecha]
        nombres = ['9. CODIGO_QR', '3. Cc', '4. Nit', '5. Act_ppa', '6. Act_sria', '7. Otras_act','8. Otras_act1','1. Nombre','2. Apellido','10. Fecha rut']
       
    
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


def docfull(ruta,folder,rutadoc,borrar):
    if not os.path.exists('output/' + rutadoc):
        os.makedirs('output/' + rutadoc)
    nombres = 'full'
    if not os.path.exists(f"output/{rutadoc}/{nombres}"):
            os.makedirs(f"output/{rutadoc}/{nombres}")
    # Ruta de la carpeta de origen
    ruta_origen = "input/zborrar/"+borrar+'/'

    # Ruta de la carpeta de destino
    ruta_destino = os.path.join('output',rutadoc,nombres)

    # Obtener la lista de archivos en la carpeta de origen
    archivos = os.listdir(ruta_origen)

    # Iterar sobre cada archivo en la carpeta de origen
    for archivo in archivos:
        # Comprobar si el archivo es una imagen
        if archivo.endswith(".jpg") or archivo.endswith(".jpeg") or archivo.endswith(".png"):
            # Construir la ruta completa de origen y destino
            ruta_archivo_origen = os.path.join(ruta_origen, archivo)
            ruta_archivo_destino = os.path.join(ruta_destino, archivo)
            # Copiar el archivo de origen al destino
            shutil.copy2(ruta_archivo_origen, ruta_archivo_destino)