from PIL import Image, ImageEnhance
import pytesseract
from PIL import Image
import pandas as pd
import datetime
import os
import openpyxl
from openpyxl import load_workbook
import re


pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'



def write(ruta, folder, numb, contador):
    # especificar la ruta de la imagen
    imagen_path = os.path.join(ruta, folder, numb, contador)
    img = Image.open(imagen_path)
    contador_numerico = ''.join(filter(str.isdigit, contador))


    # mejorar la calidad de la imagen
    img = img.convert('L')
    img = ImageEnhance.Sharpness(img).enhance(2)
    img = ImageEnhance.Contrast(img).enhance(2)

    # convertir la imagen a texto utilizando PyTesseract, se deja en global para cargar la variable texto en la funcion cargarExcel
    texto = pytesseract.image_to_string(img) 

    # Eliminar los puntos del texto
    texto = (texto.replace("~", "").replace(",","").replace("|", "").replace("?","").replace("(", "").replace(")", "").replace(";-","").replace('"', "").replace("]", "").replace("[", "").replace("5{", "").replace(";}", "").replace(";-", "").replace("-", "").replace(" /","").replace(",","")
    .replace("  ","").replace(" . “A/S \ ","").replace("fo","").replace("/","").replace("    ","").replace("&","").replace("N\ ","").replace("‘","").replace("\\",""))
    texto = re.sub(r"[^a-zA-Z0-9\s]++", "", texto)

    # Obtener la fecha actual y formatearla como YYYY-MM-DD
    fecha_actual = datetime.datetime.now().strftime("%Y_%m_%d")

    # Combinar la ruta, folder y numb para obtener la ruta completa del archivo Excel
    if numb == 'full':
        # Si numb es 'full', guardar el texto en un archivo .txt
        archivo_txt_path = os.path.join(ruta, folder,f"{contador_numerico}.txt")
        with open(archivo_txt_path, 'w') as archivo_txt:
            archivo_txt.write(texto)
    else:

        # Combinar la ruta, folder y numb para obtener la ruta completa del archivo Excel
        
        archivo_excel_path = os.path.join(ruta, folder,f"{numb}_{fecha_actual}.xlsx")

        # Inicializar la variable libro_trabajo
        libro_trabajo = None
        contador = contador.split(".")[0]
        # Verificar si el archivo Excel ya existe
        if os.path.exists(archivo_excel_path):
            # Si el archivo Excel ya existe, cargar el libro de trabajo existente
            libro_trabajo = load_workbook(archivo_excel_path)

            # Seleccionar la hoja activa del libro de trabajo
            hoja_activa = libro_trabajo.active

            # Buscar la fila correspondiente al contador
            for fila in hoja_activa.iter_rows(min_row=2, max_col=1, values_only=True):
                if fila[0] == int(contador):
                    # Si se encuentra el contador, agregar el texto a la celda de la columna correspondiente
                    columna = chr(ord('A') + len(hoja_activa[1]))
                    print(fila[0],columna,len(hoja_activa[1]))
                    hoja_activa[f"{columna}{int(fila[0])}"] = texto
                    break
            else:
                # Si no se encuentra el contador, agregar una nueva fila con el contador y el texto
                nueva_fila = [int(contador), texto]
                hoja_activa.append(nueva_fila)

        else:
            # Si el archivo Excel no existe, crear un nuevo dataframe con el contador y el texto
            df = pd.DataFrame({"Contador": [int(contador)], numb: [texto]})

            # Guardar el dataframe en un nuevo archivo Excel
            df.to_excel(archivo_excel_path, index=False)

            # Cargar el libro de trabajo nuevo
            libro_trabajo = load_workbook(archivo_excel_path)

        # Guardar los cambios en el archivo Excel y cerrar el libro de trabajo
        libro_trabajo.save(archivo_excel_path)
        libro_trabajo.close()

        # Imprimir mensaje de confirmación
        print(f"Los datos han sido agregados al archivo Excel {numb}_{fecha_actual}.xlsx en la fila correspondiente al contador {contador}")
        

        # Imprimir mensaje de confirmación



def merge_excel_files(ruta, folder, folder1):
    
    if folder1 == "":
        excel_files_out=os.path.join(ruta)
    else:
        excel_files_out=os.path.join(ruta,folder1)
    
    fecha_actual = datetime.datetime.now().strftime("%d_%m_%Y")
    # Obtener la lista de archivos Excel en la carpeta especificada
    excel_files = [f for f in os.listdir(os.path.join(ruta, folder)) if f.endswith('.xlsx')]

    # Verificar que hay al menos dos archivos Excel para combinar
    if len(excel_files) < 2:
        print("No hay suficientes archivos Excel para combinar.")
        return

    # Inicializar un diccionario para almacenar los dataframes de cada archivo Excel
    dfs = {}

    # Iterar sobre los archivos Excel y cargar cada uno como un dataframe en el diccionario
    # Iterar sobre los archivos Excel y cargar cada uno como un dataframe en el diccionario
    for file in excel_files:
        # Obtener el número de la columna a partir del nombre del archivo Excel
        numb = file.split('_')[0]

        # Cargar el archivo Excel como un dataframe y almacenarlo en el diccionario
        df = pd.read_excel(os.path.join(ruta, folder, file), index_col=0)
        df.reset_index(inplace=True)
        dfs[numb] = df


    # Concatenar los dataframes en uno solo con la función pd.concat()
    # El parámetro axis=1 indica que los dataframes se deben concatenar por columnas
    combined_df = pd.concat(dfs.values(), axis=1)

    # Crear un nuevo archivo Excel con el dataframe combinado
    combined_df.to_excel(os.path.join(excel_files_out, f"{folder}_{fecha_actual}.xlsx"), index=True)

    # Imprimir mensaje de confirmación
    print(f"archivos {combined_df}.xlsx")




def WrExcl(xlsx_name1,nombre,apell_str,CC,act_str,cant_str,uni_str,mun_str,dia_str,im):
    df2 = pd.DataFrame({'NOMBRE_1':[nombre],'APELLIDO_1':[apell_str],'CEDULA_1':[CC],'CODIGO_ACT_ECONOMICA':[act_str],'CANTIDAD_DE_MINERAL_VENDIDA':[cant_str],'UNIDAD_DE_MEDIDA':[uni_str],'MUNICIPIO_ORIGEN':[mun_str],'DIA':[dia_str],'IMG':[im]})
    writer = pd.ExcelWriter(xlsx_name1, engine='openpyxl', mode='a', if_sheet_exists='overlay')
    df2.to_excel(writer,index=False ,sheet_name='Sheet1', startrow=writer.sheets['Sheet1'].max_row, header=None)
    writer.save()
    writer.close()

def lector(columna,excel_or,rutatxt):
    rutadir= os.path.join('output',rutatxt)
    rutaerror= os.path.join('output','errores')
    excel_err= 'errores_'+rutatxt+excel_or
    columnar= columna+'_'+rutatxt
    df = pd.read_excel('output\\doc_decProd\\'+excel_or)
    df[columna] = df[columna].str.strip() #quitar espacios y saltos de linea
    if columna == '7. Fecha':
        df['7. Fecha'] = df['7. Fecha'].str.split('\n').str[-1]
    #listar txt.
    List_txt=[f for f in os.listdir(rutadir) if f.endswith('.txt')] 
    dfout = df.copy(0)
    result=[]
    for t in df[columna].to_list():
        comp = 'x'
        for txt in List_txt:
            with open(rutadir+'\\'+txt,'r') as file:
                if columna == '7. Fecha':
                    lineas = file.readlines()
                    for linea in lineas:
                        if str(t).lower() in str(linea).lower():
                            comp = linea
                else:
                    if str(t).lower() in str(file.read()).lower():
                        comp = t
        result.append(comp)
        file.close

    df[columna] = result

    df_salida = pd.DataFrame({columnar: result})
    df_er=df_salida[df_salida[columnar]=='x']
    df_er.to_excel(rutaerror+'\\'+excel_err, index=False)
    df_salida.to_excel(rutadir+'\\'+excel_or, index=False)


def mergemaster(output,salida):
 
    # Directorio donde se encuentran los archivos Excel
    directorio = os.path.join(output,salida)
    fecha_actual = datetime.datetime.now().strftime("%d_%m_%Y")

    # Lista para almacenar los nombres de los archivos Excel
    archivos_excel = []

    # Recorre todos los archivos del directorio y selecciona solo los que tienen extensión .xlsx
    for archivo in os.listdir(directorio):
        if archivo.endswith(".xlsx"):
            archivos_excel.append(archivo)

    # Concatena las columnas de todos los archivos Excel en un solo DataFrame
    df_concatenado = pd.DataFrame()
    for archivo in archivos_excel:
        df_excel = pd.read_excel(os.path.join(directorio, archivo))
        df_concatenado = pd.concat([df_concatenado, df_excel], axis=1)

    # Guarda el DataFrame concatenado en un nuevo archivo Excel
    df_concatenado.to_excel(output+"/"+salida+fecha_actual+".xlsx", index=False)




def definir(enrutador,enrutador2,colum):
    fecha_actual = datetime.datetime.now().strftime("%d_%m_%Y")
    xlsx_name1='output\\salida\\'+enrutador+'_'+fecha_actual+'.xlsx'# ruta del archivo output para el primer pdf
    xlsx_name2='output\\salida\\'+enrutador2+'_'+fecha_actual+'.xlsx'# ruta del archivo output para el primer pdf

    # Leer los archivos Excel
    df1 = pd.read_excel(xlsx_name1)
    df2 = pd.read_excel(xlsx_name2)

    # Rellenar valores faltantes en df1 con valores correspondientes de df2
    df1[colum] = df1[colum].fillna(df2[colum])

    # Rellenar valores faltantes en df2 con valores correspondientes de df1
    df2[colum] = df2[colum].fillna(df1[colum])

    # Escribir los archivos actualizados
    df1.to_excel(xlsx_name1, index=False)
    df2.to_excel(xlsx_name2, index=False)