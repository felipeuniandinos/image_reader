from PIL import Image, ImageEnhance
import pytesseract
from PIL import Image
import pandas as pd
import datetime
import os
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
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
    texto = texto.replace("~", "").replace(",","").replace("|", "").replace("?","").replace("(", "").replace(")", "").replace(";-","").replace('"', "").replace("]", "").replace("[", "").replace("5{", "").replace(";}", "").replace(";-", "").replace("-", "")


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
        # # Inicializar la variable libro_trabajo
        # libro_trabajo = None
        # contador = contador.split(".")[0]
        # # Verificar si el archivo Excel ya existe
        # if os.path.exists(archivo_excel_path):
        #     # Si el archivo Excel ya existe, cargar el libro de trabajo existente
        #     libro_trabajo = load_workbook(archivo_excel_path)

        #     # Seleccionar la hoja activa del libro de trabajo
        #     hoja_activa = libro_trabajo.active

        #     # Buscar la fila correspondiente al contador
        #     for fila in hoja_activa.iter_rows(min_row=2, max_col=1, values_only=True):
        #         if fila[0] == int(contador):
        #             # Si se encuentra el contador, agregar el texto a la celda de la columna correspondiente
        #             columna = chr(ord('A') + len(hoja_activa[1]))
        #             hoja_activa[f"{columna}{int(fila[0])}"] = texto
        #             break
        #     else:
        #         # Si no se encuentra el contador, agregar una nueva fila con el contador y el texto
        #         nueva_fila = [int(contador), texto]
        #         hoja_activa.append(nueva_fila)

        # else:
        #     # Si el archivo Excel no existe, crear un nuevo dataframe con el contador y el texto
        #     df = pd.DataFrame({"Contador": [int(contador)], numb: [texto]})

        #     # Guardar el dataframe en un nuevo archivo Excel
        #     df.to_excel(archivo_excel_path, index=False)

        #     # Cargar el libro de trabajo nuevo
        #     libro_trabajo = load_workbook(archivo_excel_path)

        # # Guardar los cambios en el archivo Excel y cerrar el libro de trabajo
        # libro_trabajo.save(archivo_excel_path)
        # libro_trabajo.close()

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
    for file in excel_files:
        # Obtener el número de la columna a partir del nombre del archivo Excel
        numb = file.split('_')[0]

        # Cargar el archivo Excel como un dataframe y almacenarlo en el diccionario
        df = pd.read_excel(os.path.join(ruta, folder, file), index_col=0)
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

def lectortxt(rutaout, folder, excel):
    # Crea un nuevo archivo Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    fecha_actual = datetime.datetime.now().strftime("%d_%m_%Y")
    # Define la ruta de la carpeta que contiene los archivos .txt
    carpeta = os.path.join(rutaout, folder)
    hoja = "Sheet1"
    columna_c = "C"
    columna_a = "A"
    columna_b = "B"
    columna_h = "H"
    columna_e = "E"

    # Abre el archivo de Excel y lee los valores de las columnas especificadas
    wb = openpyxl.load_workbook(excel)
    ws = wb[hoja]
    valores_buscar_c = [cell.value for cell in ws[columna_c]]
    valores_buscar_a = [cell.value for cell in ws[columna_a]]
    valores_buscar_b = [cell.value for cell in ws[columna_b]]
    valores_buscar_e = [cell.value for cell in ws[columna_e]]
    valores_buscar_h = [cell.value for cell in ws[columna_h]]

    # Define una lista con los nombres de las columnas
    columnas = ['A', 'B', 'C', 'E', 'H']

    for archivo in os.listdir(carpeta):
        if archivo.endswith(".txt"):
            # Abre el archivo de texto
            with open(os.path.join(carpeta, archivo), "r") as f:
                # Lee las líneas del archivo
                lineas = f.readlines()

                # Procesa cada línea en orden
                for linea in lineas:
                    # Elimina los separadores de miles del contenido
                    linea = linea.replace(",", "").replace(".", "").replace("(", "").replace(")", "").replace(";-",
                                                                                                                   "").replace(
                        '"', "").replace("]", "").replace("[", "").replace("5{", "").replace(";}", "").replace(";-", "")

                    # Busca los valores buscados en la línea
                    for i in range(len(valores_buscar_c)):
                        valor_c = valores_buscar_c[i]
                        if valor_c and str(valor_c) in linea:
                            # Busca los valores de las columnas A, B, E y H en la misma línea
                            valores_encontrados = []
                            for col in columnas:
                                if col == 'A':
                                    valor_encontrado = valores_buscar_a[i]
                                elif col == 'B':
                                    valor_encontrado = valores_buscar_b[i]
                                elif col == 'E':
                                    valor_encontrado = valores_buscar_e[i]
                                elif col == 'H':
                                    valor_encontrado = valores_buscar_h[i]
                                else:
                                    valor_encontrado = None
                                if valor_encontrado:
                                    valores_encontrados.append(valor_encontrado)
                                else:
                                    valores_encontrados.append('x')
                            # Si encuentra los valores, escribe en el archivo Excel
                            sheet.append([archivo] + [valor_c] + valores_encontrados)

                            # Remueve los valores de las columnas A, B, E y H para evitar duplicados
                            valores_buscar_a[i] = None
                            valores_buscar_b[i] = None
                            valores_buscar_e[i]= None
                            valores_buscar_h[i] = None
                    else:
                        continue
    # Guarda el archivo Excel
    workbook.save(("output/salida/"+folder+"_resultado.xlsx"))
    print("se ha guardado lo leido de:", folder)


def unir_contador_columnas(nombre_archivo):
    # Leer archivo Excel
    df = pd.read_excel(nombre_archivo)

    # Detectar todas las columnas con "contador"
    contador_cols = [col for col in df.columns if "contador" in col.lower()]

    # Unir las columnas con "contador" en la columna 1
    df["columna_1"] = df[contador_cols].apply(lambda x: '_'.join(x.dropna().astype(str)), axis=1)
    df.drop(contador_cols, axis=1, inplace=True)

    # Guardar el archivo actualizado como un nuevo archivo Excel
    nuevo_nombre_archivo = nombre_archivo.split('.')[0] + "_actualizado.xlsx"
    df.to_excel(nuevo_nombre_archivo, index=False)