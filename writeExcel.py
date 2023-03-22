import pytesseract
from PIL import Image
import pandas as pd
import datetime
import os
from openpyxl import load_workbook

def write(ruta, folder, numb, contador):
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'    
        # especificar la ruta de la imagen
    imagen_path = os.path.join(ruta, folder, numb, contador)
    img = Image.open(imagen_path)

    # convertir la imagen a texto utilizando PyTesseract, se deja en global para cargar la variable texto en la funcion cargarExcel
    texto = pytesseract.image_to_string(img)

    # Obtener la fecha actual y formatearla como YYYY-MM-DD
    
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")

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


def merge_excel_files(ruta, folder):
    """
    Combina los archivos Excel de una carpeta en uno solo.

    Parámetros:
    ruta (str): la ruta donde se encuentran los archivos Excel a combinar.
    folder (str): la carpeta donde se encuentran los archivos Excel a combinar.
    fecha_actual (str): la fecha actual en formato "YYYY-MM-DD".

    Retorna:
    None.
    """
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d")
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
    combined_df.to_excel(os.path.join(ruta, f"{fecha_actual}.xlsx"), index=True)

    # Imprimir mensaje de confirmación
    print(f"Los archivos Excel han sido combinados en {fecha_actual}.xlsx")
