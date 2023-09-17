from Pdf2Png import Pdf2Png1,ordenar_numeros, pausa
import multiprocessing
from Imagen2Letra import doc, docfull,PrePross
from writeExcel import write, merge_excel_files,lector,mergemaster,definir,unir_contador_columnas,unificarCEDULA
from lectoraws import AWS
import os
from datetime import datetime
import time             
from PIL import Image   


start_time = time.time()

# crear la carpeta input/zborrar/ si no existe
ruta_carpeta = "input/zborrar/"
if not os.path.exists(ruta_carpeta):
    os.makedirs(ruta_carpeta)
# abrir archivo de texto
with open('input/DOCUMENTOS.txt', 'r') as f:
    # leer todo el contenido del archivo
    contenido = f.read()
    # dividir el contenido en líneas utilizando el carácter de nueva línea como separador
    lineas = contenido.split('\n')
    # crear una lista de folios
    carpetas = []
    for i, linea in enumerate(lineas):
        # crear una variable de folio para cada línea
        # asignar el valor de la línea a la variable de folio correspondiente
        # agregar la variable de folio a la lista de folios
        carpetas.append(linea)
        # crear una carpeta llamada borrarX en la carpeta input/zborrar/
        ruta_carpeta_folio = os.path.join(ruta_carpeta, f"borrar{i+1}")
        if not os.path.exists(ruta_carpeta_folio):
            os.mkdir(ruta_carpeta_folio)
# verificar si hay carpetas adicionales y eliminarlas si es necesario
num_folios = len(carpetas)
num_carpetas = len(os.listdir(ruta_carpeta))
if num_carpetas > num_folios:
    carpetas_a_eliminar = num_carpetas - num_folios
    for i in range(num_folios+1, num_carpetas+1):
        ruta_carpeta_folio = os.path.join(ruta_carpeta, f"borrar{i}")
        os.rmdir(ruta_carpeta_folio)
folder = []
for i in range(1, num_folios+1):
    ruta_carpeta_folio = os.path.join(ruta_carpeta, f"borrar{i}")
    if not os.path.exists(ruta_carpeta_folio):
        os.makedirs(ruta_carpeta_folio)
    folder.append(os.path.basename(ruta_carpeta_folio))
# ########################################}
now = datetime.now() # fecha de ejecución
fecha_actual = str(now.strftime('%d_%m_%Y'))# nombre del archivo output para el primer pdf
xlsx_name1='output\\salida\\doc_decProd_'+fecha_actual+'.xlsx'# ruta del archivo output para el primer pdf
#Declaramos las varibles con el nombre de las rutas a utilizar, empezamos con las entradas
ruta_input='input'
#continuamos con los 'borrar' que tendran temporalmente las imagenes
folder_img1='zborrar'
pross='Procesado'
procesado='borrar1/'+pross
procesado1='borrar2/'+pross
procesado2='borrar3/'+pross
folpros=[procesado,procesado1]

#Rutas de salida
ruta_output= 'output'
salida= 'salida'
folder_sisben, sisben_name, sisben_lname, sisben_cc,sisben_date = ('2.1 SISBEN COOR', '1. Nombre', '2. Apellido', '3. Cc', '7. Fecha')
folder_rut, rut_Qr,rut_cc, rut_nit, rut_ppal, rut_sria, rut_oth,rut_oth1,rut_name,rut_lname,rut_fecha= ('2. RUT', '9. CODIGO_QR', '3. Cc', '4. Nit', '5. Act_ppa', '6. Act_sria', '7. Otras_act','8. Otras_act1','1. Nombre','2. Apellido','10. Fecha rut')
folder_decProd, decProd_name, decProd_lname, decProd_cc, decProd_firma, decProd_act,decProd_huel,decProd_cant,decProd_uni,decProd_mun,decProd_date = ('1. DECLARACION DE PRODUCCION','1. Nombre', '2. Apellido', '3. Cc', '9. Firma', '8. Act', '10. Huella','4. Cantidad','5. Unidad','6. Municipio','7. Fecha')


#recorre los documentos de imagenes recortadas y los lee.  
if __name__ == '__main__':
    with multiprocessing.Pool(processes=15) as pool :    

        # pdf_reader = Pdf2Png1(ruta_input, folder_img1, archivos_pdf[0], carpeta, folder)
        for borrar, carpeta in zip(folder,carpetas):
            p00 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, carpeta, borrar))
        p00.wait()   
        #FIJOS
        # p01= pool.apply_async(PrePross, args=(folder_decProd    ,pross,'borrar1'))
        # p01.wait()
        archivos_img = [f for f in os.listdir(ruta_input + '\\' + folder_img1 + '\\' + borrar) if f.endswith('.jpg')]
        archivos_img = sorted(archivos_img, key=ordenar_numeros)
        p02=pool.apply_async(PrePross, args=(ruta_input,pross,'borrar2'))
        p02.wait()
            
        # Obtener una nueva lista de carpetas comenzando en la posición 3
        sub_borrar = folder[2:]
        sub_carpeta= carpetas[2:]
        print(archivos_img)
        for archivo in archivos_img: 
            print(archivos_img)
            for carpeta_out,borrarito in zip (sub_carpeta,sub_borrar):    
                p05 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, carpeta_out, borrarito))
            p05.get()
            p03 = pool.apply_async(AWS, args=(archivo,))
            p04 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_rut, procesado1))
    
            p03.wait()
            p04.wait()
            
            
            #RUT
            p19 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_cc, archivo))
            p110 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_nit, archivo))
            p111 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_ppal, archivo))
            p112 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_sria, archivo))
            p113 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth, archivo))
            p114 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth1, archivo))
            p115 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_name, archivo))
            p116 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_lname, archivo))
            p117 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_fecha, archivo))
            
            # p005 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_sisben, 'borrar3'))
            for carpeta_out,borrarito in zip (sub_carpeta,sub_borrar):    
                p10 = pool.apply_async(write, args=(ruta_output, carpeta_out, 'full', archivo))
                
            # Esperar a que todas las tareas de write se completen antes de continuar
            
            p10.wait()
            p19.wait()
            p110.wait()
            p111.wait()
            p112.wait()
            p113.wait()
            p114.wait()
            p115.wait()
            p116.wait()
            p117.wait()
        
            

        pausa()
        #Lector
        for titulo in sub_carpeta:
            p40=pool.apply_async(lector,args=('1. DECLARACION DE PRODUCCION_'+fecha_actual+'.xlsx', titulo))
            p40.get()
        
        p402=pool.apply_async(unificarCEDULA,args=('1. DECLARACION DE PRODUCCION', '3. CEDULA'))
        p402.get()
        #MEZCLA LOS EXCEL RUT Y DECPROD   
        p20= pool.apply_async(merge_excel_files, args=(ruta_output, folder_rut, salida))
        
        p20.get()

        p400=pool.apply(definir,args=(folder_decProd,folder_rut, '1. Nombre'))
        p401=pool.apply(definir,args=(folder_decProd,folder_rut, '2. Apellido'))
        p4011=pool.apply(definir,args=(folder_decProd,folder_rut, '3. Cc'))
        #MEZCLA LOS EXC RESTANTES            
        p50=pool.apply(mergemaster,args=(ruta_output, salida))
        #p51=pool.apply(mergemaster,args=(ruta_output, 'errores'))
        p52 = pool.apply(unir_contador_columnas, args=(('output\\salida' + fecha_actual + '.xlsx',)))


    for dir_name in carpetas :
        try:
            dir_path = os.path.join(ruta_output, dir_name)
            for file in os.listdir(dir_path):
                if file.endswith('.xlsx'):
                    file_path = os.path.join(dir_path, file)
                    os.remove(file_path)
        except: 
            print("no encontro documento a eliminar: ", dir_path)

    for dir_name in folder :
        dir_path = os.path.join(ruta_input,folder_img1, dir_name)
        for file in os.listdir(dir_path):
            if file.endswith('.jpg'):
                file_path = os.path.join(dir_path, file)
                os.remove(file_path)


    #FIN DEL CONTADOR DE EJECUCIÓN DEL PROGRAMA
    end_time = time.time()
    total_time = end_time - start_time
    print("Tiempo total de ejecución: {:.2f} segundos".format(total_time))


