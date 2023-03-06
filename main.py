from Pdf2Png import ordenar_numeros, Pdf2Png1
import multiprocessing
from Imagen2Letra import doc
from writeExcel import write
import os
import pandas as pd
from datetime import datetime
import re
import time



start_time = time.time()

#Declaramos las varibles con el nombre de las rutas a utilizar, empezamos con las entradas
ruta_input='input'
in_docEq, in_rut, in_alc, in_carta, in_ced, in_crtPro, in_crtAdq, in_claHum, in_decPro, in_forVin, in_sisben, in_traDat = 'DOCUMENTO EQUIVALENTE', 'RUT', 'ALCALDIA', 'CARTA REPRESENTANTE LEGAL', 'CEDULA', 'CERTIFICACION A PROVEEDORES', 'CERTIFICADO DE ADQUISICION', 'CLAUSULA DE DERECHOS HUMANOS', 'DECLARACION DE PRODUCCION', 'FORMATO DE VINCULACION', 'SISBEN', 'TRATAMIENTO DE DATOS'
carpetas = [in_docEq,in_rut, in_alc, in_carta, in_ced, in_crtPro, in_crtAdq, in_claHum, in_decPro, in_forVin, in_sisben, in_traDat]
#continuamos con los 'borrar' que tendran temporalmente las imagenes
folder_img1='zborrar'
fol_1, fol_2, fol_3, fol_4, fol_5, fol_6, fol_7, fol_8, fol_9, fol_10, fol_11, fol_12 = "borrar1", "borrar2", "borrar3", "borrar4", "borrar5", "borrar6", "borrar7", "borrar8", "borrar9", "borrar10", "borrar11", "borrar12"
folders = [fol_1, fol_2, fol_3, fol_4, fol_5, fol_6, fol_7, fol_8, fol_9, fol_10, fol_11, fol_12]
#Rutas de salida
ruta_output= 'output'
folder_eq, eq_cc, eq_num, eq_date, eq_grBru, eq_ley = ('doc_eq', 'Cc', 'Num', 'Date', 'GrBru', 'Ley')
folder_rut, rut_cc, rut_num, rut_date, rut_grBru, rut_ley = ('doc_rut', 'Cc', 'Num', 'Date', 'GrBru', 'Ley')


# Llamar a la función ordenar_numeros
for carpeta, folder in zip(carpetas, folders):
    ordenaNum=('input/'+carpeta)
    archivos_pdf = [f for f in os.listdir(ordenaNum) if f.endswith('.pdf')]
    pdf_reader = Pdf2Png1(ruta_input, folder_img1, archivos_pdf[0], carpeta, folder)
    archivos_img = [f for f in os.listdir(ruta_input+'\\'+folder_img1+'\\'+folder) if f.endswith('.tiff')]
    archivos_img = sorted(archivos_img, key=ordenar_numeros)
    # rest of the code here

#recorre los documentos de imagenes recortadas y los lee.
if __name__ == '__main__':
    pool = multiprocessing.Pool(processes=7)
    for i in range(len(archivos_img)):
        
        # Se hace el llamado de la función write cambiando el segundo y tercer parametro de entrada
        p1 = pool.apply_async(doc, args=(ruta_input,folder_img1,archivos_img[i],folder_eq,fol_1))
        p2 = pool.apply_async(doc, args=(ruta_input,folder_img1,archivos_img[i],folder_rut,fol_2))
        p1.get()
        p2.get()
        p3 = pool.apply_async(write, args=(ruta_output,folder_eq,eq_cc,archivos_img[i]))
        p4 = pool.apply_async(write, args=(ruta_output,folder_eq,eq_num,archivos_img[i]))
        p5 = pool.apply_async(write, args=(ruta_output,folder_eq,eq_date,archivos_img[i]))
        p6 = pool.apply_async(write, args=(ruta_output,folder_eq,eq_grBru,archivos_img[i]))
        p7 = pool.apply_async(write, args=(ruta_output,folder_eq,eq_ley,archivos_img[i]))
        p1.get()
        p2.get()
        p3.get()
        p4.get()
        p5.get()
        p6.get()
        p7.get()    
    pool.close()
    pool.join()



#FIN DEL CONTADOR DE EJECUCIÓN DEL PROGRAMA
end_time = time.time()
total_time = end_time - start_time
print("Tiempo total de ejecución: {:.2f} segundos".format(total_time))

       
