from Pdf2Png import Pdf2Png1,ordenar_numeros
import multiprocessing
from Imagen2Letra import doc, PrePross
from writeExcel import write, merge_excel_files,lectortxt,unir_contador_columnas
import os
from datetime import datetime
import time                

start_time = time.time()


# ########################################}
now = datetime.now() # fecha de ejecución
fecha_actual = str(now.strftime('%d_%m_%Y'))# nombre del archivo output para el primer pdf
xlsx_name1='output\\salida\\doc_decProd_'+fecha_actual+'.xlsx'# ruta del archivo output para el primer pdf


#Declaramos las varibles con el nombre de las rutas a utilizar, empezamos con las entradas
ruta_input='input'
in_docEq, in_rut,  in_ced, in_alc, in_carta, in_decPro, in_forVin, in_sisben, in_traDat = 'DOCUMENTO EQUIVALENTE', 'RUT', 'CEDULA','ALCALDIA', 'CARTA REPRESENTANTE LEGAL',  'DECLARACION DE PRODUCCION', 'FORMATO DE VINCULACION', 'SISBEN', 'TRATAMIENTO DE DATOS'

carpetas = [in_docEq,in_rut,in_ced,in_alc,in_carta,in_decPro,in_forVin,in_sisben,in_traDat]
#continuamos con los 'borrar' que tendran temporalmente las imagenes
folder_img1='zborrar'
fol_1, fol_2, fol_3, fol_4, fol_5, fol_6, fol_7, fol_8, fol_9= "borrar1", "borrar2", "borrar3", "borrar4", "borrar5", "borrar6", "borrar7", "borrar8", "borrar9"
folders = [fol_1, fol_2, fol_3, fol_4, fol_5, fol_6, fol_7, fol_8, fol_9]
pross='Procesado'
procesado='borrar2/'+pross
procesado1='borrar6/'+pross
procesado2='borrar3/'+pross
folpros=[procesado,procesado1,procesado1]

#Rutas de salida
ruta_output= 'output'
salida= 'salida'
folder_eq, eq_cc, eq_num, eq_date, eq_grBru, eq_ley = 'doc_eq', 'Cc', 'Num', 'Date', 'GrBru', 'Ley'
folder_rut, rut_cc, rut_nit, rut_ppal, rut_sria, rut_oth,rut_oth1 = 'doc_rut', 'Cc', 'Nit', 'Act_ppa', 'Act_sria', 'Otras_act','Otras_act1'
folder_ced, ced_cc, ced_num, ced_date, ced_grBru, ced_ley = 'doc_ced', 'Cc', 'Num', 'Date', 'GrBru', 'Ley'
folder_decProd, decProd_name, decProd_lname, decProd_cc, decProd_firma, decProd_act,decProd_huel,decProd_cant,decProd_uni,decProd_mun,decProd_date = ('doc_decProd','Nombre', 'Apellido', 'Cc', 'Firma', 'Act', 'Huella','Cantidad','Unidad','Municipio','Fecha')
folder_alc, alc_full = 'doc_alc', 'full'
folder_carta, carta_full = 'doc_carta', 'full'
folder_vin, vin_full = 'doc_vin', 'full'
folder_trdatos, trdatos_full = 'doc_trdatos', 'full'
folder_sisben, sisben_full = 'doc_sisben', 'full'


#recorre los documentos de imagenes recortadas y los lee.  
if __name__ == '__main__':
    with multiprocessing.Pool(processes=10) as pool :    

        # pdf_reader = Pdf2Png1(ruta_input, folder_img1, archivos_pdf[0], carpeta, folder)
        p0 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_docEq, fol_1))
        p1 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_rut, fol_2))
        p2 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_ced, fol_3))            
        p3 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_alc, fol_4))
        p4 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_carta, fol_5))
        p5 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_decPro, fol_6))
        p6 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_forVin, fol_7))  
        p7 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_traDat, fol_8))
        p8 = pool.apply_async(Pdf2Png1, args=(ruta_input, folder_img1, in_sisben, fol_9))
        p0.get()
        p1.get()
        p2.get()
        p3.get()
        p4.get()
        p5.get()
        p6.get()
        p7.get()
        p8.get()

        PrePross(ruta_input,pross,fol_2)
        PrePross(ruta_input,pross,fol_3)
        PrePross(ruta_input,pross,fol_6)          
        for folder in folders:
            archivos_img = [f for f in os.listdir(ruta_input + '\\' + folder_img1 + '\\' + folder) if f.endswith('.jpg')]
            archivos_img = sorted(archivos_img, key=ordenar_numeros)  
            # archivos_imgpros = [f for f in os.listdir(ruta_input + '\\' + folder_img1 + '\\' +folder +'\\'+fol) if f.endswith('.jpg')]
            # archivos_imgpros = sorted(archivos_imgpros, key=ordenar_numeros)  
            
        for archivo in archivos_img:                           
            p11 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_rut, procesado))
            p15 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_decProd, procesado1))
            p10 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_eq, fol_1))
            p12 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_ced, fol_3))            
            p13 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_alc, fol_4))
            p14 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_carta, fol_5))
            p16 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_vin, fol_7))  
            p17 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_trdatos, fol_8))
            p18 = pool.apply_async(doc, args=(ruta_input, folder_img1, archivo, folder_sisben, fol_9))
            #    #DECLARACION DE PRODUCCIÓN    
            p211 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_name, archivo))
            p212 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_lname, archivo))
            p213 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_cc, archivo))
            p214 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_firma, archivo))
            p215 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_act, archivo))
            p216 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_huel, archivo))
            p217 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_cant, archivo))
            p218 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_uni, archivo))
            p219 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_mun, archivo))
            p220 = pool.apply_async(write, args=(ruta_output, folder_decProd, decProd_date, archivo))

            #EQ 
            p20 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_cc, archivo))
            p21 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_num, archivo))
            p22 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_date, archivo))
            p23 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_grBru, archivo))
            p24 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_ley, archivo))
            #RUT
            p25 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_cc, archivo))
            p26 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_nit, archivo))
            p27 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_ppal, archivo))
            p28 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_sria, archivo))
            p29 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth, archivo))
            p210 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth1, archivo))
         
            
                
            #ALCALDIA
            p30 = pool.apply_async(write, args=(ruta_output, folder_alc, alc_full, archivo))
            #CARTA
            p31 = pool.apply_async(write, args=(ruta_output, folder_carta, carta_full, archivo))
            #VIN
            p32 = pool.apply_async(write, args=(ruta_output, folder_vin, vin_full, archivo))
            #TRDATOS
            p33 = pool.apply_async(write, args=(ruta_output, folder_trdatos, trdatos_full, archivo))
            #SISBEN
            p34 = pool.apply_async(write, args=(ruta_output, folder_sisben, sisben_full, archivo))
            
            #MEZCLA LOS EXCEL 
            p35= pool.apply_async(merge_excel_files, args=(ruta_output, folder_eq, salida))
            p36= pool.apply_async(merge_excel_files, args=(ruta_output, folder_rut, salida))
            p37= pool.apply_async(merge_excel_files, args=(ruta_output, folder_decProd, salida))

        # Wait for all processes to finish before starting a new iteration
        p10.wait()
        p11.wait()
        #p12.wait()
        p13.wait()
        p14.wait()
        p15.wait()
        p20.wait()
        p21.wait()
        p22.wait()
        p23.wait()
        p24.wait()
        p25.wait()
        p26.wait()
        p27.wait()
        p28.wait()
        p29.wait()
        p210.wait()
        p211.wait()
        p212.wait()
        p213.wait()
        p214.wait()
        p215.wait()
        p216.wait()
        p217.wait()
        p218.wait()
        p219.wait()
        p220.wait()
        p30.wait()
        p31.wait()
        p32.wait()
        p33.wait()
        p34.wait()
        p35.wait()
        p36.wait()
    lectortxt(ruta_output, 'doc_carta',xlsx_name1)
    lectortxt(ruta_output, 'doc_alc', xlsx_name1)
    lectortxt(ruta_output, 'doc_vin', xlsx_name1)
    lectortxt(ruta_output, 'doc_trdatos',xlsx_name1)
    lectortxt(ruta_output, 'doc_sisben',xlsx_name1)
    # Merge excel files
    merge_excel_files(ruta_output, salida,salida)
    unir_contador_columnas('salida/salida_'+fecha_actual+'.xlsx')

    for dir_name in [folder_rut, folder_eq]:
        dir_path = os.path.join(ruta_output, dir_name)
        for file in os.listdir(dir_path):
            if file.endswith('.xlsx'):
                file_path = os.path.join(dir_path, file)
                os.remove(file_path)

    for dir_name in [fol_1, fol_2,fol_3, fol_4,fol_5, fol_6,fol_7, fol_8, fol_9]:
        dir_path = os.path.join(ruta_input,folder_img1, dir_name)
        for file in os.listdir(dir_path):
            if file.endswith('.jpg'):
                file_path = os.path.join(dir_path, file)
                os.remove(file_path)


#FIN DEL CONTADOR DE EJECUCIÓN DEL PROGRAMA
    end_time = time.time()
    total_time = end_time - start_time
    print("Tiempo total de ejecución: {:.2f} segundos".format(total_time))


