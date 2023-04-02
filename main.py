from Pdf2Png import Pdf2Png1,ordenar_numeros
import multiprocessing
from Imagen2Letra import doc, PrePross
from writeExcel import write, merge_excel_files,lector,mergemaster,definir
import os
from datetime import datetime
import time                

start_time = time.time()


# ########################################}
now = datetime.now() # fecha de ejecución
fecha_actual = str(now.strftime('%Y_%m_%d'))# nombre del archivo output para el primer pdf
xlsx_name1='output\\salida\\doc_decProd_'+fecha_actual+'.xlsx'# ruta del archivo output para el primer pdf
compname='1. Nombre_'+fecha_actual+'.xlsx'
complname='2. Apellido_'+fecha_actual+'.xlsx'
compCc='3. Cc_'+fecha_actual+'.xlsx'
compcant='4. Cantidad_'+fecha_actual+'.xlsx'
compmuni='6. Municipio_'+fecha_actual+'.xlsx'
compdate='7. Fecha_'+fecha_actual+'.xlsx'
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
folder_eq, eq_full = 'doc_eq', 'full'
folder_rut, rut_Qr,rut_cc, rut_nit, rut_ppal, rut_sria, rut_oth,rut_oth1,rut_name,rut_lname = 'doc_rut', '9. CODIGO_QR', '3. Cc', '4. Nit', '5. Act_ppa', '6. Act_sria', '7. Otras_act','8. Otras_act1','1. Nombre','2. Apellido'
folder_ced, ced_full = 'doc_ced', 'full'
folder_decProd, decProd_name, decProd_lname, decProd_cc, decProd_firma, decProd_act,decProd_huel,decProd_cant,decProd_uni,decProd_mun,decProd_date = ('doc_decProd','1. Nombre', '2. Apellido', '3. Cc', '9. Firma', '8. Act', '10. Huella','4. Cantidad','5. Unidad','6. Municipio','7. Fecha')
folder_alc, alc_full = 'doc_alc', 'full'
folder_carta, carta_full = 'doc_carta', 'full'
folder_vin, vin_full = 'doc_vin', 'full'
folder_trdatos, trdatos_full = 'doc_trdatos', 'full'
folder_sisben, sisben_full = 'doc_sisben', 'full'


#recorre los documentos de imagenes recortadas y los lee.  
if __name__ == '__main__':
    with multiprocessing.Pool(processes=8) as pool :    

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

           
            #RUT
            p25 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_cc, archivo))
            p26 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_nit, archivo))
            p27 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_ppal, archivo))
            p28 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_sria, archivo))
            p29 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth, archivo))
            p221 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_oth1, archivo))
            p222 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_name, archivo))
            p223 = pool.apply_async(write, args=(ruta_output, folder_rut, rut_lname, archivo))
            
         
            
                
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
            #CEDULA
            p35 = pool.apply_async(write, args=(ruta_output, folder_ced, ced_full, archivo))
            #EQ 
            p20 = pool.apply_async(write, args=(ruta_output, folder_eq, eq_full, archivo))
            
            #MEZCLA LOS EXCEL RUT Y DECPROD 
            
            p37= pool.apply_async(merge_excel_files, args=(ruta_output, folder_rut, salida))
            p38= pool.apply_async(merge_excel_files, args=(ruta_output, folder_decProd, salida))

        # Wait for all processes to finish before starting a new iteration
        p10.wait()
        p11.wait()
        p12.wait()
        p13.wait()
        p14.wait()
        p15.wait()
        p20.wait()
        p25.wait()
        p26.wait()
        p27.wait()
        p28.wait()
        p29.wait()
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
        p221.wait()
        p222.wait()
        p223.wait()
        p30.wait()
        p31.wait()
        p32.wait()
        p33.wait()
        p34.wait()
        p35.wait()

        p37.wait()
        p38.wait()
        p400=pool.apply(definir,args=(folder_decProd,folder_rut, '1. Nombre'))
        p401=pool.apply(definir,args=(folder_decProd,folder_rut, '2. Apellido'))
        
        #Lectura ALCALDIA  
        p40=pool.apply_async(lector,args=('1. Nombre',compname, folder_alc))
        p41=pool.apply_async(lector,args=('2. Apellido', complname, folder_alc))    
        p42=pool.apply_async(lector,args=('3. Cc', compCc, folder_alc))
        p43=pool.apply_async(lector,args=('4. Cantidad', compcant, folder_alc))
        p44=pool.apply_async(lector,args=('6. Municipio', compmuni, folder_alc))
        p45=pool.apply_async(lector,args=('7. Fecha', compdate, folder_alc))

        #Lectura CEDULA  
        p46=pool.apply_async(lector,args=('1. Nombre',compname, folder_ced))
        p47=pool.apply_async(lector,args=('2. Apellido', complname, folder_ced))    
        p48=pool.apply_async(lector,args=('3. Cc', compCc, folder_ced))
        p49=pool.apply_async(lector,args=('4. Cantidad', compcant, folder_ced))
        p410=pool.apply_async(lector,args=('6. Municipio', compmuni, folder_ced))
        p411=pool.apply_async(lector,args=('7. Fecha', compdate, folder_ced))

        #Lectura DOCUMENTO EQUIVALENTE  
        p412=pool.apply_async(lector,args=('1. Nombre',compname, folder_eq))
        p413=pool.apply_async(lector,args=('2. Apellido',complname, folder_eq))    
        p414=pool.apply_async(lector,args=('3. Cc',compCc, folder_eq))
        p415=pool.apply_async(lector,args=('4. Cantidad',compcant, folder_eq))
        p416=pool.apply_async(lector,args=('6. Municipio',compmuni, folder_eq))
        p417=pool.apply_async(lector,args=('7. Fecha',compdate, folder_eq))
        
        #Lectura SISBEN  
        p418=pool.apply_async(lector,args=('1. Nombre',compname, folder_sisben))
        p419=pool.apply_async(lector,args=('2. Apellido',complname, folder_sisben))    
        p420=pool.apply_async(lector,args=('3. Cc',compCc, folder_sisben))
        p421=pool.apply_async(lector,args=('4. Cantidad',compcant, folder_sisben))
        p422=pool.apply_async(lector,args=('6. Municipio',compmuni, folder_sisben))
        p423=pool.apply_async(lector,args=('7. Fecha',compdate, folder_sisben))
        
        #Lectura TRATAMIENTO DE DATOS  
        p424=pool.apply_async(lector,args=('1. Nombre',compname, folder_trdatos))
        p425=pool.apply_async(lector,args=('2. Apellido',complname, folder_trdatos))    
        p426=pool.apply_async(lector,args=('3. Cc',compCc, folder_trdatos))
        p427=pool.apply_async(lector,args=('4. Cantidad',compcant, folder_trdatos))
        p428=pool.apply_async(lector,args=('6. Municipio',compmuni, folder_trdatos))
        p429=pool.apply_async(lector,args=('7. Fecha',compdate, folder_trdatos))

        #Lectura VINCULACIÓN  
        p430=pool.apply_async(lector,args=('1. Nombre',compname, folder_vin))
        p431=pool.apply_async(lector,args=('2. Apellido',complname, folder_vin))    
        p432=pool.apply_async(lector,args=('3. Cc',compCc, folder_vin))
        p433=pool.apply_async(lector,args=('4. Cantidad',compcant, folder_vin))
        p434=pool.apply_async(lector,args=('6. Municipio',compmuni, folder_vin))
        p435=pool.apply_async(lector,args=('7. Fecha',compdate, folder_vin))
        p40.wait()
        p41.wait()
        p42.wait()
        p43.wait()
        p44.wait()
        p45.wait()
        p46.wait()
        p47.wait()
        p48.wait()
        p49.wait()
        p410.wait()
        p411.wait()
        p412.wait()
        p413.wait()
        p414.wait()
        p415.wait()
        p416.wait()
        p417.wait()
        p418.wait()
        p419.wait()
        p420.wait()
        p421.wait()
        p422.wait()
        p423.wait()
        p424.wait()
        p425.wait()
        p426.wait()
        p427.wait()
        p428.wait()
        p429.wait()
        p430.wait()
        p431.wait()
        p432.wait()
        p433.wait()
        p434.wait()
        p435.wait()
     # Merge excel files
        #MEZCLA LOS EXCRESTANTES            
        p39= pool.apply(merge_excel_files, args=(ruta_output, folder_alc, salida))
        p310= pool.apply(merge_excel_files, args=(ruta_output, folder_carta, salida))
        p311= pool.apply(merge_excel_files, args=(ruta_output, folder_ced, salida))
        p312= pool.apply(merge_excel_files, args=(ruta_output, folder_eq, salida))
        p313= pool.apply(merge_excel_files, args=(ruta_output, folder_sisben, salida))
        p314= pool.apply(merge_excel_files, args=(ruta_output, folder_trdatos, salida))
        p315= pool.apply(merge_excel_files, args=(ruta_output, folder_vin, salida))
        p50=pool.apply(mergemaster,args=(ruta_output, salida))
        p51=pool.apply(mergemaster,args=(ruta_output, 'errores'))
    # unir_contador_columnas('output/salida/salida_'+fecha_actual+'.xlsx')

    for dir_name in [folder_rut,folder_decProd,folder_alc,folder_eq,folder_carta,folder_sisben,folder_trdatos,folder_vin,folder_ced]:
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


