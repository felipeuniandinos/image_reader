import os
import Pdf2Png
import Imagen2Letra
import writeExcel
import pandas as pd
from datetime import datetime
import re

def ordenar_numeros(nombre_archivo): #funcion para obtener imagenes en orden
    # Extraer el número del nombre de archivo utilizando una expresión regular
    numero = re.search(r'\d+', nombre_archivo).group()
    # Convertir el número de cadena a entero
    return int(numero)

ruta_input = 'input'
ruta_output = 'output'
folder_img1 = 'borrar1' #ruta donde se van a alojar las imagenes del primer formato de pdf
folder_out1 = 'firma'
folder_out2 = 'huella'
now = datetime.now() # fecha de ejecución
format1 = str(now.strftime('%d_%m_%Y'))# nombre del archivo output para el primer pdf
xlsx_name1='output\\salida_'+format1+'.xlsx'# ruta del archivo output para el primer pdf
df1 = pd.DataFrame(columns=['NOMBRE_1','APELLIDO_1','CEDULA_1','CODIGO_ACT_ECONOMICA','CANTIDAD_DE_MINERAL_VENDIDA','UNIDAD_DE_MEDIDA','MUNICIPIO_ORIGEN','DIA','IMG'])#crear df con columnas primer pdf
df1.to_excel(xlsx_name1,index=False)#crear excel con columnas primer pdf


archivos_pdf = [f for f in os.listdir(ruta_input) if f.endswith('.pdf')] #listar archivos pdf

pdf_reader = Pdf2Png.Pdf2Png1(ruta_input,folder_img1,archivos_pdf[0]) #generar imagenes en carpeta ruta_input\\borrar1 

archivos_img = [f for f in os.listdir(ruta_input+'\\'+folder_img1) if f.endswith('.tiff')] #listar archivos img

archivos_img = sorted(archivos_img, key=ordenar_numeros) #ordenar los archivos

etiqueta_nom='nombre_' #se ubica la zona donde podremos encontrar el nombre
x_etiqueta_nom_ini=185
x_etiqueta_nom_fin=185+709
y_etiqueta_nom_ini=1454
y_etiqueta_nom_fin=1454+119

etiqueta_apell='apell_' #se ubica la zona donde podremos encontrar el apellido
x_etiqueta_apell_ini=1115
x_etiqueta_apell_fin=1115+540
y_etiqueta_apell_ini=1470
y_etiqueta_apell_fin=1470+109

etiqueta_cc='cc_' #se ubica la zona donde podremos encontrar CC
x_etiqueta_cc_ini=1800
x_etiqueta_cc_fin=1800+309
y_etiqueta_cc_ini=1479
y_etiqueta_cc_fin=1479+133

etiqueta_act='act_' #se ubica la zona donde podremos encontrar actividad economica
x_etiqueta_act_ini=2355
x_etiqueta_act_fin=2355+507
y_etiqueta_act_ini=1470
y_etiqueta_act_fin=1470+79

etiqueta_cant='cant_' #se ubica la zona donde podremos encontrar cantidad de mineral
x_etiqueta_cant_ini=335
x_etiqueta_cant_fin=335+227
y_etiqueta_cant_ini=965
y_etiqueta_cant_fin=965+67

etiqueta_uni='uni_' #se ubica la zona donde podremos encontrar unidad de mineral
x_etiqueta_uni_ini=945
x_etiqueta_uni_fin=945+360
y_etiqueta_uni_ini=969
y_etiqueta_uni_fin=969+69

etiqueta_mun='mun_' #se ubica la zona donde podremos encontrar municipio de mineral
x_etiqueta_mun_ini=1623
x_etiqueta_mun_fin=1623+475
y_etiqueta_mun_ini=971
y_etiqueta_mun_fin=971+73

etiqueta_dia='dia_' #se ubica la zona donde podremos encontrar diaicipio de mineral
x_etiqueta_dia_ini=2160
x_etiqueta_dia_fin=2160+687
y_etiqueta_dia_ini=60
y_etiqueta_dia_fin=60+157

etiqueta_firma='firma_' #se ubica la zona donde podremos encontrar firma
x_etiqueta_firma_ini=830
x_etiqueta_firma_fin=830+551
y_etiqueta_firma_ini=33
y_etiqueta_firma_fin=33+243

etiqueta_huella='huella_' #se ubica la zona donde podremos encontrar huella
x_etiqueta_huella_ini=1380
x_etiqueta_huella_fin=1380+195
y_etiqueta_huella_ini=29
y_etiqueta_huella_fin=29+235

for im in archivos_img:
    nom_str = 'error'
    apell_str = 'error'
    cc_str = 'error'
    act_str = 'error'
    cant_str = 'error'
    uni_str = 'error'
    mun_str = 'error'
    dia_str = 'error'
    try:
        img_proc=Imagen2Letra.PrePross(ruta_input,im,folder_img1)
        os.remove(ruta_input+'\\'+folder_img1+'\\'+im)#borrar imagen original

        img_nom=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_nom,x_etiqueta_nom_ini,x_etiqueta_nom_fin,y_etiqueta_nom_ini,y_etiqueta_nom_fin)
        nom_str=Imagen2Letra.Img2Str(ruta_input,img_nom,folder_img1)
        nom_str=nom_str.replace(' | ','')
        nom_str=nom_str.replace('\n','')
        nom_str=nom_str.replace('|','')
        nom_str=nom_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_nom)

        img_apell=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_apell,x_etiqueta_apell_ini,x_etiqueta_apell_fin,y_etiqueta_apell_ini,y_etiqueta_apell_fin)
        apell_str=Imagen2Letra.Img2Str(ruta_input,img_apell,folder_img1)
        apell_str=apell_str.replace(' | ','')
        apell_str=apell_str.replace('\n','')
        apell_str=apell_str.replace('|','')
        apell_str=apell_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_apell)

        img_cc=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_cc,x_etiqueta_cc_ini,x_etiqueta_cc_fin,y_etiqueta_cc_ini,y_etiqueta_cc_fin)
        cc_str=Imagen2Letra.Img2Str(ruta_input,img_cc,folder_img1)
        cc_str=cc_str.replace(' | ','')
        cc_str=cc_str.replace('\n','')
        cc_str=cc_str.replace('|','')
        cc_str=cc_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_cc)
        
        img_act=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_act,x_etiqueta_act_ini,x_etiqueta_act_fin,y_etiqueta_act_ini,y_etiqueta_act_fin)
        act_str=Imagen2Letra.Img2Str(ruta_input,img_act,folder_img1)
        act_str=act_str.replace(' | ','')
        act_str=act_str.replace('\n','')
        act_str=act_str.replace('|','')
        act_str=act_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_act)

        img_cant=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_cant,x_etiqueta_cant_ini,x_etiqueta_cant_fin,y_etiqueta_cant_ini,y_etiqueta_cant_fin)
        cant_str=Imagen2Letra.Img2Str(ruta_input,img_cant,folder_img1)
        cant_str=cant_str.replace(' | ','')
        cant_str=cant_str.replace('\n','')
        cant_str=cant_str.replace('|','')
        cant_str=cant_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_cant)

        img_uni=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_uni,x_etiqueta_uni_ini,x_etiqueta_uni_fin,y_etiqueta_uni_ini,y_etiqueta_uni_fin)
        uni_str=Imagen2Letra.Img2Str(ruta_input,img_uni,folder_img1)
        uni_str=uni_str.replace(' | ','')
        uni_str=uni_str.replace('\n','')
        uni_str=uni_str.replace('|','')
        uni_str=uni_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_uni)

        img_mun=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_mun,x_etiqueta_mun_ini,x_etiqueta_mun_fin,y_etiqueta_mun_ini,y_etiqueta_mun_fin)
        mun_str=Imagen2Letra.Img2Str(ruta_input,img_mun,folder_img1)
        mun_str=mun_str.replace(' | ','')
        mun_str=mun_str.replace('\n','')
        mun_str=mun_str.replace('|','')
        mun_str=mun_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_mun)

        img_dia=Imagen2Letra.ImgCut(ruta_input,img_proc,folder_img1,etiqueta_dia,x_etiqueta_dia_ini,x_etiqueta_dia_fin,y_etiqueta_dia_ini,y_etiqueta_dia_fin)
        dia_str=Imagen2Letra.Img2Str(ruta_input,img_dia,folder_img1)
        dia_str=dia_str.replace(' | ','')
        dia_str=dia_str.replace('\n','')
        dia_str=dia_str.replace('|','')
        dia_str=dia_str.strip()
        os.remove(ruta_input+'\\'+folder_img1+'\\'+img_dia)

        img_firma=Imagen2Letra.ImgCutFirm(im,cc_str,ruta_output,folder_out1,ruta_input,img_proc,folder_img1,etiqueta_firma,x_etiqueta_firma_ini,x_etiqueta_firma_fin,y_etiqueta_firma_ini,y_etiqueta_firma_fin)
        img_huella=Imagen2Letra.ImgCutFirm(im,cc_str,ruta_output,folder_out2,ruta_input,img_proc,folder_img1,etiqueta_huella,x_etiqueta_huella_ini,x_etiqueta_huella_fin,y_etiqueta_huella_ini,y_etiqueta_huella_fin)

        #os.remove(ruta_input+'\\'+folder_img1+'\\'+img_firma)
    except Exception as e:
        print('lectura fallida numero : ',e) 
    writeExcel.WrExcl(xlsx_name1,nom_str,apell_str,cc_str,act_str,cant_str,uni_str,mun_str,dia_str,str(im))
    os.remove(ruta_input+'\\'+folder_img1+'\\'+img_proc)


    




