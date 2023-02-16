import pandas as pd

def WrExcl(xlsx_name1,nombre,apell_str,CC,act_str,cant_str,uni_str,mun_str,dia_str,im):
    df2 = pd.DataFrame({'NOMBRE_1':[nombre],'APELLIDO_1':[apell_str],'CEDULA_1':[CC],'CODIGO_ACT_ECONOMICA':[act_str],'CANTIDAD_DE_MINERAL_VENDIDA':[cant_str],'UNIDAD_DE_MEDIDA':[uni_str],'MUNICIPIO_ORIGEN':[mun_str],'DIA':[dia_str],'IMG':[im]})
    writer = pd.ExcelWriter(xlsx_name1, engine='openpyxl', mode='a', if_sheet_exists='overlay')
    df2.to_excel(writer,index=False ,sheet_name='Sheet1', startrow=writer.sheets['Sheet1'].max_row, header=None)
    writer.save()