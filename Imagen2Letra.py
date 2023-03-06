from PIL import Image


def doc(ruta,folder,numb,rutadoc,borrar):
    # abrir la imagen
    imagen = Image.open(ruta+'//'+folder+'//'+borrar+'//'+numb)

    if rutadoc=='doc_eq':
        # especificar el área que deseas cortar en  ES IM 004 DOCUMENTO EQUIVALENTE
        dimQr, dimCc, dimnum, dimDate, dimGrBru, dimLey = ((1321,333,1549,561), (503,652,857,704), (1700,319,2221,409), (1683,499,2223,565), (807,891,1183,959), (827,1225,1189,1295))

        dimensiones = [dimQr, dimCc, dimnum, dimDate, dimGrBru, dimLey]
        nombres = ['CODIGO_QR', 'Cc', 'Num', 'Date', 'GrBru', 'Ley']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

        
    elif rutadoc=='doc_rut':
        # especificar el área que deseas cortar en  RUT
        dimQr, dimCc, dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 = ((993,457,1248,699), (1363,934,1715,976),(343,794,775,838),(149,1870,327,1926),(705,1872,859,1920),(1397,1880,1553,1928),(1571,1880,1719,1928))

        dimensiones = [dimQr, dimCc,  dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 ]
        nombres = ['CODIGO_QR', 'Cc', 'Nit', 'Act_ppa', 'Act_sria', 'Otras_act','Otras_act1']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

        # especificar el área que deseas cortar en  RUT"""


    
    
