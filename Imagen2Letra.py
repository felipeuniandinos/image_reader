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
        dimQr, dimCc, dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 = ((993,457,1248,699), (1323,934,1687,970),(343,794,775,838),(169,1868,323,1924),(705,1872,859,1920),(1397,1880,1553,1928),(1571,1880,1719,1928))
        dimensiones = [dimQr, dimCc,  dimNit, dimActpal, dimActsria, dimOtact1, dimOtact2 ]
        nombres = ['CODIGO_QR', 'Cc', 'Nit', 'Act_ppa', 'Act_sria', 'Otras_act','Otras_act1']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_ced':
        # especificar el área que deseas cortar en  CEDULA
        ancho, alto = 5,5
        coordenadas = (327, 168, 532 + ancho, 198 + alto)
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775+ ancho,838+alto),(212,211,528+ancho,243+alto), (327, 168, 532 + ancho, 198 + alto),(203,432,674+ancho ,472+alto),(745,123,1091+ancho,512+alto),(187,1085,466+ancho,1397+alto),(521,1331,668+ancho,1357+alto))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_alc':
        # especificar el área que deseas cortar en  ALCALDIA
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_carta':
        # especificar el área que deseas cortar en  CARTA REPRESENTANTE LEGAL
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_decProd':
        # especificar el área que deseas cortar en  ES IM 004 DECLARACION DE PRODUCCION
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_forVin':
        # especificar el área que deseas cortar en ES IM 004 FORMATO DE VINCULACION
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")


    elif rutadoc=='doc_sisben':
        # especificar el área que deseas cortar en  SISBEN
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")

    elif rutadoc=='doc_tratDatos':
        # especificar el área que deseas cortar en  TRATAMIENTO DE DATOS
        dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp = ((343,794,775,838),(212,211,528,243), (1363,934,1715,976),(203,432,674 ,472),(745,123,1091,512),(187,1085,466,1397),(521,1331,668,1357))
        

        dimensiones = [dimName,dimLastname, dimCc,  dimFirma, dimFoto, dimHuel, dimDateexp]
        nombres = ['Nombre', 'Apellido', 'Cc', 'Firma', 'Foto', 'Huella','Fecha_exp']
        for i, dim in enumerate(dimensiones):
            imagen.crop(dim).save(f"output/{rutadoc}/{nombres[i]}/{numb}")


    
    
