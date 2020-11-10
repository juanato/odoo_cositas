import urllib3
import base64
import requests
import time 
import re
import pandas as oExcel
from pandas import ExcelWriter
from pandas import ExcelFile
from urllib.request import urlopen
import urllib.request
import math
import os
import csv

  



'''

Odoo 11 CE
Codifica una hoja Excel para importar el producto en odoo 11
Usamos la Excel para leer, pero por motivos de incompatibilidad, 
troceamos en múltiples CSV para importar poco a poco los productos
en Odoo, con la imagen sustituida desde la url a base64

Desde odoo 12, no es necesario ya que lo hace direcatmente durante la importación
Si Odoo 12 detecta en la columna tipada como imagen de producto (ver ejemplo), 
codifica la imagen que detecta en base64

Debes guardar la importación en Odoo 11 como conversión-odoo.xlsx
Debe existir una hoja dentro del libro Excel llamada 'comercio'
El campo que convierte el script url2base64 debe denominarse la columna como 'image',
tal como la reconoce la documentación de Odoo.

'''

def HuboError(cTexto):
  print(cTexto)
  oFichero.write(str(cTexto)+"\n")

oFichero = open("xls2base64v12-errores.txt", "a")
cFichero = "zerimar.jpg"
print("Codificando zerimar.jpg BASE64 ")

oImagen1 =  open( cFichero,'rb')
cImagen1 = oImagen1.read()


cImagen1b64 = base64.b64encode( cImagen1 )
#HuboError(cImagen1)
#HuboError( cImagen1b64 )
cBase64Zerimar = str( cImagen1b64 )
cBase64Zerimar = cBase64Zerimar.replace("b'", "")
cBase64Zerimar = cBase64Zerimar.replace("'", "")
cBase64Zerimar += '=' * (-len( cBase64Zerimar ) % 4) 
dComienzo = time.time()
print(dComienzo)
oHoja = oExcel.read_excel('etl.xlsx', sheet_name='comercio')
oHojaColumnaImage = oHoja['image']

nCuantosLlevo = 501
nNumerador  = 0


for nContador in oHoja.index:
    
    
    print("=======Conversor de imágenes desde columna Excel url2base64 para Odoo versión 11 jpg/JPG===========")
    
    print('Inspeccionando fila de la hoja excel '+str(nContador))
    print('Inspeccionando columna image '+str(nContador))
    cUrl = oHoja['image'][nContador]
    print("--- %s segundos ---" % (time.time() - dComienzo ))
    
       

    if  type( cUrl ) is str:
      oPatronURL = re.compile('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')
      lEncontrado = oPatronURL.search(cUrl)
      if lEncontrado:# Sí, podemos descargar la url
             
             print (cUrl) 
             #oRespuesta = urlopen(cUrl)
             oResultado = requests.get(cUrl)
             #if oResultado.ok and oResultado.content and oResultado.status_code == 200 :
             HuboError(oResultado.ok)
             HuboError(oResultado.status_code)
             oImagen = oResultado.content
             #'\033[71m'+
             #'\033[72m'+
             HuboError('Imagen descargada jpg' )
             cUrl1Base64 = base64.b64encode(oImagen)
             #.decode() 
             #cUrl2Base64 = base64.encodestring( oRespuesta.read() )
             # convert back to a string
             #cUrlStr = cUrl2Base64.decode()
             #oRespuesta.close()
             cUrl1Base64 = str(cUrl1Base64)
             cUrl1Base64 = cUrl1Base64.replace("b'", "")
             cUrl1Base64 = cUrl1Base64.replace("'", "")
            
             #cUrlStr     += '=' * (-len( cUrlStr ) % 4)
             cUrl1Base64 += '=' * (-len( cUrl1Base64 ) % 4)
             #oRespondeme = requests.get(cFichero)
             #oImg = Image.open(BytesIO( oRespondeme.content))
             #print(cUrl1Base64)
             HuboError('Imagen codificada jpg'  )
             HuboError(cUrl)
             
             oHoja['image'][nContador] = cUrl1Base64
             cIdentificador = oHoja['id'][nContador]
             cNombre  = oHoja['name'][nContador]
             print('\033[92m'+"BASE64- "+cUrl1Base64[0:40])
             print('\033[72m'+" ID "+str(cIdentificador) )
             nCuantosLlevo   += 1
             if nCuantosLlevo > 500:
                nCuantosLlevo  = 1
                nNumerador  += 1

                oFicheroCSV = open('odoo-parte'+str(nNumerador)+'.csv', 'a')
                oFlujoCSV = csv.writer(oFicheroCSV, delimiter=';', quotechar='|', quoting=csv.QUOTE_MINIMAL) 
                oFlujoCSV.writerow([ 'name' ,'id', 'image' ]) 
                HuboError("Creado lote nº "+str(nNumerador))
             HuboError("Grabando línea CSV " + str(nCuantosLlevo)+" perteneciente al lote nº "+str(nNumerador))
             oFlujoCSV.writerow( [  cNombre, cIdentificador, cUrl1Base64 ],  )
             
             HuboError('Columna imagen convertida a BASE64 en la fila ' + str(nContador) )
              
      else:    
        #cUrl != "" and len(cUrl)> 0 and
        #oHoja['image'][nContador] = cBase64Zerimar
        HuboError( cUrl )
        HuboError('Columna imagen sin URL.'+str(nContador) )
    else:    
        #  oHoja['image'][nContador] = cBase64Zerimar
        HuboError( cUrl )
        HuboError('Columna imagen sin URL.'+str(nContador) )
    print('+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++'+"\n") 
    #os.system('clear')    
oExcelSalvar = oExcel.ExcelWriter('conversión-odoo.xlsx')
oHoja.to_excel( oExcelSalvar,'comercio',index=False)
oExcelSalvar.save()
nSegundos = time.time() - dComienzo
print("--- %s segundos ---" % ( nSegundos)) 
print("--- %s minutos ---" % ( nSegundos/60)) 
print("=Conversor ha terminado ===========")
oFichero.close()