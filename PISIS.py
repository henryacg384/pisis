from datetime import datetime
import sys
import ftplib
import os
from ftplib import FTP
from shutil import rmtree
from os import remove
from os import path
import selenium
import pathlib
import urllib
import shutil
import smtplib
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import selenium
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import date
from time import sleep
import os, time, base64
import openpyxl
import pandas as pd
import xlrd as xl
import zipfile
import shutil
from zipfile import ZipFile
import os

#ruta="E:/arus/varios/"
#options = webdriver.ChromeOptions()
#options.add_argument('--no-sandbox')
#options.add_argument("--lang=es")
#options.add_argument('--allow-running-insecure-content')
#options.add_argument('--ignore-certificate-errors')
#options.add_argument("--start-maximized")
#options.add_argument('--ignore-ssl-errors')
#capabilities = DesiredCapabilities.CHROME.copy()
#capabilities['acceptSslCerts'] = True
#capabilities['acceptInsecureCerts'] = True
#capabilities['browserName']='chrome'
#capabilities['javascriptEnabled']=True
#fecha_actual=ct.strftime("%d-%m-%y")
#hora=str(ct.strftime("%H"))
#horas= ct.strftime("%H:%M")

def esperar_elemento(_xpath):
    element = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, _xpath)))
    return element



def elemento_clickeable(_xpath):
    clickeable = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, _xpath)))
    return clickeable

def captura_elemento_picture(elemento,nombre,ruta_temporal):
    driver.execute_script("return arguments[0].scrollIntoView();", elemento)
    elemento.screenshot(ruta_temporal+nombre+".png")



def esperar_elemento2(_xpath):
    element = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, _xpath)))
    driver.execute_script("arguments[0].scrollIntoView()",element)
    driver.execute_script('window.scrollBy(0, -100)')
    return element



def elemento_clickeable2(_xpath):
    clickeable = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, _xpath)))
    driver.execute_script("return arguments[0].scrollIntoView();", clickeable)
    time.sleep(0.2)
    driver.execute_script("return arguments[0].click();", clickeable)



def get_file_names_moz(driver):
  driver.command_executor._commands["SET_CONTEXT"] = ("POST", "/session/$sessionId/moz/context")
  driver.execute("SET_CONTEXT", {"context": "chrome"})
  return driver.execute_async_script("""
    var { Downloads } = Components.utils.import('resource://gre/modules/Downloads.jsm', {});
    Downloads.getList(Downloads.ALL)
      .then(list => list.getAll())
      .then(entries => entries.filter(e => e.succeeded).map(e => e.target.path))
      .then(arguments[0]);
    """)
  driver.execute("SET_CONTEXT", {"context": "content"})




def get_file_content_moz(driver, path):
  driver.execute("SET_CONTEXT", {"context": "chrome"})
  result = driver.execute_async_script("""
    var { OS } = Cu.import("resource://gre/modules/osfile.jsm", {});
    OS.File.read(arguments[0]).then(function(data) {
      var base64 = Cc["@mozilla.org/scriptablebase64encoder;1"].getService(Ci.nsIScriptableBase64Encoder);
      var stream = Cc['@mozilla.org/io/arraybuffer-input-stream;1'].createInstance(Ci.nsIArrayBufferInputStream);
      stream.setData(data.buffer, 0, data.length);
      return base64.encodeToString(stream, data.length);
    }).then(arguments[1]);
    """, path)
  driver.execute("SET_CONTEXT", {"context": "content"})
  return base64.b64decode(result)
capabilities_moz = { \
    'browserName': 'firefox',
    'marionette': True,
    'acceptInsecureCerts': True,
    'moz:firefoxOptions': { \
      'args': [],
      'prefs': {
        # 'network.proxy.type': 1,
        # 'network.proxy.http': '12.157.129.35', 'network.proxy.http_port': 8080,
        # 'network.proxy.ssl':  '12.157.129.35', 'network.proxy.ssl_port':  8080,
        'browser.download.dir': '',
        'browser.helperApps.neverAsk.saveToDisk': 'application/octet-stream,application/vnd.ms-excel,application/pdf',
        'browser.download.useDownloadDir': True,
        'browser.download.manager.showWhenStarting': False,
        'browser.download.animateNotifications': False,
        'browser.safebrowsing.downloads.enabled': False,
        'browser.download.folderList': 2,
        'pdfjs.disabled': True
      }
    }
  }


############inicia proceso#######################

print("INICIANDO EL PROCESO")
sleep(3)


ftp=FTP("10.0.0.106")
ftp.login("cedex.pisis","=aA3~dAhtfS4j]Q")

def downloadFiles(path,destination):
#path & destination are str of the form "/dir/folder/something/"
#path should be the abs path to the root FOLDER of the file tree to download
    try:
        ftp.cwd(path)
        #clone path to destination
        os.chdir(destination)
        #os.mkdir(destination[0:len(destination)-1]+path)
        print(destination[0:len(destination)-1]+path+" built")
    except OSError:
        #folder already exists at destination
        pass
    except ftplib.error_perm:
        #invalid entry (ensure input form: "/dir/folder/something/")
        print("error: could not change to "+path)
        sys.exit("ending session")

    #list children:
    filelist=ftp.nlst()
    print(filelist)
    for file in filelist:
        try:
            #this will check if file is folder:
            ftp.cwd(path+file+"/")
            #if so, explore it:
          # downloadFiles(path+file+"/",destination)
        except ftplib.error_perm:
            #not a folder with accessible content
            #download & return
            #os.chdir(destination[0:len(destination)-1]+path)
            #possibly need a permission exception catch:
            ftp.retrbinary("RETR "+file, open(os.path.join("/tmp/",file),"wb").write)
            print(file + " downloaded")
    return


# %Y%m%d fecha completa %Y% ano m% mes d dia

fecha_completa = datetime.today().strftime('%Y%m%d')
year=datetime.today().strftime('%Y')
month= datetime.today().strftime('%-m')
day= datetime.today().strftime('%-d')
print (fecha_completa)

#fecha con un digito
#%Y%-m%-d   %Y% ano -m% mes sin un digito -d día sin un digito
fecha_un_digito  = datetime.today().strftime('%-m%-d')
print (fecha_un_digito)


#si es 31 de diciembre matar el proceso

if "1231" == fecha_un_digito:
  print("es 31 de diciembre")
  exit()

#Ingresar en el FTP

day=int(day)

source="/{}/{}/{}/".format(year,month,day) #esta es la ruta de inicio
dest="/tmp/"  #esta es la ruta destino
downloadFiles(source,dest) #descargar de la ruta de origen a la ruta destino

ruta_zip = "/tmp/{}{}{}.zip".format(year,month,day,day)
ruta_extraccion = "/tmp/{}{}{}/".format(year,month,day)
archivo_zip = ZipFile(ruta_zip, "r")
try:
    print(archivo_zip.namelist())
    archivo_zip.extractall(path=ruta_extraccion)
except:
    pass
    archivo_zip.close()


#  Eliminar carpeta Reprocesos

if os.path.isdir("/tmp/{}{}{}/{}/Reprocesos".format(year,month,day,day)):
   shutil.rmtree("/tmp/{}{}{}/{}/Reprocesos".format(year,month,day,day))
   print("ELIMINANDO CARPETA REPROCESOS")
   sleep(3)

else:
   print("CONTINUAR CON EL PROCESO")
   sleep(2)


# Validar si hay plantillas que no son del día actual y eliminar

fecha_completa="{}{}{}".format(year,month,day)
archivos=os.listdir("/tmp/{}{}{}/{}/".format(year,month,day,day))
print (len (archivos))
for file in archivos:
   archivo=file.split("PIL019PILA")
#   print("archivosplit",archivo)
   archivo=archivo[1][:-29]
#   print(archivo)
   if archivo != fecha_completa:
      print("EL ARCHIVO PRESENTA ERRORES")
      os.remove("/tmp/{}{}{}/{}/{}".format(year,month,day,day,file))
      print("PLANTILLA ELIMINADA")
#      archivos=os.listdir("/tmp/{}{}{}/{}/".format(year,month,day,day))
#      print(len(archivos))




########################## pagina web



#ingreso al zalenium

driver = webdriver.Remote('https://172.30.5.7:4444/wd/hub', capabilities_moz)

driver.get ("https://gestion.suaporte.com.co/Gestion#/login")

#ingressamos a la url

#driver.get ("https://gestion.suaporte.com.co/Gestion")
time.sleep(10)
print('INGRESO A LA PAGINA EXITOSO')


#autentificacion  cedula y contraseña




#def correo():
#  remitente = 'monitoreomosaico@arus.com.co'
#  destinatarios = ['sandra.castaneda@arus.com.co','yuri.gallego@arus.com.co','soporte3@arus.com.co','alejandra.chamorro@arus.com.co','soporte20@arus.com.co','jhonatan.valencia@arus.com.co','maryory.lezcano@arus.com.co','henry.calderon@arus.com.co' ]
#  asunto = 'Falla para ingresar al aplicativo Su Aporte Gestión RPA Gestionar Archivo Pisis'
#  cuerpo = 'Cordial Saludo,\n\n Se informa que no se puede acceder al aplicativo Su Aporte Gestión.\n\n Atte:\n\n Asistente Virtual.'
#  mensaje = MIMEMultipart()
#  mensaje['From'] = remitente
#  mensaje['To'] = ", ".join(destinatarios)
#  mensaje['Subject'] = asunto
#  mensaje.attach(MIMEText(cuerpo, 'plain'))
#  sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
#  sesion_smtp.starttls()
#  sesion_smtp.login('monitoreomosaico@arus.com.co','Mayo4321*')
#  texto = mensaje.as_string()
#  sesion_smtp.sendmail(remitente, destinatarios, texto)
#  sesion_smtp.quit()
#  print("correo enviado satisfactoramente")


while True:
  i=0
  select = Select(driver.find_element_by_id('doc-types'))
  select.select_by_value('CC')
  text=driver.find_element_by_xpath("//*[@id='nro-doc-login']")
  text.send_keys("2019")
  WebDriverWait(driver, 5)\
  .until(EC.element_to_be_clickable((By.ID,"login-continue")))\
  .click()
  time.sleep(5)
  try:
    driver.find_element_by_id('5')
    break
  except:
    i=i+1
    if i ==2:
      correo()
      #print("nodio")
      exit()
print("salio")
password = [5,5,5,5]
WebDriverWait(driver, 5)\
    .until(EC.element_to_be_clickable((By.ID,password[0])))\
    .click()
sleep(5)
driver.find_element_by_id(password[1]).click()
driver.find_element_by_id(password[2]).click()
driver.find_element_by_id(password[3]).click()
WebDriverWait(driver, 5)\
     .until(EC.element_to_be_clickable((By.ID,'login')))\
     .click()
sleep(5)
#

print("INICIO DE SESION EXITOSO")

Login_suaporte = True


WebDriverWait(driver, 5)\
        .until(EC.element_to_be_clickable((By.ID,'aMenu_61')))\
        .click()
sleep(5)
WebDriverWait(driver, 5)\
        .until(EC.element_to_be_clickable((By.ID,'aSubmenu63')))\
        .click()
sleep(8)
driver.switch_to.frame(driver.find_element_by_xpath('//*[@id="iframeApp"]'))
driver.find_element_by_xpath('//*[@id="radioDetallado:0"]').click()
today = date.today()
f= today.strftime("%d/%m/%Y")
#f = "27/08/2021"
time.sleep(3)
esto=driver.find_element_by_xpath('//*[@id="tx_fechaDetallada:textoFecha"]')
driver.execute_script("arguments[0].removeAttribute('readonly')", esto)
esto.send_keys(str(f))
time.sleep(1)
driver.find_element_by_xpath('//*[@id="tx_fechaDetallada:textoFecha"]').click()
sleep(5)
driver.find_element_by_xpath('//*[@id="btDetallada"]/img').click()
sleep(3)


files = WebDriverWait(driver, 20, 1).until(get_file_names_moz)

# get the content of the last downloaded file
content = get_file_content_moz(driver, files[0])
print(content)
#nombre=files[0]
#print(nombre)
# save the content in a local file ain the working directory
with open("/home/adminapp/archivo.xlsx", 'wb') as f:
  f.write(content)


#contar numero de filas

loc = ("/home/adminapp/archivo.xlsx")          #Giving the location of the file

wb = xl.open_workbook(loc)                    #opening & reading the excel file
s1 = wb.sheet_by_index(0)                     #extracting the worksheet
s1.cell_value(0,0)                            #initializing cell from the excel file mentioned through the cell position

print("No. of rows:", s1.nrows-1)               #Counting & Printing thenumber of rows & columns respectively


if len(archivos) != s1.nrows-1:

  print ("es diferente ")

shutil.make_archive ("/tmp/{}{}{}/{}".format(year,month,day,day),"zip","/tmp/{}{}{}/{}".format(year,month,day,day))
source = "/tmp/{}{}{}/{}.zip".format(year,month,day,day)
destination = r'/opt/win/Integrador de Recursos/Operador de Información/pisis/Inconsistente/'
try:
   shutil.copy(source,destination)
except:
   print("el archivo inconsistente no se puede copiar")

# Notificacion correo
#  remitente = 'monitoreomosaico@arus.com.co'
 # destinatarios = ['sandra.castaneda@arus.com.co','yuri.gallego@arus.com.co','soporte3@arus.com.co','alejandra.chamorro@arus.com.co','soporte20@arus.com.co','jhonatan.valencia@arus.com.co','maryory.lezcano@arus.com.co','henry.calderon@arus.com.co' ]
 # asunto = 'Inconsistencia para Procesar Archivo Pisis'
 # cuerpo = 'Cordial Saludo,\n\n Se informa que realizando la ejecución del dia de hoy, se valida que el numero de planillas del archivo FTP\n es diferente al numero de planillas del informe de Gestión. \n\n  Atte:\n\n Asistente Virtual.'
 # mensaje = MIMEMultipart()
#  mensaje['From'] = remitente
 # mensaje['To'] = ", ".join(destinatarios)
 # mensaje['Subject'] = asunto
 # mensaje.attach(MIMEText(cuerpo, 'plain'))
 # sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
 # sesion_smtp.starttls()
 # sesion_smtp.login('monitoreomosaico@arus.com.co','Mayo4321*')
 # texto = mensaje.as_string()
 # sesion_smtp.sendmail(remitente, destinatarios, texto)
 # sesion_smtp.quit()
 # print("correo enviado satisfactoramente")
#  exit()


plantillas_reemplazar=os.listdir("/opt/win/Integrador de Recursos/Operador de Información/pisis/Planillas a reemplazar")
for item in  plantillas_reemplazar:
 if "txt" in item:
  shutil.copyfile("/opt/win/Integrador de Recursos/Operador de Información/pisis/Planillas a reemplazar/"+item,"/tmp{}{}{}/{}"
.format(,year,month,day,day,item))
print ("plantilla")
print("archivo reemplazado correctamente")


if os.path.isdir("/tmp/folder/"):
    print('La carpeta existe.');
else:
    print('destination');
    os.mkdir("/tmp/folder/")



if os.path.isdir("/tmp/folder/MINPS_{}-{}-{}/".format(year,month,day)):
    print('La carpeta existe.');
else:
    print('destination');
    os.mkdir("/tmp/folder/MINPS_{}-{}-{}/".format(year,month,day))
copiar_archivos=os.listdir( "/tmp/{}{}{}/{}/".format(year,month,day,day))
for item in copiar_archivos:
 shutil.copyfile("/tmp/{}{}{}/{}/{}".format(year,month,day,day,item),"/tmp/folder/MINPS_{}-{}-{}/{}".format(year,month,day,item))


#shutil.make_archive ("/tmp/{}{}{}/{}".format(year,month,day,day),"zip","/tmp/{}{}{}/{}".format(year,month,day,day))

shutil.make_archive ("/tmp/folder/".format(year,month,day),"zip","/tmp/folder/".format(year,month,day))

#shutil.make_archive ("/tmp/{}{}{}/{}".format(year,month,day,day),"zip", "/tmp/MINPS_{}-{}-{}".format(year,month,day))





destination = r'/opt/win/Integrador de Recursos/Operador de Información/pisis/Resultado/{}'.format(year)
if os.path.isdir(destination):
    print('La carpeta existe.');
else:
    print('destination');
    os.mkdir(destination)

month_largo= datetime.today().strftime('%m')

destination = r'/opt/win/Integrador de Recursos/Operador de Información/pisis/Resultado/{}/{}/'.format(year,month_largo)
if os.path.isdir(destination):
    print('La carpeta existe.');
else:
    print('destination');
    os.mkdir(destination)

source = "/tmp/folder.zip".format(year,month,day)
destination = r'/opt/win/Integrador de Recursos/Operador de Información/pisis/Resultado/{}/{}/{}.zip'.format(year,month_largo,day)
#destination = r'/opt/{}.zip'.format(day)
try:
  shutil.copy(source,destination)
except:
  print("problema al copiar el archivo")

f= datetime.today().strftime("%Y%m%d")
os.rename(destination,r'/opt/win/Integrador de Recursos/Operador de Información/pisis/Resultado/{}/{}'.format(year,month_largo)+'/PIL019LOTE{}OP000000000089091300.PIL'.format(f))
print("CAMBIO EN NOMBRE SATISFACTORIO")
######## nototificacion correo ####

#remitente = 'monitoreomosaico@arus.com.co'
#destinatarios = ['sandra.castaneda@arus.com.co','yuri.gallego@arus.com.co','soporte3@arus.com.co','alejandra.chamorro@arus.com.co','soporte20@arus.com.co','jhonatan.valencia@arus.com.co','maryory.lezcano@arus.com.co','henry.calderon@arus.com.co']
#asunto = 'Gestión RPA Gestionar Archivo Pisis {0}'
#cuerpo = 'Cordial Saludo, \n\n Se informa que el archivo para cargar en la plataforma Pisis quedo gestionado. En total el numero de planillas en el archivo son {}  \n\n El número total de planillas que se remplazaron fue de {} \n\n El archivo quedo almacenado en la Ruta: {}. \n\n atte:  \n\n Asistente Virtual.'

#mensaje = MIMEMultipart()
#mensaje['From'] = remitente
#mensaje['To'] = ", ".join(destinatarios)
#mensaje['Subject'] = asunto
#mensaje.attach(MIMEText(cuerpo, 'plain'))
#sesion_smtp = smtplib.SMTP('smtp.office365.com', 587)
#sesion_smtp.starttls()
#sesion_smtp.login('monitoreomosaico@arus.com.co','Mayo4321*')
#texto = mensaje.as_string()
#sesion_smtp.sendmail(remitente, destinatarios, texto)
#sesion_smtp.quit()
#print("correo enviado satisfactoramente")
exit()

