#Envio masivo de pdfs por correo

## Contexto
____________
Se tienen una cantidad de archivos para enviar en forma individual por correo.
Haremos el ejemplo con pdfs, aunque no hay problema, si es otro tipo de archivo, esto es configurable.

Luego de la generación de los archivos en una carpeta x con un nombre especifico,
se requiere enviar los pdfs generados individualmente a unas casillas de correo electrónico
 para cada persona. Pues el contenido es diferente para cada persona, no es el mismo adjunto para todos. 
 Ver proyecto cuentas de cobro.
y el correo es según el dueño del adjunto, se debe enviar el archivo de acuerdo 
a una ruta establecida donde estaran los archivos.
Toda la información que se necesita para el envio estara en una hoja electrónica que contiene:
un consecutivo del archivo, nombre del artista, correo al que se envia, asunto del correo, cuerpo del correp y la ruta del archivo.

# Solución
En el archivo en Excel se han organizado previamente con la información necesaria a los destinatarios.
Se han estandarizado las rutas del archivo en una carpeta.

1. Se procede a leer el archivo en Excel
2. Se debe configurar las credenciales del correo saliente :
(smtp_server, puerto, usuario,contraseña y remitente)
3. Se recorre los registros en excel 
4. se crea el mensaje
5. se adjunta el archivo de acuerdo a la ruta definida en el excel
6. se envia el correo electrónico con el remitente, correo, adjunto
7. se cierra la conexion SMTP

#REQUISITOS
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

#Archivo en excel con los campos:
nCuenta_cobro,	nombre_del_artista,	correo, Asunto_del_correo ,	 Cuerpo_del_correo,
 Carpeta_donde_estan_los_archivos ,	 RutaArchivoAdjuntar

 # Video Explicativo
 https://www.youtube.com/watch?v=u7IFh7IpJ5c 
