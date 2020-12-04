#!/usr/bin/python
# -*- coding: utf-8 -*-
# ------------------------------------------------------------------------
# PaaSOS
# TipeSoft S.L.U.
# 
# Copyright (C) 2009-2014, TipeSoft S.L.U.
# 
# Licensed under the EUPL V.1.1
# 
# The European Union Public Licence (the "EUPL") applies to the Work
# or Software which is provided under the terms of this Licence.
# Any use of the Work, other than as authorised under this Licence
# is prohibited (to the extent such use is covered by a right
# of the copyright holder of the Work).
# 
# The Work is provided under the Licence on an "as is" basis and
# WITHOUT WARRANTIES OF ANY KIND concerning the Work,
# including without limitation MERCHANTABILITY, FITNESS FOR A
# PARTICULAR PURPOSE, ABSENCE OF DEFECTS OR ERRORS, ACCURACY,
# NON-INFRINGEMENT OF INTELLECTUAL PROPERTY RIGHTS other than
# copyright as stated in Article 6 of the Licence.
# 
# You should have received a copy of the European Union Public Licence
# along with this program.  If not, see http://ec.europa.eu/idabc/eupl
# 
# ------------------------------------------------------------------------
# 
# Send mails
# 

import os
import smtplib
import sys
import datetime

from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email.Utils import formatdate
from email import Encoders
		
def win32_unicode_argv():
    """Uses shell32.GetCommandLineArgvW to get sys.argv as a list of Unicode
    strings.

    Versions 2.x of Python don't support Unicode in sys.argv on
    Windows, with the underlying Windows API instead replacing multi-byte
    characters with '?'.
    """

    from ctypes import POINTER, byref, cdll, c_int, windll
    from ctypes.wintypes import LPCWSTR, LPWSTR

    GetCommandLineW = cdll.kernel32.GetCommandLineW
    GetCommandLineW.argtypes = []
    GetCommandLineW.restype = LPCWSTR

    CommandLineToArgvW = windll.shell32.CommandLineToArgvW
    CommandLineToArgvW.argtypes = [LPCWSTR, POINTER(c_int)]
    CommandLineToArgvW.restype = POINTER(LPWSTR)

    cmd = GetCommandLineW()
    argc = c_int(0)
    argv = CommandLineToArgvW(cmd, byref(argc))
    if argc.value > 0:
        # Remove Python executable and commands if present
        start = argc.value - len(sys.argv)
        return [argv[i] for i in
                xrange(start, argc.value)]
				
			
# v1.0	
# Leemos los argumentos desde la linea de comandos
# Parametros:
# 	outputserver: Servidor
# 	outputport = Puerto
# 	outputmailaccount: Cuenta
# 	outputmailpassword: Contraseña
# 	outputmaildisplayname: Nombre a mostrar
# 	outputmailaddress: Dirección de correo
# 	outputmailtoaddress: Dirección de correo del destinatario
# 	outputmailsubject: Dirección de correo del destinatario
# 	outputmailfiletxt: Path al fichero txt con el mensaje de correo
# 	outputmailfilehtml: Path al fichero html con el mensaje de correo
# 	outputmailattachments: Anexos
# 	outputmailimages: Imágenes incrustadas
# 	outputmailacknowledgment: Acuse de recibo
# 	outputssl: SSL-TLS

now = datetime.datetime.now()

print "Envio de correo iniciado: " + str(now)

if sys.version_info[0] < 3 and os.name == 'nt':
    sys.argv = win32_unicode_argv()

print "Leyendo argumentos: " + str(now)

# Leemos argumentos
outputserver = sys.argv[1]
outputport = sys.argv[2]
outputmailaccount = sys.argv[3]
outputmailpassword = sys.argv[4]
outputmaildisplayname = sys.argv[5]
outputmailaddress = sys.argv[6]
outputmailtoaddress = sys.argv[7]
outputmailccaddress = sys.argv[8]
outputmailsubject = sys.argv[9]
outputmailfiletxt = sys.argv[10]
outputmailfilehtml = sys.argv[11]
outputmailattachments = sys.argv[12]
outputmailimages = sys.argv[13]
outputmailacknowledgment = sys.argv[14]
outputssl = sys.argv[15]

# Convertimos las cadenas en listas
outputmailattachments = outputmailattachments.split(";")
outputmailimages = outputmailimages.split(";")

# Verificamos que contengan direcciones
if (outputmailtoaddress == "N"):
    outputmailtoaddress = ""
if (outputmailccaddress == "N"):
    outputmailccaddress = ""

# Verificamos que son listas de elementos
if outputmailattachments:
    assert type(outputmailattachments) == list
if outputmailimages:
    assert type(outputmailimages) == list

if not (os.path.exists(outputmailaccount.replace('@', '.'))):
        os.makedirs(outputmailaccount.replace('@', '.'))
log = open(str(outputmailaccount).replace('@', '.')+'/logsend.txt', 'wb')

print "Preparando el mensaje: " + str(now)

# Preparamos el mensaje a enviar
# http://www.peterbe.com/plog/zope-html-emails
messagemime = MIMEMultipart('related')

# Añadimos la cabecera
messagemime['From'] = outputmaildisplayname + ' <' + outputmailaddress + '>'
messagemime['To'] = outputmailtoaddress
messagemime['Cc'] = outputmailccaddress
messagemime['Subject'] = outputmailsubject
messagemime['Date'] = formatdate(localtime=True)

# Añadimos el comprobante
if outputmailacknowledgment:
    messagemime ['Disposition-Notification-To'] = outputmailaddress

# Enviamos los dos formatos text y html
messagealternative = MIMEMultipart('alternative')
messagemime.attach(messagealternative)

# Añadimos el mensaje de texto alternativo
if outputmailfiletxt:
    f = open(outputmailfiletxt, 'rb')
    stringtext = MIMEText( f.read(), 'plain', "ISO-8859-1")
    messagealternative.attach(stringtext)

# Añadimos el mensaje html como principal
if outputmailfilehtml:
    f = open(outputmailfilehtml, 'rb')
    stringhtml = MIMEText( f.read(), 'html', "UTF-8")
    messagealternative.attach(stringhtml)

print "Añadimos imágenes al mensaje: " + str(now)

# Añadimos imagenes
try:
    for imagen in outputmailimages:
        if (imagen != ""):
            # Cargarmos la imagen desde el fichero binario
            fileimage = open(imagen, 'rb')
            messageimage = MIMEImage(fileimage.read())
            fileimage.close()

            # Hemos de adjuntar la imagen en el content-id.
            # En el archivo html se debe hacer referencia al content-id
            # como fuente en el source de la imagen, por ejemplo:
            # <img src="cid:/nombre/de_la_ruta_entera/imagen.jpg">
            messageimage.add_header('Content-ID', '<' + imagen + '>')
            messagemime .attach(messageimage)
except:
    log.write('Error al anexar imagenes como Content-ID'+'\n')

print "Añadimos anexos al mensaje: " + str(now)

# Anadimos anexos
try:
    for line in outputmailattachments:
        if (line != ""):
            attachment = MIMEBase('application', "octet-stream")
            attachment.set_payload(open(line, "rb").read())
            Encoders.encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment; filename = "%s"' % os.path.basename(line))
            messagemime.attach(attachment)
except:
    log.write('Error al anexar ficheros'+'\n')
	
# Enviamos el mensaje mediante SMTP
try:
    manager = smtplib.SMTP(outputserver)
except:
    log.write('Error al conectar'+'\n')

#manager.set_debuglevel(1)

try:
    manager.ehlo()
    if (outputssl == "Y"):
        manager.starttls()
except:
    log.write('Error al hacer EHLO'+'\n')

try:
    manager.login(outputmailaccount, outputmailpassword)
except:
    log.write('Error de autenticación'+'\n')

# multiple-recipients
# http://www.jmcpdotcom.com/blog/2012/01/06/multiple-recipients-with-pythons-smtplib-sendmail-method/
# limpiamos la cadena del To asegurandonos de que cumple el formato delimitado por espacios Por ejemplo:malcom@example.com,reynolds@example.com,firefly@example.com
try:
	print "Enviamos el mensaje: " + str(now)
	outputmailtoaddresses = outputmailtoaddress.replace(";", ",") + "," + outputmailccaddress.replace(";", ",")
	manager.sendmail(outputmailaddress, outputmailtoaddresses.split(","), messagemime.as_string())
except:
    log.write('Error en el envío'+'\n')

log.close()
manager.quit()


