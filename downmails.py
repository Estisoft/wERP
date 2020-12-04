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
# Download mails
# 

# Importamos todo lo necesario
import sys
import os.path
import poplib
from datetime import date
from datetime import timedelta
from email import utils
from email.header import decode_header
from email.parser import Parser
from email.utils import getaddresses

# v1.0
# Funcion de extracción del texto desde el HTML
# http://effbot.org/zone/textify.htm
def textify(html_snippet, maxwords=50):
    import formatter, htmllib, StringIO, string
    class Parser(htmllib.HTMLParser):
        def anchor_end(self):
            self.anchor = None
    class Formatter(formatter.AbstractFormatter):
        pass
    class Writer(formatter.DumbWriter):
        def send_label_data(self, data):
            self.send_flowing_data(data)
            self.send_flowing_data(" ")
    o = StringIO.StringIO()
    p = Parser(Formatter(Writer(o)))
    p.feed(html_snippet)
    p.close()
    words = o.getvalue().split()
    if len(words) <= 2*maxwords:
        return string.join(words)
    return string.join(words[:maxwords]) + " ..."

# v1.0	
# Parseamos una direccion de correo
def addrparser(addrstring):
    addrlist = ['']
    quoted = False

    # ignore comma at beginning or end
    addrstring = addrstring.strip(',')

    for char in addrstring:
        if char == '"':
            # toggle quoted mode
            quoted = not quoted
            addrlist[-1] += char
        # a comma outside of quotes means a new address
        elif char == ',' and not quoted:
            addrlist.append('')
        # anything else is the next letter of the current address
        else:
            addrlist[-1] += char

    return getaddresses(addrlist)
	
	
# v1.0	
# Leemos los argumentos desde la linea de comandos
# Parametros:
# 	inputmailserver: Servidor
# 	inputmailaccount: Cuenta
# 	inputmailpassword: Contraseña
# 	inputport = Puerto
# 	inputssl = SSL (Y/N)
# 	inputpermanence = Días de permanencia de mensajes en el servidor

# Preparamos las variables de entrada
inputserver = ''
inputport = ''
inputssl = ''
inputmailaccount = ''
inputmailpassword = ''
inputpermanence = 30

inputserver = sys.argv[1]
inputmailaccount = sys.argv[2]
inputmailpassword = sys.argv[3]
inputport = sys.argv[4]
inputssl = sys.argv[5]
inputpermanence = int(sys.argv[6])

# Establecemos los valores por defecto
if inputpermanence == 0:
   inputpermanence = 30

# Preparamos la carpeta de trabajo   
if not (os.path.exists(inputmailaccount.replace('@', '.'))):
		os.makedirs(inputmailaccount.replace('@', '.'))

# Guardamos el log en el directorio del usuario
try:
	log = open(inputmailaccount.replace('@', '.')+'/logdown.txt', 'wb')
except:
    log.write('No se pudo preparar el entorno' + '\n')
	
# Se establece conexion con el servidor POP
try:
    if (inputssl=="Y"):
       m = poplib.POP3_SSL(inputserver, inputport)
    if (inputssl=="N"):
       m = poplib.POP3(inputserver, inputport)
except:
    log.write('Error de conexión con el servidor POP' + '\n')

# Se establecen las credenciales de conexion con el servidor POP
try:
    m.user(inputmailaccount)
    m.pass_(inputmailpassword)
except:
    log.write('Error de autenticación POP' + '\n')

# Se obtiene el numero de mensajes
try:
    messagesnumber = len(m.list()[1])
except:
    log.write('Error recuperando la lista de mensajes desde el servidor POP' + '\n')

deletemessages = []

# Recorremos cada mensaje
for i in range (messagesnumber):
    # Leemos la cabecera del mensaje
    resp, header, bytes = m.top(i + 1, 0)
    header = '\n'.join(header)
	# Parseamos
    x = Parser()	
    header = x.parsestr(header)

    # Obtenemos el From
    fromaccount = header.get("From")
    fromaccount = utils.parseaddr(fromaccount)
    fromaccountoriginal = str(fromaccount[1])
    fromaccount = str(fromaccount[1])
    fromaccount = fromaccount.replace("@", ".")
		
    #Obtenemos la Fecha
    messagedate = header["Date"]
    messagedate = utils.parsedate(messagedate)
	# Preparamos el nombre de la carpeta en función de la fecha del mensaje
    if messagedate:
        folderyear = str(messagedate[0])
        foldermonth = str(messagedate[1]).rjust(2, "0")
        folderday = str(messagedate[2]).rjust(2, "0")
        folderhour = str(messagedate[3]).rjust(2, "0")
        folderminute = str(messagedate[4]).rjust(2, "0")
        foldersecond = str(messagedate[5]).rjust(2, "0")
    else:
        folderyear = str(2000)
        foldermonth = str(00)
        folderday = str(00)
        folderhour = str(00)
        folderminute = str(00)
        foldersecond = str(00)

    # Anotamos que mensajes borrar. Se borrará si es antiguo (inputpermanence)
    if inputpermanence > 0:
        today = int(str(date.today() - timedelta(days=inputpermanence)).replace("-", ""))
    if int(folderyear+foldermonth+folderday) < today:
        deletemessages.append(i+1)

    # Comprobamos si ya se descargó el correo
    currentpath = inputmailaccount.replace("@", ".")+"/"+fromaccount+"/inbox"+"/"+folderyear+foldermonth+folderday+folderhour+folderminute+foldersecond+"/"
    print ("---" + currentpath)
	
	# Descargamos el correo completo
	# http://chuwiki.chuidiang.org/index.php?title=Enviar_y_leer_email_con_python_y_gmail
	# http://code.ohloh.net/project?pid=hsZeNqdiLuQ&did=python-gmail&cid=6j_iiFxcLDQ&fp=275900&mp=&projSelected=true
	# http://www.portalhacker.net/index.php?topic=139358.0
	# http://listas.gcoop.coop/pipermail/gnu-linux/2011-June/000087.html
    if not (os.path.exists(currentpath)):
        os.makedirs(currentpath)

        # Si no está lo descargamos
        response, headerLines, bytes = m.retr(i+1)

        # Guardamos todo el mensaje en un unico string
        messagecontent='\n'.join(headerLines)

        # Parsea el mensaje
        p = Parser()
        email = p.parsestr(messagecontent)

        #Se guarda el mensaje completo EML
        filepath = currentpath + "index.eml"
        fp = open(filepath,'wb')
        fp.write(messagecontent)
        fp.close()

        # Se sacan por pantalla los datos básicos
		# Date
        print ("Date: "+folderyear+foldermonth+folderday+folderhour+folderminute+foldersecond)

		# From		
        fromaccount = email["From"]
        if ( fromaccount == "" ):
            fromaccount = "From"
        print ("From: "+fromaccount)
		
        # Obtenemos el To
        toaccount = header.get('To')
        if (toaccount is None) or (len(toaccount) == 0):
			toaccount = ""
        else:
			toaccount = addrparser(toaccount)

        # Obtenemos un string con las cuentas del To
        toaccounts = ""
        for name, address in toaccount:
            if (toaccounts == ""):
                toaccounts = address
            else:
                toaccounts = toaccounts + "," + address

        print ("To: " + toaccounts)	

        # Temp
        tos = email.get_all('to', [])
        ccs = email.get_all('cc', [])
        resent_tos = email.get_all('resent-to', [])
        resent_ccs = email.get_all('resent-cc', [])
		
        # Obtenemos el CC
        ccaccount = header.get('cc')
        if (ccaccount is None) or (len(ccaccount) == 0):
			ccaccount = ""
        else:
			ccaccount = addrparser(ccaccount)	
        
		# Obtenemos un string con las cuentas del To
        ccaccounts = ""
        for name, address in ccaccount:
            if (ccaccounts == ""):
                ccaccounts = address
            else:
                ccaccounts = ccaccounts + "," + address

        print ("Cc: " + ccaccounts)	
	
        emailsubject = email["Subject"]
        emailsubject = decode_header(emailsubject)
        emailsubject = emailsubject[0]
        charset = emailsubject[1]
        if not charset:
            charset = 'utf8'
        emailsubject = str(emailsubject[0])
        emailsubject = emailsubject.replace("\n", " ")
        emailsubject = unicode(emailsubject, charset, 'replace').encode('utf8', 'replace')
        if ( emailsubject == "" ):
            emailsubject = "No subject"
		# Subject	
        print ("Subject: "+emailsubject)

        # Preparamos las variables de entorno
        html = ""
        text = ""
        attachments = ""
		
        # Diferenciamos los mensajes multipart de los simples
        if (email.is_multipart()):
            # Leemos cada parte del mensaje
            for part in email.walk():
                for subpart in part.walk():
                    # Se mira el mime type de la subparte
                    contenttype = subpart.get_content_type()
                    # Leemos el charset
                    charset = part.get_content_charset()
                    if not charset:
                        charset = 'utf8'
                    # Si es texto plano, se guarda el txt
                    if ("text/plain" == contenttype):
                        text = unicode(subpart.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                    # Si es html se guarda en memoria para modificar los enlaces inline
                    if ("text/html" == contenttype):
                        html = unicode(subpart.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')						
						
                # Se mira el mime type de la parte
                contenttype = part.get_content_type()
                # Leemos el charset
                charset = part.get_content_charset()
                if not charset:
                    charset = 'utf8'
                # Si hay filename o content-id es un adjunto
                if (part.get_filename()) or part.get('content-id'):
                    # Vemos si es inline
                    if str(part.get('content-disposition')).count(';') == 1:
                        ctype = part.get('content-disposition').split(';')[0]
                    else:
                        ctype = part.get('content-disposition')

                    # leemos el nombre del adjunto
                    filepath = part.get_filename()

                    # Leemos el identificador de la parte
                    cid = part.get('content-id')
                    bad_characters = ["<", ">"]
                    for letter in bad_characters:
                        if cid:
                            cid = cid.replace(letter, "")

                    # Si tenemos CID
                    if cid:

                        if html.count('src="cid:' + cid):
                            filepath = cid
                            link = cid.replace("@", ".")
                            link = link.replace("/", ".")
                            link = link.replace("//", ".")
                            link = link.replace("\\", ".")
                            # Reemplazamos el identificador en el html si existe
                            html = html.replace('src="cid:' + filepath, 'src="' + link)
                            print ("filepath: " + filepath)							
                            # Guardamos el fichero
                            fp = open(currentpath + link, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()

                        elif filepath:
                            if html.count('src="cid:' + part.get_filename()):
                                filepath = part.get_filename()
                                bad_characters = ["/", "//", "\\", ":", "(", ")", "<", ">", "|", "?", "¿", "*", "@", "#", "=", "\n", "\r", "\t", "\p", "\s"]
                                for letter in bad_characters:
                                    filepath = filepath.replace(letter, "")
                                # Reemplazamos el identificador en el html si existe
                                html = html.replace('src="cid:' + filepath, 'src="' + filepath)
                                # Guardamos el fichero
                                fp = open(currentpath + filepath, 'wb')
                                fp.write(part.get_payload(decode=True))
                                fp.close()

                    # Si no tenemos CID
                    else:
                        filepath = part.get_filename()
                        bad_characters = ["/", "//", "\\", ":", "(", ")", "<", ">", "|", "?", "¿", "*", "@", "#", "=", "\n", "\r", "\t", "\p", "\s"]
                        for letter in bad_characters:
                            filepath = filepath.replace(letter, "")

                        # Si es html o texto y la disposicion es inline lo añadimos al cuerpo
                        if ctype == "inline" and contenttype == "text/html":
                            if (html == ""):
                                html = unicode(part.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                        elif ctype == "inline" and contenttype == "text/plain":
                            if (text == ""):
                                text = unicode(part.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                        else:
                            # Si no, guardamos el fichero
                            if not (os.path.exists(currentpath + "attach/")):
                                os.makedirs(currentpath + "attach/")
                            fp = open(currentpath + "attach/" + filepath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                            # Guardamos el nombre del adjunto para el log
                            attachments = attachments + filepath + ";"

                else:
                    # Si es texto plano, se guarda el txt
                    if (text == ""):					
                        if ("text/plain" == contenttype):
                            text = unicode(part.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                    # Si es html se guarda en memoria para modificar los enlaces inline
                    if (html == ""):
                        if ("text/html" == contenttype):
                            html = unicode(part.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')

            if text:
                filepath = currentpath + "index.txt"
                fp = open(filepath, 'wb')
                fp.write(text)
                fp.close()

            if html:
                filepath = currentpath + "index.htm"
                fp = open(filepath, 'wb')
                fp.write(html)
                fp.close()

        else:
            # Si es un mensaje simple leemos el html y el plano de forma directa
            contenttype = email.get_content_type()
            charset = email.get_content_charset()
            if not charset:
                charset = 'utf8'

            if ("text/plain" == contenttype):
                filepath = currentpath + "index.txt"
                fp = open(filepath, 'wb')
                data = unicode(email.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                fp.write(data)
                fp.close()

            if ("text/html" == contenttype):
                filepath = currentpath + "index.htm"
                fp = open(filepath, 'wb')
                data = unicode(email.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                fp.write(data)
                fp.close()
                filepath = currentpath + "index.txt"
                fp = open(filepath, 'wb')
                data = unicode(email.get_payload(decode=True), charset, 'replace').encode('utf8', 'replace')
                fp.write(textify(data, 1000))
                fp.close()

		# Si no pudimos extraer el mensaje en formato txt lo obtenemos desde el html
        if not (os.path.exists(currentpath + "index.txt")):
            filepath = currentpath + "index.txt"
            html = unicode(html, charset, 'replace').encode('utf8', 'replace')
            fp = open(filepath, 'wb')
            fp.write(textify(html, 10000))
            fp.close()

        # Limpiamos el subject de | para no romper el formato de salida del fichero
        emailsubject = emailsubject.replace("|", "-")

        # Guardamos el log de los descargados
        fichero = inputmailaccount.replace("@", ".") + "/download.txt"
        log = open(fichero, 'a')
        log.write("|R|" + fromaccountoriginal + "|F|" + folderyear + foldermonth + folderday + "|H|" + folderhour + folderminute + foldersecond + "|A|" + emailsubject + "|X|" + attachments + "|S|" + currentpath + "|O|" + toaccounts + "|C|" + ccaccounts + "\n")
        log.close()

# Borramos los mensajes marcados para borrar
for b in deletemessages:
    print (b)
    #m.dele(b)
    
log.close()
# Cierre de la conexion
m.quit()
