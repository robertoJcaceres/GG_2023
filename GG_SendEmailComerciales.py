import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Configuración del servidor SMTP
smtp_server = 'smtp.office365.com'
smtp_port = 587
smtp_username = 'reportes@ggrafica.es'
smtp_password = 'QvTYc,Z3wF'

# Dirección de correo electrónico de origen y destinatarios
from_email = 'reportes@ggrafica.es'
to_emails = ['rcaceres@ggrafica.es']

# Asunto y cuerpo del correo electrónico
subject = 'Reporte semanal'
body = 'Adjunto encontrarás el reporte semanal de ventas.'

# Ruta de los archivos adjuntos
adjuntos = [r'C:\Users\rcaceres\Desktop\GGrafica-Coding\GGrafica\StkSinMovimiento_DAVID MARTINEZ (BERNABEU).xlsx']

# Configurar el mensaje del correo electrónico
message = MIMEMultipart()
message['From'] = from_email
message['To'] = ', '.join(to_emails)
message['Subject'] = subject

# Agregar el cuerpo del mensaje
message.attach(MIMEText(body, 'plain'))

# Adjuntar los archivos Excel
for adjunto in adjuntos:
    part = MIMEBase('application', 'octet-stream')
    with open(adjunto, 'rb') as file:
        part.set_payload(file.read())
    encoders.encode_base64(part)
    filename = os.path.basename(adjunto)  # Obtener solo el nombre del archivo
    part.add_header('Content-Disposition', f'attachment; filename={filename}')
    message.attach(part)

# Enviar el correo electrónico
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()
    server.login(smtp_username, smtp_password)
    server.send_message(message)

print('Correo enviado exitosamente.')
