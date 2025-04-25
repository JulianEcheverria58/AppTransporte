import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from config import EMAIL_FROM, EMAIL_TO, SMTP_SERVER, SMTP_PORT, USERNAME, PASSWORD
import os

class EmailSender:
    def send_report(self, html_content, report_date, report_type='detallado'):
        """Envía reporte HTML por email"""
        msg = MIMEMultipart('related')
        msg['Subject'] = f"Reporte {report_type} - {report_date.strftime('%d/%m/%Y')}"
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        
        msg.attach(MIMEText(html_content, 'html'))
        
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(USERNAME, PASSWORD)
                server.send_message(msg)
            return True, "Correo enviado exitosamente"
        except Exception as e:
            return False, f"Error al enviar correo: {str(e)}"

    def send_with_attachment(self, subject, body, file_path):
        """Envía email con archivo adjunto"""
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = EMAIL_FROM
        msg['To'] = EMAIL_TO
        
        msg.attach(MIMEText(body, 'plain'))
        
        with open(file_path, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)
        
        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(USERNAME, PASSWORD)
                server.send_message(msg)
            return True, "Correo con adjunto enviado"
        except Exception as e:
            return False, f"Error al enviar adjunto: {str(e)}"