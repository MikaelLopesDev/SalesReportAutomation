import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


class EmailSender:
    def __init__(self,smtp_server,smtp_port,sender_email,sender_password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = sender_password
        self.attachments = []

    def add_attachment(self,file_path):
        self.attachments.append(file_path)

    def send_email(self, recipient_email, subject, body):

        msg = MIMEMultipart()
        msg['From'] = self.sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        for file_path in self.attachments:
            try:
                with open(file_path, 'rb') as file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {file_path.split("/")[-1]}'
                    )
                    msg.attach(part)
            except FileNotFoundError:
                print(f"Erro: File Not Found - {file_path}")

                # Send Email
        try:
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.sender_email, self.sender_password)
                server.send_message(msg)
            print("E-mail Sent Successfully!")
        except Exception as e:
            print(f"Failed sent e-mail: {str(e)}")
