import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

# Configuracoes padroes do email
password = open("senha", "r").read()
from_email = "cursosthiz@gmail.com"
to_email = "cursosthiz@gmail.com"
subject = "Automacao"
body = """
Ola. Segue em anexo a automacao da planilha para a empresa XYZ

Qualquer duvida estou a disposicao!
"""

# Cria a mensagem
message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

message.set_content(body)
safe = ssl.create_default_context()

#Cria o Anexo
anexo = "data/Final.xlsx"
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].split("/")

# Adiciona o anexo
with open(anexo, "rb") as a:
    message.add_attachment(
        a.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename=anexo
    )

# Envia o email
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as smtp:
    smtp.login(from_email, password)
    smtp.sendmail (
        from_email,
        to_email,
        message.as_string()
    )

