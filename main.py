import smtplib
import config

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

class myMail():

    def cabecalho(self, consultor: str, destino: str, remetente):

        global msg

        # cabecalho do email:
        msg = MIMEMultipart('alternative')
        msg['From'] = remetente  # user
        match destino:
            case 'consultores':
                msg['Subject'] = f'Relatório {consultor} (período: {start} - {end})'
            case 'diretoria':
                msg['Subject'] = f'Relatório Estados (período: {start} - {end})'

        # imagem da Header:
        with open(r'C:\Users\kevin.krause\Desktop\relatorios_bi_por_email\performance_vendas_crm\Header.jpg', 'rb') as header:
            # define o tipo de arquivo, tipo de img (png/jpg):
            mimeh = MIMEBase(
                'image', 'jpg', filename='Header.jpg')
            # adiciona dados da Header:
            mimeh.add_header('Content-Disposition',
                             'attachment', filename='img1.jpg')
            mimeh.add_header('X-Attachment-Id', 'header')
            mimeh.add_header('Content-ID', '<header>')
            # insere dados dentro do Mime object:
            mimeh.set_payload(header.read())
            # encode com base64:
            encoders.encode_base64(mimeh)
            # adciona o mimeobject no MIMEMultipart object:
            msg.attach(mimeh)

            return msg
        

            part = MIMEBase('application', "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', 'attachment; filename="WorkBook3.xlsx"')
            message.attach(part)