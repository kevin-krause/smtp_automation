import smtplib
import config
import locale

from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from datetime import datetime
from datetime import timedelta

locale.setlocale(locale.LC_ALL, 'pt_br.UTF-8')
today = datetime.today()
first = today.replace(day=1)
dt_comissao = first - timedelta(days=1)
mes = dt_comissao.strftime('%B')
ano = dt_comissao.strftime('%Y')

remetente = config.usuario


def cabecalho():

    global msg

    msg = MIMEMultipart('alternative')
    msg['From'] = remetente
    msg['Subject'] = f'Comissões Rental {mes}'

    # imagem da Header:
    with open(r'C:\Users\kevin.krause\Desktop\relatorios_bi_por_email\performance_vendas_crm\Header.jpg', 'rb') as header:
        mimeh = MIMEBase(
            'image', 'jpg', filename='Header.jpg')
        mimeh.add_header('Content-Disposition',
                         'attachment', filename='img1.jpg')
        mimeh.add_header('X-Attachment-Id', 'header')
        mimeh.add_header('Content-ID', '<header>')
        mimeh.set_payload(header.read())
        encoders.encode_base64(mimeh)
        msg.attach(mimeh)

        return msg


def anexarArquivo(filepath: str):
    with open(filepath, 'rb') as arquivo:
        mimea = MIMEBase(
            'excel', 'xlsx', filename='Comissões Consultores.xlsx')
        mimea.add_header('Content-Disposition',
                         'attachment', filename='Comissões Consultores.xlsx')
        mimea.set_payload(arquivo.read())
        encoders.encode_base64(mimea)
        msg.attach(mimea)

        return msg


def body(index: str):
    index = open(index).read().format(mes=mes, ano=ano)
    body = MIMEText(index, 'html')
    msg.attach(body)

    return msg


def sendmail(destinatario: str):

    msg['To'] = destinatario  # variavel email do consultor

    s = smtplib.SMTP(host='zmail.grupomotormac.com.br', port=587)
    s.starttls()
    s.login(config.usuario, config.senha)
    s.ehlo()

    s.sendmail(remetente, destinatario, msg.as_string())
    print(f"Email Enviado {today}")

    s.quit()


cabecalho()
anexarArquivo(
    filepath=r"C:\Users\kevin.krause\Downloads\Modelo Comissoes.xlsx")
body(index=r'C:\Users\kevin.krause\Desktop\smtp_automation\index.html')
sendmail(destinatario='kevin.krause@motormac.com.br')
