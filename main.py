import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None)
print(tabela_vendas)

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

def setup():
    email_user = 'seu_email@gmail.com'
    email_pass = 'sua senha que será gerada no gmail'
    server_smtp = 'smtp.gmail.com'
    server_port = 465
    timeout = 10.0

    return email_user, email_pass, server_smtp, server_port, timeout

def setup_and_send(receiver_email='', subject_line='', message_body=''):
    you = receiver_email
    texto_padrao =  f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Seu nome</p>
'''
    message = texto_padrao
    subject = subject_line
    text = message_body

    email_user, email_pass, server_smtp, server_port, timeout = setup()

    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = email_user
    msg['To'] = you

    html = f'''<strong>{ message } </strong>'''
    plain_version = MIMEText(text, 'plain')
    html_version = MIMEText(html, 'html')

    #msg.attach(plain_version)
    msg.attach(html_version)

    with smtplib.SMTP_SSL(server_smtp, server_port, timeout=timeout) as server:
        server.set_debuglevel(0)
        to_addrs = [you]
        server.login(email_user, email_pass)
        server.sendmail(email_user, to_addrs, msg.as_string())

def main(receiver_email='email_destinatario@gmail.com', subject_line='Relatório', message_body=''):
    if receiver_email:
        setup_and_send(receiver_email, subject_line, message_body)
    else:
        print('receiver_email is False.')

main()


