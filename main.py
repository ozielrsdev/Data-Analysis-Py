import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#Importar DB
sales_table = pd.read_excel("Vendas.xlsx")
#Ver base de dados
pd.set_option("display.max_columns", None)

#Faturamento por Loja
billing = sales_table[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()

#Quantidade de produtos vendidos por loja
sold_products_by_shop = sales_table[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()


#Ticket medio por produto em cada loja

average_ticket = (billing["Valor Final"] / sold_products_by_shop["Quantidade"]).to_frame()
average_ticket = average_ticket.rename(columns={0: "Ticket Médio"})


#Enviar um relatório por e-mail
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "email_do_remetente"
password = "senha_do_email"

# Criando a mensagem
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = "email_do_destinatáro"
message["Subject"] = "Relatório de Vendas"

html = f''' 
<p>Olá!</p>

<p>Segue abaixo o relatório de vendas de cada loja</p>

<p>Faturamento:</p>
{billing.to_html(formatters={"Valor Final" : "R${:,.2f}".format})}

<p>Quantidade Vendida:</p>
{sold_products_by_shop.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{average_ticket.to_html(formatters={"Ticket Médio": "R${:,.2f}".format})}

<p>Qualquer dúvida estou à disposição</p>

<p>Att.,</p>
<p>Oziel Sousa</p>
'''

message.attach(MIMEText(html, "html"))

# Enviando o e-mail
with smtplib.SMTP(smtp_server, smtp_port) as server:
    server.starttls()  # Conexão segura via TLS
    server.login(sender_email, password)  # Fazendo login no servidor SMTP
    server.sendmail(sender_email, message["To"], message.as_string())

print("E-mail enviado com sucesso!")