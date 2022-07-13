#instalar antes a bliblioteca pip install pywin32
import win32com.client as win32

#criar a integração com o outlook
outlook =  win32.dispatch('outlook.application')

#criar um email
email= outlook.Creatitem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

#configurar as informações do seu e-mail
email.To = "destino; destino2"
email.Subject = "Email automatico do python"
email.HTMLBody = f"""
<p> Ola Aldair,aqui e o codigo python </p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
"""

#anexo = "colocar o caminho do anexo aqui exemplo :C://Users/aldair/Downloads/arquivo.xlsx "
#email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")

