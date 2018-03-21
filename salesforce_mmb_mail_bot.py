# Importa dependencias
import email, getpass, imaplib, os


#Variaveis que serao usadas durante o processo
detach_dir = 'INSIRA O ARQUIVO AQUI\.' # Diretorio onde sera salvo o anexo (default: current)
user = "INSIRA O EMAIL EXCHANGE AQUI" # Endereco de e-mail onde o anexo esta
pwd = "INSIERA O PASSWORD AQUI" # Senha do e-mail inserido


# Conectando no servidor IMAP do exchange
m = imaplib.IMAP4_SSL("outlook.office365.com") #Endereco do servidor IMAP
m.login(user,pwd) # Usuario e senha 
m.select("Emails_Salesforce") # Pasta do e-mail onde você quer procurar o email


#Filtra o e-mail em questao
resp, items = m.search(None, "ALL") # A palavra ALL é um filtro especifico IMAP 
# (veja mais filtros em http://www.example-code.com/csharp/imap-search-critera.asp)
items = items[0].split() # Pega o ID dos e-mails
for item in items:
    print(item)


# Pega dados do primeiro e-mail
resp, data = m.fetch(items[-1], "(RFC822)")
email_body = data[0][1] # Pega conteudo do e-mail
mail = email.message_from_bytes(email_body) # parsing da mensagem de e-mail
print("["+mail["From"]+"] :" + mail["Subject"])


# Metodo WALK() utilizado para passar pelas diferentes partes do e-mail
for part in mail.walk():
    # multiparts sao containers, entao podemos passar direto por eles
    if part.get_content_maintype() == 'multipart':
        continue

    # Testa se a parte lida e um anexo
    if part.get('Content-Disposition') is None:
        continue

    # Da o nome do arquivo que sera gravado na pasta indicada no comeco do script
    # filename = part.get_filename() # Deixe essa linha caso queira manter o nome original do arquivo
    filename = "download.xlsx"
    att_path = os.path.join(detach_dir, filename)

    #Verifica se o arquivo ja nao existe
    if not os.path.isfile(att_path) :
        # Escreve o arquivo
        fp = open(att_path, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()