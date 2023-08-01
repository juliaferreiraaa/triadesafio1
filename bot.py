# Import for the Desktop Bot
from botcity.core import DesktopBot

# Para o Arquivo Excel
import pandas as pd
import os

import zipfile  # Pasta Zip

# Para o envio do E-mail
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header


def main():

    nome_do_arquivo = 'desafiopla.xlsx'
    df = pd.read_excel(nome_do_arquivo)  # Ler planilha


    bot = DesktopBot()
    bot.browse('https://docs.google.com/forms/d/e/1FAIpQLSd5VaVQ6z4zeUHd5Flh_fBi49jIh3GiRfYUASLCOqYhiikt4w/viewform?pli=1')


    for index, row in df.iterrows():
        if (row["Sexo"]) == "feminino":
           if not bot.find( "feminino", matching=0.97, waiting_time=10000):
               not_found("feminino")
           bot.click_relative(-24, 14)

        else:
           if not bot.find( "masculino", matching=0.97, waiting_time=10000):
               not_found("masculino")
           bot.click_relative(-26, 17)

        bot.page_down()

        if not bot.find( "nome", matching=0.97, waiting_time=10000):
            not_found("nome")
        bot.click_relative(30, 76)
        bot.paste(row["Nome"])
    
        if not bot.find( "email", matching=0.97, waiting_time=10000):
            not_found("email")
        bot.click_relative(14, 75)
        bot.paste(row["E-mail"])
    
        if not bot.find( "endereco", matching=0.97, waiting_time=10000):
            not_found("endereco")
        bot.click_relative(43, 77)
        bot.paste(row["Endereço"])
    
        if not bot.find( "numero", matching=0.97, waiting_time=10000):
            not_found("numero")
        bot.click_relative(50, 78)
        bot.paste(row["Telefone"])
        bot.tab()

        # Imprimir a página atual
        bot.control_p()


        if not bot.find( "clicareselecionardestino", matching=0.97, waiting_time=10000):
            not_found("clicareselecionardestino")
        bot.click_relative(286, 22)

        bot.wait(2000)
        if not bot.find( "salvarcomopdf2", matching=0.97, waiting_time=10000):
            not_found("salvarcomopdf2")
        bot.click()


        if not bot.find( "clickemsalvar", matching=0.97, waiting_time=10000):
            not_found("clickemsalvar")
        bot.move()
        bot.click()
        
                    
        if not bot.find( "cliqueparacolarcaminho", matching=0.97, waiting_time=10000):  # Icone se estiver na pasta donwload
            not_found("cliqueparacolarcaminho")
            if not bot.find( "iconealternativopsalva", matching=0.97, waiting_time=10000):  # Icone se estiver em outras pastas
                not_found("iconealternativopsalva")
        bot.click()
        bot.delete()
        bot.delete()
        bot.paste("C:\\Users\julia\PycharmProjects\Desafio1\desafiobot")
        bot.enter()
                    

        if not bot.find( "nomearpdf", matching=0.97, waiting_time=10000):
            not_found("nomearpdf")
        bot.click_relative(81, 9)
        bot.delete()
        bot.paste(row["Nome"])

        bot.tab()
        bot.tab()
        bot.tab()
        bot.enter()

        if not bot.find( "limparformulario", matching=0.97, waiting_time=10000):
            not_found("limparformulario")
        bot.click()

        if not bot.find( "limparformulario2", matching=0.97, waiting_time=10000):
            not_found("limparformulario2")
        bot.click()
        
        if index == 5:
            break

    arquivo_zip = zipfile.ZipFile(r'C:\\Users\julia\PycharmProjects\Desafio1\desafiobot\formularioszip', 'w')
    for pasta, subpastas, arquivos in os.walk(r'C:\\Users\julia\PycharmProjects\Desafio1\desafiobot'):
        for arquivo in arquivos:
            if arquivo.endswith('.pdf'):
                arquivo_zip.write(os.path.join(pasta, arquivo), os.path.relpath(os.path.join(pasta, arquivo), r'C:\\Users\julia\PycharmProjects\Desafio1\desafiobot'), compress_type=zipfile.ZIP_DEFLATED)
    arquivo_zip.close()


    # Cria um objeto MIME Multi-Parts (Nosso E-mail em partes)
    new_email = MIMEMultipart()

    new_email['From'] = 'julia.ferreira@triasoftware.com.br'

    new_email['To'] = 'julia.isabela_ferreira@hotmail.com , julia.ferreira@triasoftware.com.br'

    # Titulo do E-mail
    new_email['Subject'] = 'Desafio 1 - Formulários'

    # Mensagem no corpo do E-mail.
    body = ('Olá,boa tarde.\n'
            'Segue anexo dos formulários em formato PDF conforme solicitado.\n'
            'Atenciosamente,\n'
            'Julia Ferreira ')

    # Attach anexa arquivos no E-mail, nesse caso a variavel body definida na linha de cima, e define o formato do texto.
    # Plain = Texto plano,so texto.
    new_email.attach(MIMEText(body, 'plain'))

    # O caminho do arquivo que queremos enviar.
    filepath = r'C:\\Users\julia\PycharmProjects\Desafio1\desafiobot\formularioszip'

    # Guarda a instancia aberta do arquivo para a leitura, para depois enviar o E-mail.
    # (r - READ) (b - BINARY)
    attachment = open(filepath, 'rb')

    # Base do formato para conseguir ler e depois anexar.
    part = MIMEBase('application', 'octet-stream')

    # Configura o arquivo que abrimos nas linhas anteriores, para ser inserido no E-mail usando a função read.
    part.set_payload(attachment.read())

    # Determina o formato da plataforma que estamos usando para fazer isso, no caso a base do windows é 64BITS.
    # Dentro da função colocarmos o PART que tem a configuração que precisamos que definimos anteriormente.
    encoders.encode_base64(part)

    """ Importante passar no header da variavel filename o formato UTF-8 ,porque isso define o tipo de base de caracteres 
    que vamos usar e permite mudar o nome do arquivo anexado e para nao enviar em formato binário que mexemos anteriormente """

    part.add_header('Content-Disposition', "attachment", filename=(Header('formularioszip.zip', 'utf-8').encode()))

    # Inserimos toda a configuração feita acima dentro do E-mail.
    new_email.attach(part)

    # Iniciamos uma instancia do servidor que vai conectar nosso codigo com o app outlook.
    # Indicamos a porta que vai ser usada no modem.
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)

    # Inicia o metodo de segurança TLS.
    server.starttls()

    server.login('julia.ferreira@triasoftware.com.br', 'Juju2012!')

    # Variavel que vai receber todo o E-mail configurado acima, como string.
    text = new_email.as_string()

    # Envia o E-mail, (Quem envia, Quem recebe, Conteudo)
    destinatarios = ['julia.ferreira@triasoftware.com.br', 'julia.isabela_ferreira@hotmail.com']
    server.sendmail('julia.ferreira@triasoftware.com.br', destinatarios, text)

    # Fecha a conexão do servidor.
    server.quit()

    print('e-mail enviado')

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()