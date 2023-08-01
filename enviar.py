import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header



def enviar_email() -> object:
# Cria um objeto MIME Multi-Parts (Nosso E-mail em partes)
    new_email = MIMEMultipart()
    new_email['From'] = 'julia.ferreira@triasoftware.com.br'
    new_email['To'] = 'julia.isabela_ferreira@hotmail.com , julia.ferreira@triasoftware.com.br'  #Mudar para francir.silverio@triasoftware.com.br

    # Titulo do E-mail
    new_email['Subject'] = 'Desafio 2 - Consumo de API '

    # Mensagem no corpo do E-mail.
    body = ('Olá,boa tarde.\n'
            'Segue anexo dos formulários em formato PDF conforme solicitado.\n'
            'Atenciosamente,\n'
            'Julia Ferreira ')

    # Attach anexa arquivos no E-mail, nesse caso a variavel body definida na linha de cima, e define o formato do texto.
    # Plain = Texto plano,so texto.
    new_email.attach(MIMEText(body, 'plain'))

    # O caminho do arquivo que queremos enviar.
    filepath = r"C:\Users\julia\PycharmProjects\botcep\botdesafio\formularioszip"

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
    destinatarios = ['julia.ferreira@triasoftware.com.br', 'julia.isabela_ferreira@hotmail.com'] #Mudar para francir.silverio@triasoftware.com.br
    server.sendmail('julia.ferreira@triasoftware.com.br', destinatarios, text)

    # Fecha a conexão do servidor.
    server.quit()

    print('e-mail enviado')