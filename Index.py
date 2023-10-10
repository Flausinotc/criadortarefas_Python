import speech_recognition as sr
import openpyxl
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configurações do servidor de e-mail (Gmail neste exemplo)
email_host = '###'
email_port = 587
email_user = '###'
email_password = '###'

def transcrever_audio():
    gravar = sr.Recognizer()
    with sr.Microphone() as source:
        print("Fale p/ criar um novo evento:")
        audio = gravar.listen(source)

    try:
        texto = gravar.recognize_google(audio, language='pt-BR')
        return texto
    except sr.UnknownValueError:
        print("Não consegui entender")
        return None
    except sr.RequestError as e:
        print("Erro no serviço de reconhecimento de fala; {0}".format(e))
        return None

def escrever_no_excel(texto):
    arquivo_excel = "C:/Users/thiago.cardoso/Documents/Project_python/Agendador de tarefas/transcricoes.xlsx"
    try:
        workbook = openpyxl.load_workbook(arquivo_excel)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(['Data e Hora', 'Texto Transcrito'])

    data_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sheet.append([data_hora, texto])

    workbook.save(arquivo_excel)

    print(f'Transcrição "{texto}" gravada em: {data_hora}')
    
    # Envio de e-mail
    enviar_email(texto, data_hora)

def enviar_email(texto, data_hora):
    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = '###'  # Substitua pelo endereço de e-mail do destinatário
    msg['Subject'] = 'Novo evento foi criado!'

    corpo_email = f'Um novo evento foi criado: {data_hora}:\n\n{texto}'
    msg.attach(MIMEText(corpo_email, 'plain'))

    try:
        server = smtplib.SMTP(email_host, email_port)
        server.starttls()
        server.login(email_user, email_password)
        text = msg.as_string()
        server.sendmail(email_user, msg['To'], text)
        server.quit()
        print("E-mail enviado com sucesso!")
    except smtplib.SMTPException as e:
        print("Erro ao enviar e-mail:", str(e))

texto_transcrito = transcrever_audio()
if texto_transcrito:
    escrever_no_excel(texto_transcrito)
