# Agendador de Tarefas por Reconhecimento de Fala

Este é um projeto em Python que permite criar eventos de agendamento de tarefas por meio de reconhecimento de fala, armazenar as transcrições em uma planilha Excel e enviar notificações por e-mail sobre os eventos criados. O código utiliza as bibliotecas `speech_recognition`, `openpyxl`, `datetime`, e `smtplib` para realizar essas tarefas.

## Configuração

Antes de usar o código, é necessário configurar as informações do servidor de e-mail, como o host, a porta, o nome de usuário e a senha. Isso é feito nas seguintes variáveis no código-fonte:

```python
email_host = '###'          # Host do servidor de e-mail (por exemplo, 'smtp.gmail.com')
email_port = 587             # Porta do servidor de e-mail (por exemplo, 587 para TLS no Gmail)
email_user = '###'           # Nome de usuário do e-mail
email_password = '###'       # Senha do e-mail

Uso
Execute o código Python.
O programa gravará sua fala quando você falar "Fale para criar um novo evento:".
Após falar, o texto reconhecido será transcrito e armazenado em uma planilha Excel.
Um e-mail será enviado para o destinatário especificado com as informações do evento.
# Execução Principal
texto_transcrito = transcrever_audio()
if texto_transcrito:
    escrever_no_excel(texto_transcrito)

Requisitos
Certifique-se de ter as seguintes bibliotecas Python instaladas:

speech_recognition
openpyxl
Notas
Certifique-se de substituir as variáveis de configuração (email_host, email_port, email_user, email_password, e o endereço de e-mail do destinatário) com suas informações pessoais antes de usar o código.
Para usar serviços de e-mail, verifique as políticas de segurança do seu provedor de e-mail, como permitir o acesso de aplicativos menos seguros, se necessário, para que o envio de e-mails funcione corretamente.


Você pode copiar o texto acima e colá-lo diretamente em seu repositório do GitHub para documentar o seu código. Certifique-se de substituir `"###"` com as suas informações reais de configuração de e-mail e fazer outras personalizações conforme necessário.
