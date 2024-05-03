# MeuPrimeiroCodigoPython
Automatização de Envio de E-mails e Relatório em PDF com Outlook - Este script Python automatiza o envio de e-mails pelo Outlook com base em uma lista de convidados em um arquivo Excel. Além disso, ao final do envio, é gerado um relatório em PDF indicando os nomes dos destinatários e se o e-mail foi enviado com sucesso.

Funcionalidades

Envia e-mails individualizados para cada destinatário com base em um modelo HTML.

Anexa imagens aos e-mails.

Gera um relatório em PDF com o status do envio de e-mails.

O relatório deve conter a logomarca da sua empresa (logo)

Como Utilizar

Instale as bibliotecas necessárias executando pip install pandas, pywin32, openpyxl, reportlab, fpdf e bs4.

Certifique-se de ter o Outlook instalado e configurado.

Coloque os arquivos Lista_convidados.xlsx, Convite.jpg, Assinatura_do_e-mail.jpg e a logo.png na mesma pasta do script.

Execute o script envio_emails.py.

Realize testes com uma lista de e-mails de forma a se certificar que está tudo conforme deseja, caso contrário realize os ajustes necessários.

Verifique a saída no console para confirmar o envio dos e-mails.

O relatório em PDF será gerado como relatorio_envio_emails.pdf e será salvo na mesma pasta do script.

Certifique-se de ter permissões necessárias para acessar e-mails e executar o script.
