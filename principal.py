import openpyxl 
from urllib.parse import quote
import webbrowser
from time import sleep

webbrowser.open('https://web.whatsapp.com/')
sleep(15)  

workbook = openpyxl.load_workbook('Cópia de clientes.xlsx')
paginas_clientes = workbook['Sheet1']

for linha in paginas_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    if not nome or not telefone or not vencimento:
        continue  # pula se faltar algum dado

    # Formatação
    telefone = str(telefone).strip()
    if not telefone.startswith('+'):
        telefone = f'+55{telefone}'  # Adiciona código do país

    mensagem = f'Olá {nome}, o seu boleto vence na data {vencimento.strftime("%d/%m/%Y")}!'
    print(f'Enviando para {nome} ({telefone}): {mensagem}')

    link_mensagem_wpp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_wpp)
    sleep(10)  # Tempo para envio manual ou leitura
