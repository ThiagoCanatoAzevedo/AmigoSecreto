import pandas as pd
import random
import smtplib
import os
from collections import OrderedDict
from email.message import EmailMessage

os.system('cls')

excel = pd.read_excel (r"CAMINHO DO SEU EXCEL COM NOMES DOS PARTICIPANTES", engine='openpyxl' )
numerosSorteados     = []

lenght_participantes = 0
lenght_participantes = len(excel ['Nomes'])
randomizacao = 0
cont  = 0
cont2 = 0

OrderedDict.fromkeys(numerosSorteados)

while(cont < lenght_participantes):

    participante = excel['Nomes'][cont]

    randomizacao = random.randint(0,lenght_participantes-1)
    
    cont2 = 0
    while( cont2 < 3 and (randomizacao in numerosSorteados or excel['Família'][cont] == excel['Família'][randomizacao])):
        randomizacao = random.randint(0,lenght_participantes-1)
        cont2 += 1

    if(cont2 == 3):
        numerosSorteados = []
        cont = 0
        os.system('cls')
        continue

    familiaParticipante           = excel['Família'][cont]
    amigoSecretoParticipante      = excel['Nomes'][randomizacao]
    amigoSecretoParticipanteEmail = excel['Emails'][randomizacao]

  
    numerosSorteados.append(randomizacao)

    listaFinalNomesSorteados = list(OrderedDict.fromkeys(numerosSorteados))


    cont += 1

print(listaFinalNomesSorteados) #-----> Números Sorteados 


quantidadeNomesExcel = 0
quantidadeNomesExcel = len(excel['Nomes'])


cont3 = 0
while(cont3 < quantidadeNomesExcel):
    EMAIL_ADRESS = SEU EMAIL
    EMAIL_PASSWORD = SUA SENHA
    
    msg = EmailMessage()
    msg['Subject'] = "Amigo secreto da Família Canato!"
    msg['From'] = EMAIL_ADRESS
    msg['To'] = excel['Emails'][cont3]
    msg.set_content("Você, infelizmente, tirou: " +  excel['Nomes'][listaFinalNomesSorteados[cont3]] + '\n' + 'Essa pessoa gostaria de receber: ' + excel['Pedidos'][listaFinalNomesSorteados[cont3]])
    
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
        print("Email enviado")
        
    cont3+=1

print('Todos email-s foram enviados!')
