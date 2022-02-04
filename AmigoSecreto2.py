import pandas as pd
import random
import smtplib
import os
from collections import OrderedDict
from email.message import EmailMessage
from DadosEnviarEmails import senha, email

os.system('cls')

excel = pd.read_excel (r"C:\Users\thica\OneDrive\Área de Trabalho\AmigoSecreto.xlsx", engine='openpyxl')

numerosSorteados = []
lenght_participantes = 0
lenght_participantes = len(excel ['Nomes'])
randomizacao = 0

cont  = 0
cont2 = 0
cont3 = 0

OrderedDict.fromkeys(numerosSorteados)


def parte_inicio ():
    global cont, numerosSorteados

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

        print(participante, ' tirou: ', amigoSecretoParticipante, ' (email ', amigoSecretoParticipante,')\n')

        numerosSorteados.append(randomizacao)

        #listaFinalNomesSorteados  = list(OrderedDict.fromkeys(numerosSorteados))
        #print("Esse é o listFinalNomesSorteados: ", listaFinalNomesSorteados)

        cont += 1

parte_inicio()

lenght_participantes = len(excel ['Nomes'])

def parte_fim ():
    global cont3, numerosSorteados, lenght_participantes
        
    print(cont3)

    while(cont3 < lenght_participantes):
        EMAIL_ADRESS = email
        EMAIL_PASSWORD = senha
            
        msg = EmailMessage()
        msg['Subject'] = "Amigo secreto da Família Canato!"
        msg['From'] = EMAIL_ADRESS
        msg['To'] = excel['Emails'][cont3]
        msg.set_content("Você, infelizmente, tirou: " +  excel['Nomes'][numerosSorteados[cont3]] + '\n' + 'Essa pessoa gostaria de receber: ' + excel['Pedidos'][numerosSorteados[cont3]])
        
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(EMAIL_ADRESS, EMAIL_PASSWORD)
            smtp.send_message(msg)
            print("Email enviado")
                
        cont3+=1

#parte_fim()