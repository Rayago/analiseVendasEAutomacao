# -*- coding: utf-8 -*-
"""
Created on Tue Feb  9 07:43:21 2021

@author: Rumpel
"""

import pandas as pd

df = pd.read_excel(r'Vendas.xlsx')

faturamento = df[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

faturamento = faturamento.sort_values(by='Valor Final', ascending=False)

# display(faturamento)

quantidade = df[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
quantidade = quantidade.sort_values(by='ID Loja', ascending=False)

# display(quantidade)

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
ticket_medio = ticket_medio.sort_values(by='Ticket Medio', ascending=False)
display(ticket_medio)

# Função enviar_email
import smtplib
import email.message

def enviar_email(resumo_loja, loja):
    
    server = smtplib.SMTP('smtp.gmail.com:587')  
    corpo_email = f'''
    <p>Segue a tabela de faturamento:</p>
    {resumo_loja.to_html()}
    <p>Grande abraço!</p>'''
      
    msg = email.message.Message()
    msg['Subject'] = f'Loja: {loja}'
      
    # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
    # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534)
    # => Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
    msg['From'] = 'seuEmailAqui@gmail.com'
    msg['To'] = 'emailDoDestinatarioAqui@gmail.com'
    password = "suaSenhaAqui"
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)
      
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')
    
lojas = df['ID Loja'].unique()

for loja in lojas:
    tabela_loja = df.loc[df['ID Loja'] == loja, ['ID Loja', 'Quantidade', 'Valor Final']]
    resumo_loja = tabela_loja.groupby('ID Loja').sum()
    resumo_loja['Ticket Médio'] = resumo_loja['Valor Final'] / resumo_loja['Quantidade']
    enviar_email(resumo_loja, loja)
    
# E-mail para a diretoria
tabela_diretoria = faturamento.join(quantidade).join(ticket_medio)
enviar_email(tabela_diretoria, 'Todas as Lojas')