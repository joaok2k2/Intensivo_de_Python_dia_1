import pandas as pd
import smtplib
import email.message



def enviar_email(nome_da_loja, tabela):
    corpo_email = f'''
    <p>Prezados, </p>

    <p>Segue relatório de vendas</p>
     {tabela.to_html()}
     <p>Qualquer dúvida estou à disposição.<p/>

     '''

    msg = email.message.Message()
    msg['Subject'] = f"Relatório de Vendas - {nome_da_loja}"

    # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
    # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
    # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
    # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
    msg['From'] = 'remetente@gmail.com'
    msg['To'] = 'destinatario@gmail.com'
    password = "senha"
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


#Passo 1 - Importar a base de dados

tabela_vendas = pd.read_excel("Vendas.xlsx")


#Passo 2 - Calcular o faturamento de cada loja


# Colunas filtradas e fazendo os agrupamentos, com a saida para somar.
tabela_faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()

# Ordenando em ordem crescente por Valor final
tabela_faturamento = tabela_faturamento.sort_values(by="Valor Final", ascending=True)


#Passo 3 - Calcular a quantidade de produtos vendidos de cada loja

# Filtrando por quantidade de vendas
tabela_quantidade = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()




#Passo 4 - Calculcar o tciker médio dos produtos de cada loja

# Calculando o ticket medio e convertenando a saida para tabela, para um data frame.
ticket_medio = (tabela_faturamento["Valor Final"] / tabela_quantidade["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})


#Passo 5 - enviar o e-mail para diretória.

# Juntando as tabelas

tabela_completa = tabela_faturamento.join(tabela_quantidade).join(ticket_medio)

enviar_email("Diretoria", tabela_completa)




#Passo 6 - enviar o e-mail para cada loja

#Criando um lista com todos os nomes das lojas.
#Pegando da base de dados, filtrando por coluna "Id Loja", e usando o método unique, para remover duplicadas.
lista_lojas = tabela_vendas["ID Loja"].unique()


#Crindo um loop para enviar os e-mails de cada loja com seu devido processamento de tabela
#Como por exemplo, a quantidade vendida, o valor final das vendas e o ticket médio.

for loja in lista_lojas:

    #Filtrando linhas e colunas de uma loja, linhas referente a lojas e colunas referente a quantidade e valor final"
    tabela_loja = tabela_vendas.loc[tabela_vendas["ID Loja"] == loja, ["ID Loja", "Quantidade", "Valor Final"]]

    #Agrupando a tabela e calculando o valor final
    tabela_loja = tabela_loja.groupby("ID Loja").sum()

    #Criando a coluna ticket_medio à tabela_loja.
    tabela_loja["Ticket Médio"] = tabela_loja["Valor Final"] / tabela_loja["Quantidade"]

    #Enviando e-mail
    enviar_email(loja, tabela_loja)
