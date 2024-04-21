import win32com.client as client
import pandas as pd
import datetime as dt


# Carregar o arquivo Excel
tabela = pd.read_excel('Exemplo Contas a Receber.xlsx')

hoje = dt.datetime.now()


#Coleta só os dados daqueles com Status em aberto
tabela_devedores = tabela.loc[tabela['Status']=='Em aberto']
tabela_devedores = tabela_devedores.loc[tabela_devedores['Data Prevista para pagamento']<hoje]
print(tabela_devedores)


#Cria lista com os valores da tabela dos devedores que vão ser usadas no E-mail
infos = tabela_devedores[['Valor em aberto','Data Prevista para pagamento','E-mail', 'NF']].values.tolist()


#Encaminhando as Mensagens para os Clientes devedores
for info in infos:
    destinatario = info[2]
    nf = info[3]
    prazo = info[1]
    prazo = prazo.strftime("%d/%m/%Y")
    valor = info[0]

    envmail = client.Dispatch('Outlook.Application')
    emissor = envmail.session.Accounts['seuemail@dominio.com']
    msg = envmail.CreateItem(0)
    #msg.Display() #Abre o e-mail (se tirar essa linhas ele manda direto em segundo plano)
    msg.To = destinatario
    msg.Subject = 'Identificamos um Atraso no Pagamento'
    msg.Body = f'''
        Prezado Cliente, 
    Identificamos que você ainda não realizou o pagamento da NF({nf}) no valor de R${valor:.2f} com vencimento em {prazo}. Se estiver com dificuldades para fazer o pagamento em nossa plataforma favor entre em contato.

        Atenciosamente,
        Felipe
    '''
    msg._oleobj_.Invoke(*(64209,0,8,0,emissor))
    msg.Save()
    msg.Send()