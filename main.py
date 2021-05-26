import pandas as pd
import win32com.client as win32


#baixar pandas
#baixar o pywin32

#   IMPORTAR A BASE DE DADOS

tabela_vendas = pd.read_excel('Vendas.xlsx')



#   VISUALIZAR A BASE DE DADOS PARA CONFERENCIA


pd.set_option('display.max_columns', None)


#aqui temos que passar duas coisas, a opcao q vc quer e o valor para aquela opcao.
# display.max_columns = vemos que ele fala num valor máximo de colunas. Em quantidade pedido para q n faça isso.
# Ou seja, nao existirá valor máximo para colunas. Quantas tiverem, serao mostradas.

#   FATURAMENTO DA LOJA

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# no formato 1 tabela_vendas [['ID Loja', 'Valor Final']], ele tá pegando apenas essas colunas de todas que
# existem. no formato 2 grouby('ID Loja').sum() ele agrupa o id das lojas e soma o resto.
# como separamos o id da loja e o valor final, entao o sum, somará o valor final.
# print(faturamento) vc irá visualizar todas as lojas e seus faturamentos


#   QUANTIDADE DE PRODUTOS VENDIDOS POR LOJA

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()


# print(quantidade) vc irá visualizar a quantidade de produtos vendidos por loja


#   TICKET MÉDIO POR PRODUTO EM CADA LOJA

ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
# um colchete pq quero passar apenas uma coluna.
# uma lista de colunas, dois colchetes.
# Sempre que vc faz um equacao de uma coluna com outra coluna e vc quer que ela vire uma tabela
# vc tem que colocar o .to_frame()
#Se vc perceber, usei a mudança do nome da coluna pq ele estava indo, antes da modificaçao, como coluna 0.


#   ENVIAR UM EMAIL COM O RELÁTÓRIO COLHIDO ACIMA


outlook = win32.Dispatch('outlook.application')
#Se conectando com o app do outlook
mail = outlook.CreateItem(0)
#Criar um novo email
mail.To = 'enviarpraesseemail@hotmail.com'
#Para qual endereço de email
mail.Subject = 'Relatório Geral'
#Assunto do Email
#mail.Body = 'Message body'
#Corpo do email!  Mas neste caso, usaremos o html body.
mail.HTMLBody = f'''
<p>Prezado,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format })}

<p>Quantidade vendida:</p>
{quantidade.to_html()}

<p>Ticket médio dos produtos de cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format })}
<p>Qualquer dúvida, estarei à disposição</p>

<p>Att,</p>

<p>Jefferson Andrade.</p>
'''
#COM TRES ASPAS SIMPLES VC MOSTRA AO PYTHON QUE VAI ESCREVER MAIS DO QUE APENAS UMA LINHA
# To attach a file to the email (optional):
#attachment  = "Path to the attachment"
#mail.Attachments.Add(attachment)

mail.Send()

print('Email Enviado com sucesso!!')