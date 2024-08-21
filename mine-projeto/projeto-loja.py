import pandas as pd
import win32com.client as win32

# Carregar a tabela de vendas
tabela = pd.read_excel('Vendas.xlsx')

# Configurar para visualizar todas as colunas
pd.set_option('display.max_columns', None)

# Calcular o faturamento por loja
faturamento = tabela[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Calcular a quantidade vendida por loja
quantidade_vendida = tabela[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_vendida)
print('-' * 50)

# Calcular o ticket médio por loja
ticket_medio = (faturamento['Valor Final'] / quantidade_vendida['Quantidade']).to_frame()  # Convertendo para DataFrame
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
print('-' * 50)

# Agora vamos enviar o e-mail
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gamesadrg@gmail.com'  # Destinatário
mail.Subject = 'Relatório de Vendas por Loja'

# Convertendo os DataFrames para HTML para o corpo do e-mail
faturamento_html = faturamento.to_html()
quantidade_vendida_html = quantidade_vendida.to_html()
ticket_medio_html = ticket_medio.to_html()

# Corpo do e-mail em HTML
mail.HTMLBody = f'''
<p>Seres, estou encaminhando o relatório de vendas:</p>

<h3>Faturamento:</h3>
{faturamento_html}

<h3>Quantidade Vendida:</h3>
{quantidade_vendida_html}

<h3>Ticket Médio por Loja:</h3>
{ticket_medio_html}

<p>Qualquer dúvida, se foda.</p>
'''

# Enviar o e-mail
mail.Send()

print('E-mail enviado com sucesso!')
