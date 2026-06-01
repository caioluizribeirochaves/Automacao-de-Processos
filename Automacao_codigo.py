# 1 - Importar Arquivos e Bibliotecas

import pandas as pd
from IPython.display import display
import pathlib


# Para ler os arquivos presentes na pasta, como os emails, lojas e vendas

emails = pd.read_excel(r'Bases de Dados/Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados/Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)

# 2 - Criar uma tabela para cada loja e definir o dia do indicador

# atribui um ID a cada loja

vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)

# Criar um dicionário com todas as lojas e seus dados

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]

display(dicionario_lojas['Iguatemi Esplanada'])

# Localizar o dia indicador, que nesse caso será a data mais recente, por isso é utilizado o .max()

dia_indicador = vendas['Data'].max()
print(dia_indicador)
print(f'{dia_indicador.day}/{dia_indicador.month}')

# 3 - Salvar na Pasta de Backup

# Identificar se a pasta de cada loja já existe e caso ainda não tenha, ela será criada

caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pastas_backup = caminho_backup.iterdir()
lista_backup_lojas = [arquivo.name for arquivo in arquivos_pastas_backup]

for loja in dicionario_lojas:
    if loja not in lista_backup_lojas:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

# Para poder salvar o arquivo dentro das respectivas pastas de backup

    nome_do_arquivo = f'{dia_indicador.day}_{dia_indicador.month}_{loja}.xlsx'
    local_do_arquivo = caminho_backup / loja / nome_do_arquivo
    dicionario_lojas[loja].to_excel(local_do_arquivo)
# 4 - Calcular o indicador para loja

# Calcular o faturamento do dia e o faturamento total de cada loja

loja = 'Norte Shopping'
vendas_loja = dicionario_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

# Faturamento

faturamento_ano = vendas_loja['Valor Final'].sum()
print(faturamento_ano)
faturamento_dia = vendas_loja_dia['Valor Final'].sum()
print(faturamento_dia)

# Diversidade de produtos: vê a quantidade de produtos.

quantidade_prod_ano = len(vendas_loja['Produto'].unique())
print(quantidade_prod_ano)
quantidade_prod_dia = len(vendas_loja_dia['Produto'].unique())
print(quantidade_prod_dia)

# Ticket Médio

# valor para o ano

valor_vendas = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
ticket_medio_ano = valor_vendas['Valor Final'].mean()
print(ticket_medio_ano)

# valor para o dia

valor_vendas_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
ticket_medio_dia = valor_vendas_dia['Valor Final'].mean()
print(ticket_medio_dia)

# Definição das metas

meta_faturamento_dia = 1000
meta_faturamento_ano = 16500000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

# 5 - Disparar os e - mails

