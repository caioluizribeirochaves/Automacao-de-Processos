# Automação de Processos: Filtragem de Dados e Relatórios por E-mail

## 🎯 Objetivo do Projeto
O objetivo deste projeto é automatizar o processamento de uma base de dados unificada de Shoppings de diversas regiões do Brasil e realizar o envio direcionado de relatórios de desempenho para gerentes e diretores.

A base de dados original contém as seguintes informações:
* **Identificação:** Nome do Shopping e e-mails dos representantes legais.
* **Vendas:** Histórico diário de produtos vendidos, quantidade e preço unitário.

---

## 🛠️ Fluxo de Processamento dos Dados

Para garantir a organização e a escalabilidade do código, o processo foi dividido em quatro etapas principais:

### 1. Tratamento e Segregação de Dados
O sistema atribui um ID exclusivo para cada Shopping, permitindo isolar as transações de forma individual. Em seguida, os dados são estruturados em dicionários específicos para cada unidade, garantindo que as informações não se misturem.

### 2. Geração de Backup Automatizado
Para cada loja, o script aplica um filtro baseado na **data mais recente** de vendas e gera um arquivo em Excel (`.xlsx`). Esses arquivos são salvos automaticamente em pastas exclusivas, nomeadas de acordo com o Shopping correspondente.

### 3. Disparo de E-mails para os Representantes
Cada gestor recebe um e-mail personalizado contendo um panorama geral da sua unidade e o relatório detalhado em anexo. No corpo do e-mail, são destacados os seguintes indicadores financeiros:
* **Faturamento** (Diário e Anual)
* **Quantidade de Produtos Vendidos** (No dia e no ano)
* **Ticket Médio** (Diário e Anual)

### 4. Relatório Consolidado para a Diretoria
Por fim, o sistema gera e envia um e-mail exclusivo para a diretoria executiva contendo um **ranking geral de faturamento** entre todos os Shoppings, facilitando a tomada de decisão estratégica baseada na performance diária e anual de cada unidade.

---

## 🚀 Proposta de Valor

> **Frequência:** Diária  
> **Impacto:** Otimizar o controle sobre as vendas de cada unidade, eliminar o trabalho manual de filtragem/envio e garantir que a diretoria e os representantes tenham dados precisos logo no início do dia.
