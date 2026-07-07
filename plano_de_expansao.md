# Plano de Expansão: Dashboard de Histórico e Estatísticas

## 1. O Valor Comercial desta Funcionalidade

Esta ideia eleva o sistema para outro patamar. Ao adicionar um histórico, a ferramenta passa a permitir:

Análise de Tendência: A empresa melhorou ou piorou os erros de cobrança neste mês?

Gestão de Equipe: Se as divergências aumentaram em julho, houve algum problema com os vendedores?

Visualização Executiva: Gráficos bonitos são o que justificam o preço de venda/assinatura de um software (conforme discutimos na Proposta Comercial).

## 2. A Arquitetura Profissional (Como fazer)

Para dar "memória" ao sistema sem o deixar pesado ou difícil de instalar, a forma mais profissional é usar um Banco de Dados Embutido (SQLite).
O SQLite já vem instalado dentro do Python (não precisa de instalar servidores pesados) e guarda tudo num único ficheirinho leve (ex: banco_dados.db).

O que vamos guardar no Banco de Dados?

NÃO vamos guardar os milhares de registos do Excel (isso deixaria o sistema lento). Vamos guardar apenas o "Sumário Executivo" de cada vez que apertar o botão "Iniciar Conciliação".

Estrutura da Tabela (conciliacoes):

id: Identificador único (ex: 1, 2, 3...)

data_processamento: A data e hora exata em que ela gerou a planilha.

periodo_referencia: O mês/quinzena (ex: "Julho/2026 - Quinzena 1"). Podemos pedir para ela digitar isso na tela antes de enviar, ou o Python pode deduzir lendo as datas dos Excel.

qtd_perfeitos: (ex: 142)

qtd_historico: (ex: 18)

qtd_desmembrados: (ex: 5)

qtd_divergencias: (ex: 12)

taxa_sucesso: (ex: 92.7%)

## 3. Passo a Passo da Implementação (Roadmap)

Para não quebrar o que já está a funcionar perfeitamente, vamos fazer isso em 3 fases:

### FASE 1: Backend (Criar a Memória)

Atualizar o app.py para usar a bibliot  eca sqlite3 (nativa do Python) ou SQLAlchemy (mais profissional).

Sempre que a conciliação terminar e o JSON for gerado, o Python faz um INSERT no banco de dados guardando os números daquela sessão.

### FASE 2: Backend (Nova Rota de Leitura)

Criar um novo endpoint no FastAPI: GET /api/historico.

Quando o Front-end chamar esta rota, o Python vai ao banco de dados e devolve uma lista com todas as conciliações passadas, ordenadas por data.

### FASE 3: Frontend (Os Gráficos)

Criar uma nova "Aba" ou página na interface web chamada "Histórico & Analytics".

Usar uma biblioteca profissional de gráficos para React, como o Recharts ou Chart.js (são gratuitas, lindas e fáceis de usar com Tailwind).

Desenhar dois gráficos principais:

Gráfico de Linha: Mostrando a evolução da % de Sucesso ao longo das quinzenas.

Gráfico de Barras: Comparando o número de Divergências (erros) mês a mês.

Uma tabela simples abaixo dos gráficos com o histórico de execuções.

## 4. Arquivo Morto

o Python pode parar de apagar o ficheiro Conciliacao_Final.xlsx gerado. Ele pode movê-lo para uma pasta permanente chamada /arquivos_antigos e guardar o caminho no banco de dados.
Assim, na tela de Histórico, teria um botão para "Re-baixar" a planilha de Janeiro, mesmo que já estejamos em Dezembro!