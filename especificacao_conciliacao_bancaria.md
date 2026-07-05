# Especificação Técnica Funcional: Sistema de Conciliação Bancária Automatizada (Argos vs. Bancos)

Este documento descreve detalhadamente os requisitos, a arquitetura de software, as regras de negócio e a suíte de testes para a criação de um script automatizado em Python utilizando a biblioteca `pandas`. O objetivo é realizar a conciliação automática entre os relatórios de baixa do sistema interno (Argos) e os extratos bancários brutos (Caixa Econômica, Banese, etc.), tratando exceções complexas de valores divergentes e lançamentos desmembrados.

---

## 1. Visão Geral do Sistema e Fluxo de Dados

O sistema opera de forma offline e local. Recebe dois conjuntos de arquivos de entrada (em formatos Excel `.xlsx` ou CSV) e gera um arquivo consolidado em Excel como saída, categorizando os lançamentos de acordo com o nível de correspondência (Sucesso, Ajuste por Histórico, Lançamentos Desmembrados ou Divergência).


### 1.1 Arquivos de Entrada

1. **Relatório do Argos (Baixas de Vendedores):** Contém colunas como `Parceiro Descrição`, `Data`, `Valor`, `Evento Descrição` e `Histórico`. O script deve ser capaz de processar o arquivo bruto removendo colunas desnecessárias de metadados internos (como `Ordem`, `Conta`, `Evento`, `Tipo Evento`).
2. **Extratos Bancários:** Arquivos brutos baixados diretamente do Internet Banking de cada instituição (Caixa Econômica Federal, Banese, etc.). As colunas principais a serem mapeadas são: Data do lançamento, Valor do crédito e Descrição/Documento.

### 1.2 Arquivo de Saída

Um único arquivo Excel (`conciliacao_final.xlsx`) contendo quatro abas estruturadas:

- **1. Conciliado_Perfeito:** Lançamentos que bateram exatamente em valor e data (janela de tolerância).
- **2. Conciliado_Via_Historico:** Lançamentos em que o valor nominal da baixa divergia, mas o valor real foi extraído com sucesso do texto da coluna `Histórico`.
- **3. Conciliado_Desmembrado:** Lançamentos combinados (múltiplas baixas do Argos que somadas equivalem a um único crédito no banco).
- **4. Divergencias_Pendentes:** Registros do banco que não encontraram par no Argos, ou baixas do Argos sem reflexo no banco. Esta é a única aba onde o usuário precisará atuar manualmente.

---

## 2. Pipeline de Limpeza e Normalização (Data Cleaning)

Antes de executar o motor de busca, os dados brutos devem passar por um tratamento rigoroso:

### 2.1 Normalização do Argos

- **Remoção de Colunas Inúteis:** Descartar automaticamente as colunas `Ordem`, `Conta`, `Evento` e `Tipo Evento`.
- **Tratamento de Datas:** Converter a coluna `Data` para o tipo `datetime64[ns]`. Tratar formatos brasileiros (`DD/MM/AAAA`).
- **Tratamento de Valores:** Converter a coluna `Valor` para o tipo numérico float (`float64`). Garantir que separadores de milhar (ponto) e decimais (vírgula) sejam interpretados corretamente.
- **Limpeza de Texto:** Remover espaços em branco sobressalentes nas colunas `Parceiro Descrição` e `Histórico`, convertendo-os para caixa alta (Uppercase).

### 2.2 Normalização dos Extratos Bancários

- **Identificação de Créditos:** Filtrar apenas as entradas de dinheiro (valores positivos / créditos). Transações de débito devem ser desconsideradas para esta finalidade.
- **Padronização de Colunas:** Mapear dinamicamente os cabeçalhos do banco para uma estrutura interna padrão: `data_banco`, `valor_banco`, `descricao_banco`.

---

## 3. Regras de Negócio e Algoritmos de Match

O motor de conciliação processará os dados sequencialmente através de três regras prioritárias. Quando um registro encontra um par, ele é marcado como "Conciliado" e removido do pool de busca das próximas regras.

### Regra 1: O Cenário Perfeito (Busca Exata por Janela Temporal)

- **Critério:** O script busca no extrato bancário um crédito cujo valor seja **exatamente igual** ao valor de baixa do Argos.
- **Janela Temporal:** Como depósitos em fins de semana ou após o horário bancário entram na conta em datas posteriores, a busca deve considerar uma janela de tolerância de **0 a +3 dias** (ex: se a baixa no Argos foi 16/06/2026, o crédito correspondente no banco pode estar entre 16/06/2026 e 19/06/2026).
- **Ação:** Remove ambos os registros dos pools e insere na aba `Conciliado_Perfeito`.

### Regra 2: Busca por Extração de Texto no Histórico (Regex Overriding)

- **Critério:** Aplicável quando o valor na coluna `Valor` do Argos difere do banco, mas o vendedor anotou o valor correto no campo `Histórico` (ex: Baixa de `199,17`, mas no histórico consta: `"pix no valor de 200,00 dia 22/06"`).
- **Mecanismo:** O script deve aplicar uma Expressão Regular (Regex) na coluna `Histórico` para capturar padrões numéricos financeiros precedidos de termos como `"valor de"`, `"valor"`, `"pix de"`, ou números isolados flutuantes que correspondam ao formato de moeda.
  - _Regex Sugerida:_ `(?:valor de|valor|pix de)?\s*([0-9]{1,3}(?:\.[0-9]{3})*\,[0-9]{2})`
- **Validação:** Se o valor extraído via Regex do histórico bater exatamente com um crédito no banco dentro da janela de **-2 a +3 dias** da data do histórico, o lançamento é considerado validado.
- **Ação:** Registra na aba `Conciliado_Via_Historico`, documentando o valor original da baixa e o valor real conciliado.

### Regra 3: Agrupamento e Soma de Partes (Subset Sum para Desmembrados)

- **Critério:** Casos em que o cliente faz um único Pix de valor alto (ex: `800,00`), mas os vendedores dão baixas separadas no sistema que somam este valor (ex: uma baixa de `445,96` e outra de `354,04`, ambas vinculadas ao mesmo período/cliente).
- **Mecanismo:** Para os registros que restaram sem conciliação nas etapas 1 e 2:
  1. Agrupar os lançamentos do Argos por proximidade de data (mesmo dia ou diferença de 1 dia) ou por similaridade textual de fragmentos do nome do cliente se houver.
  2. Utilizar um algoritmo combinatório (módulo `itertools.combinations`) para testar se a soma de 2 ou 3 baixas do Argos resulta exatamente no valor de algum crédito órfão do banco.
  3. Para otimização de performance, o algoritmo deve testar combinações de no máximo 3 elementos dentro do mesmo intervalo de 3 dias.
- **Ação:** Vincula os N lançamentos do Argos ao único lançamento do banco e envia para a aba `Conciliado_Desmembrado`.

---

## 4. Estrutura do Código Fonte Sugerida

O projeto deve seguir princípios de Clean Code, modularidade e orientação a objetos (ou funções puras com tipagem clara).

```python
import pandas as pd
import re
from datetime import timedelta
import itertools

class DataCleaner:
    @staticmethod
    def clean_argos(df_raw: pd.DataFrame) -> pd.DataFrame:
        \"\"\"Remove colunas desnecessárias, limpa strings e converte tipos.\"\"\"
        pass

    @staticmethod
    def clean_bank(df_raw: pd.DataFrame, bank_name: str) -> pd.DataFrame:
        \"\"\"Normaliza o extrato bruto baseado nas colunas do banco específico.\"\"\"
        pass

class ReconciliationEngine:
    def __init__(self, df_argos: pd.DataFrame, df_bank: pd.DataFrame):
        self.df_argos = df_argos.copy()
        self.df_bank = df_bank.copy()

    def match_exact(self):
        \"\"\"Implementa a Regra 1.\"\"\"
        pass

    def match_by_history_regex(self):
        \"\"\"Implementa a Regra 2.\"\"\"
        pass

    def match_split_sums(self):
        \"\"\"Implementa a Regra 3.\"\"\"
        pass

    def execute_pipeline(self) -> dict:
        \"\"\"Executa as regras em ordem e retorna os dicionários de DataFrames prontos.\"\"\"
        pass

class ExcelReporter:
    @staticmethod
    def generate_report(data_sheets: dict, output_path: str):
        \"\"\"Salva o arquivo final usando openpyxl aplicando formatação visual leve.\"\"\"
        pass
```

## 5 - Arquitetura de Testes de Software

Para garantir a confiabilidade matemática e lógica da aplicação, o projeto obrigatoriamente deve incluir uma suíte de testes unitários utilizando a biblioteca nativa unittest ou pytest. Os testes simulam os dados exatos extraídos das imagens reais do problema.

```python
import unittest
import pandas as pd
from datetime import datetime

class TestReconciliationEngine(unittest.TestCase):
    
    def setUp(self):
        # Configuração de dados mockados baseados nas imagens reais
        self.argos_base = pd.DataFrame({
            'Parceiro Descrição': ['JUNIO SILVA DE MENDOCA', 'JOAO DE OLIVEIRA LIMA JUNIOR', 'LATICINIO SERTANEJO', 'LATICINIO SERTANEJO'],
            'Data': [pd.to_datetime('2026-06-16'), pd.to_datetime('2026-06-22'), pd.to_datetime('2026-06-16'), pd.to_datetime('2026-06-16')],
            'Valor': [512.50, 199.17, 445.96, 354.04],
            'Evento Descrição': ['Inclusão Vendas Atacado'] * 4,
            'Histórico': [
                'pix no valor de 512,50 dia 16/06',
                'pix no valor de 200,00 dia 22/06', # Caso de desvio (Regra 2)
                '16/06 comprovante no valor de 800 reais.', # Desmembrado parte 1 (Regra 3)
                '16/06 comprovante no valor de 800 reais.'  # Desmembrado parte 2 (Regra 3)
            ]
        })
        
        self.bank_base = pd.DataFrame({
            'data_banco': [pd.to_datetime('2026-06-16'), pd.to_datetime('2026-06-22'), pd.to_datetime('2026-06-17')],
            'valor_banco': [512.50, 200.00, 800.00],
            'descricao_banco': ['CREDITO PIX', 'CREDITO PIX', 'CREDITO PIX']
        })

    def test_regra_1_match_perfeito(self):
        \"\"\"Testa se o valor de 512.50 é conciliado perfeitamente na mesma data.\"\"\"
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_exact()
        # Valida se o registro de 512.50 foi para a aba correta
        self.assertTrue(any(res['Conciliado_Perfeito']['Valor'] == 512.50))

    def test_regra_2_ajuste_historico(self):
        \"\"\"Testa se a baixa de 199.17 encontra o crédito de 200.00 lendo o histórico.\"\"\"
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        # Força execução da regra 2
        res = engine.match_by_history_regex()
        self.assertTrue(any(res['Conciliado_Via_Historico']['valor_real_banco'] == 200.00))

    def test_regra_3_valores_desmembrados(self):
        \"\"\"Testa se as duas baixas (445.96 e 354.04) combinam com o crédito de 800.00 no banco.\"\"\"
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_split_sums()
        self.assertEqual(len(res['Conciliado_Desmembrado']), 2) # Duas linhas do Argos unificadas

if __name__ == '__main__':
    unittest.main()
```