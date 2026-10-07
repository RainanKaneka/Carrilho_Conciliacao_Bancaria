# Conciliação Bancária · Carrilho Distribuidora

Sistema de conciliação entre relatórios de baixas do **Argos** e extratos bancários. A aplicação cruza valores, datas, nomes de clientes e informações do histórico para gerar uma planilha Excel com resultados, divergências e indicadores de integridade.

O projeto reúne uma interface web em português, API em FastAPI, motor de processamento em Python e integração com Supabase para autenticação, histórico e armazenamento dos relatórios.

## Sumário

- [Funcionalidades](#funcionalidades)
- [Arquitetura e tecnologias](#arquitetura-e-tecnologias)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Requisitos](#requisitos)
- [Instalação e execução](#instalação-e-execução)
- [Configuração do Supabase](#configuração-do-supabase)
- [Arquivos de entrada](#arquivos-de-entrada)
- [Como usar](#como-usar)
- [Regras de conciliação](#regras-de-conciliação)
- [Relatório Excel](#relatório-excel)
- [Variáveis de ambiente](#variáveis-de-ambiente)
- [API](#api)
- [Uso do motor em Python](#uso-do-motor-em-python)
- [Testes](#testes)
- [Publicação do serviço](#publicação-do-serviço)
- [Solução de problemas](#solução-de-problemas)
- [Limitações e cuidados operacionais](#limitações-e-cuidados-operacionais)
- [Contribuição e documentação complementar](#contribuição-e-documentação-complementar)
- [Autoria e licença](#autoria-e-licença)

## Funcionalidades

- Envio de múltiplos relatórios Argos e extratos bancários na mesma execução.
- Normalização de cabeçalhos, datas brasileiras e valores como `R$ 1.234,56`.
- Identificação do banco pelo conteúdo ou nome do arquivo.
- Conciliação por valor, proximidade de datas, nome do cliente e texto do histórico.
- Combinação de baixas do mesmo cliente para encontrar um único crédito bancário.
- Tratamento de pequenas diferenças monetárias e descontos/acréscimos informados no histórico.
- Classificação de estornos e saídas reconhecidos pelo motor.
- Exportação para Excel com sete abas, resumo executivo, alertas temporais e verificação de integridade financeira.
- Login com Supabase Auth.
- Histórico de execuções com edição de período, anotações, exclusão individual ou em lote e download de relatórios anteriores.
- Dashboard com indicadores, gráficos e filtros do histórico.

> Os formatos e layouts efetivamente processados estão descritos em [Arquivos de entrada](#arquivos-de-entrada).

## Arquitetura e tecnologias

```mermaid
flowchart TD
    U["Usuário"] --> UI["Interface web · index.html"]
    UI --> AUTH["Supabase Auth"]
    UI --> API["API FastAPI · app.py"]
    API --> CLEAN["DataCleaner · leitura e normalização"]
    CLEAN --> ENGINE["ReconciliationEngine · regras de conciliação"]
    ENGINE --> REPORT["ExcelReporter · relatório XLSX"]
    REPORT --> API
    API --> UI
    API --> DB["Supabase · tabela conciliacoes"]
    API --> STORAGE["Supabase Storage · arquivos_antigos"]
    DB --> ANALYTICS["AnalyticsService · métricas e séries"]
    ANALYTICS --> UI
```

| Camada | Tecnologias |
| --- | --- |
| Backend | Python, FastAPI e Uvicorn |
| Processamento | pandas, openpyxl, expressões regulares e itertools |
| Leitura de PDF | pdfplumber |
| Frontend | HTML, JavaScript e Tailwind CSS via CDN |
| Gráficos e ícones | Chart.js e Phosphor Icons |
| Autenticação e persistência | Supabase Auth, Database e Storage |
| Configuração | Variáveis de ambiente e python-dotenv |
| Testes | unittest e FastAPI TestClient |

O FastAPI serve a interface em `/` e as imagens em `/img`. O frontend utiliza rotas relativas para se comunicar com a API; não há etapa de build com Node.js.

O processamento dos arquivos ocorre no servidor Python. O uso completo da interface depende de conexão com o Supabase e com os serviços externos que fornecem os recursos do frontend.

## Estrutura do projeto

```text
Carrilho_Conciliacao_Bancaria/
├── app.py                                 # API, autenticação e histórico
├── conciliacao.py                         # Limpeza, conciliação, Excel e analytics
├── config.py                              # Parâmetros operacionais e CORS
├── index.html                             # Interface web
├── requirements.txt                       # Dependências Python
├── test_conciliacao.py                     # Suíte de testes
├── img/
│   ├── carrilho.png
│   ├── gato.png
│   └── icon.png
├── diretrizes.md                           # Diretrizes originais
├── especificacao_conciliacao_bancaria.md    # Especificação inicial
├── plano_de_expansao.md                    # Planejamento do histórico/dashboard
├── .gitignore
└── README.md
```

Arquivos locais, não versionados:

- `.env`: credenciais e configuração do ambiente.
- `venv/`: ambiente virtual, caso criado com os comandos abaixo.
- `temp_uploads/`: arquivos temporários das requisições.
- `historico.db`: base SQLite opcional já existente, utilizada como fallback de leitura.
- Planilhas, CSVs e PDFs: ignorados pelo Git para evitar o versionamento dos dados de entrada.

## Requisitos

- Python; para os comandos deste guia, utilize um ambiente com **Python 3.11 ou superior**.
- Git para clonar o repositório.
- Navegador com JavaScript habilitado.
- Projeto Supabase configurado para utilizar login, histórico e armazenamento.
- Permissão de escrita na pasta do aplicativo.
- Relatórios Argos em `.xlsx` e extratos compatíveis com os leitores implementados.

As dependências estão em [requirements.txt](requirements.txt). O arquivo utiliza versões mínimas e não fixa todo o ambiente com um lockfile.

## Instalação e execução

### 1. Clonar o repositório

```bash
git clone https://github.com/RainanKaneka/Carrilho_Conciliacao_Bancaria.git
cd Carrilho_Conciliacao_Bancaria
```

### 2. Criar e ativar o ambiente virtual

**Windows — PowerShell:**

```powershell
py -3 -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Se a ativação do ambiente estiver bloqueada no PowerShell, utilize o executável diretamente:

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe app.py
```

**Linux/macOS:**

```bash
python3 -m venv venv
source venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### 3. Configurar o ambiente

Crie um arquivo `.env` na raiz do projeto:

```dotenv
SUPABASE_URL=https://SEU_PROJETO.supabase.co
SUPABASE_KEY=SUA_CHAVE_SERVICE_ROLE_DO_SERVIDOR

CONCILIACAO_JANELA_DIAS=31
CONCILIACAO_JANELA_DIAS_CENTAVOS=3
CONCILIACAO_TOLERANCIA_CENTAVOS=1.50
CONCILIACAO_TOLERANCIA_DESCONTO_PIX=15.00

CORS_ORIGINS=http://localhost:8000,http://127.0.0.1:8000
```

Substitua os valores ilustrativos e siga a seção [Configuração do Supabase](#configuração-do-supabase). O `app.py` carrega o `.env` antes de importar o motor e seus parâmetros.

### 4. Iniciar o servidor

```bash
python app.py
```

Alternativamente:

```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000 --reload
```

| Recurso | Endereço local |
| --- | --- |
| Interface web | http://localhost:8000 |
| Swagger UI | http://localhost:8000/docs |
| ReDoc | http://localhost:8000/redoc |
| Health check | http://localhost:8000/api/health |

O comando `python app.py` inicia o servidor em `0.0.0.0:8000` com recarga automática. O comando alternativo acima restringe o acesso ao computador local.

Sem credenciais Supabase, o servidor pode servir a interface e o health check, mas as rotas protegidas não estarão operacionais.

## Configuração do Supabase

A aplicação utiliza o Supabase para três funções: **autenticação**, **histórico de conciliações** e **armazenamento dos relatórios Excel**.

### Autenticação e frontend

1. Configure o login por e-mail e senha no projeto Supabase.
2. Crie os usuários que terão acesso ao sistema.
3. Configure `SUPABASE_URL` e `SUPABASE_KEY` no ambiente do backend.
4. Em [index.html](index.html), localize as constantes `SUPABASE_URL` e `SUPABASE_ANON_KEY` e substitua pelos dados públicos do mesmo projeto.

O frontend já contém configuração de um projeto Supabase. Ela não é atualizada automaticamente pelo `.env`; frontend e backend precisam apontar para o mesmo projeto.

Utilize apenas a chave pública `anon` no navegador. A chave `service_role`, utilizada no exemplo do backend para as operações de Database e Storage, deve ficar exclusivamente no servidor. O arquivo `.env` está incluído no `.gitignore`.

### Tabela de histórico

O código espera uma tabela chamada `conciliacoes`. O repositório não inclui uma migração SQL; o exemplo abaixo define uma estrutura compatível para **um projeto novo**:

```sql
create table public.conciliacoes (
    id bigint generated by default as identity primary key,
    data_processamento timestamptz not null default now(),
    periodo text not null default '',
    perfeitos integer not null default 0,
    historico integer not null default 0,
    desmembrados integer not null default 0,
    saidas_estornos integer not null default 0,
    divergencias integer not null default 0,
    taxa_sucesso double precision not null default 0,
    caminho_arquivo text,
    anotacao text not null default ''
);

alter table public.conciliacoes enable row level security;
```

Esse exemplo considera as operações de banco realizadas pelo backend com a chave `service_role`. Se utilizar uma chave com permissões diferentes no servidor, configure as políticas necessárias para as operações da aplicação.

### Storage

Crie um bucket **privado** chamado `arquivos_antigos`. O backend envia os arquivos para esse bucket e fornece o download por uma rota autenticada.

A criação da tabela, do bucket e dos usuários não ocorre automaticamente ao iniciar o aplicativo.

O histórico atual é compartilhado entre usuários autenticados: as rotas não filtram registros por proprietário. Se precisar de isolamento entre usuários ou empresas, será necessário adaptar o modelo e a autorização no backend.

## Arquivos de entrada

### Formatos efetivamente processados

| Origem | Formato | Comportamento atual |
| --- | --- | --- |
| Argos | `.xlsx` | Leitura com pandas/openpyxl e mapeamento de cabeçalhos |
| Banco | `.xlsx` | Leitura com pandas/openpyxl e normalização |
| Banco | `.pdf` | Leitores específicos para Banco do Nordeste e Nature/Sicoob |
| Argos ou banco | `.csv` e `.xls` | A seleção/validação permite essas extensões, mas os leitores tabulares atuais utilizam openpyxl; converta para `.xlsx` |
| Argos | `.pdf` | Não implementado pelo leitor Argos |

Os PDFs precisam conter texto extraível e seguir os layouts reconhecidos. Não há OCR para documentos digitalizados como imagem.

A identificação por conteúdo ou nome contempla Caixa Econômica, Banese, BNB, Sicoob/Nature, Bradesco, Itaú e Santander. Essa identificação não garante a leitura de todos os layouts exportados por essas instituições.

### Relatórios Argos

O leitor procura cabeçalhos na primeira linha ou nas primeiras 20 linhas do relatório. Exemplos de campos reconhecidos:

| Campo interno | Cabeçalhos reconhecidos por fragmentos |
| --- | --- |
| `Cliente` | `Cliente` ou `Parceiro Descrição` |
| `Valor` | Coluna contendo `valor` |
| `Data` | Coluna contendo `data`, sem `baixa` |
| `Histórico` | Coluna contendo `obs` ou `hist` |
| `Banco` | Coluna contendo `banco`, sem `baixa` |
| `Tipo Evento` | Coluna contendo `evento descri` |

A coluna de valor é essencial. Quando ausente, o leitor retorna uma base vazia. Cliente, data e histórico preenchidos ajudam a reduzir ambiguidades.

### Extratos bancários

Para as planilhas bancárias, utilize colunas identificáveis como:

- `Data`
- `Valor`
- `Histórico` ou `Descrição`
- `Tipo`, quando disponível

O leitor tabular mantém valores positivos; o motor separa saídas reconhecidas por `Tipo = D` e estornos do Argos.

Utilize datas completas em `DD/MM/AAAA`. Datas sem ano recebem o ano corrente do servidor, o que pode produzir referências incorretas em arquivos antigos.

## Como usar

1. Inicie o servidor e abra http://localhost:8000.
2. Faça login com um usuário do Supabase.
3. Selecione ou arraste os arquivos Argos para a área correspondente.
4. Adicione os extratos bancários do período.
5. Inicie a conciliação e aguarde o processamento.
6. Confira as contagens por categoria e baixe a planilha Excel.
7. Revise as divergências, a coluna `REGRA APLICADA`, os alertas temporais e o resumo de integridade.
8. Consulte o histórico e o dashboard para acompanhar execuções anteriores.

A API concatena os arquivos de cada origem antes de executar o motor. Evite enviar arquivos com períodos sobrepostos ou lançamentos repetidos.

Se o processamento terminar, mas a gravação no Supabase falhar, o relatório da execução ainda pode ser entregue ao navegador. Nesse caso, a execução pode não aparecer no histórico e `id_historico` pode ser `null`.

## Regras de conciliação

O pipeline remove os registros já utilizados das etapas seguintes. A ordem atual prioriza informações do histórico antes das correspondências puramente numéricas.

| Ordem | Regra | Critério principal |
| --- | --- | --- |
| 1 | 2 e 2.5 — Histórico e desmembramento guiado | Extrai valores e possíveis datas do histórico; procura o crédito e, quando necessário, outras baixas do mesmo cliente |
| 2 | 0.5 — Nome no histórico bancário | Verifica palavras do nome do cliente, diferença de valor dentro da tolerância e proximidade temporal |
| 3 | 1.1 — Valor único | Concilia valores que aparecem uma única vez em cada conjunto restante, sem limite de datas |
| 4 | 1 — Valor exato com concorrência | Prioriza coincidências do nome e datas próximas para valores iguais |
| 5 | 3 — Combinação de baixas | Busca a soma de múltiplas baixas do mesmo cliente para um crédito bancário |
| 6 | 3.5 — Aproximação de centavos | Tenta diferenças monetárias pequenas dentro de uma janela temporal menor |
| 7 | 4 — Divergências | Classifica o que restou como `Falta no Banco` ou `Sobrou no Banco / Faltou no Argos` |

Parâmetros padrão:

- Janela geral: **31 dias**, usando diferença absoluta entre datas nas regras que aplicam esse limite.
- Aproximação de centavos: até **R$ 1,50**, em até **3 dias**.
- Desconto/acréscimo via histórico: até **R$ 15,00**.
- Alerta temporal no Excel: diferença superior a **5 dias**.
- Timeout de busca combinatória: **5 segundos** por busca monitorada.

O parâmetro `CONCILIACAO_MAX_COMBINACOES` vale `4`, mas os laços atuais usam limite superior exclusivo: a Regra 3 testa grupos de 2 ou 3 baixas; a Regra 2.5 pode acrescentar até 3 baixas à principal.

A categoria “Conciliado Perfeito” também recebe correspondências por nome e aproximação de centavos. Consulte a regra registrada em cada linha para interpretar o resultado.

## Relatório Excel

O motor retorna sete conjuntos de dados. Na gravação do Excel, os underscores dos nomes são substituídos por espaços:

| Aba no Excel | Conteúdo |
| --- | --- |
| `0 Resumo Executivo` | Totais, volumes, taxa de sucesso, integridade e hash SHA-256 do processamento |
| `1 Conciliado Perfeito` | Correspondências por valor, nome e aproximação de centavos |
| `2 Conciliado Via Historico` | Ajustes identificados pelo texto do histórico |
| `3 Conciliado Desmembrado` | Baixas combinadas para um crédito |
| `4 Saidas Estornos` | Estornos e saídas reconhecidos pelo motor |
| `5 Divergencias Pendentes` | Lançamentos sem correspondência |
| `6 Resumo Integridade` | Comparação dos valores de entrada Argos com os valores classificados na saída |

As abas de lançamentos apresentam campos como banco, cliente, valor da baixa, datas, histórico, motivo da divergência, regra aplicada e observação.

Diferenças temporais acima do limite configurado recebem observação e destaque visual. Categorias vazias recebem uma mensagem informativa.

A verificação financeira compara os valores positivos do Argos com os valores conciliados e as divergências “Falta no Banco”. Ela verifica a conservação dos valores classificados; não comprova, sozinha, a identidade do pagador.

O hash do resumo representa os dados do processamento, incluindo seu timestamp. O `ExcelReporter` também calcula o hash SHA-256 do arquivo gerado e o backend o registra no log. Esses hashes não são uma assinatura com certificado digital.

As contagens representam linhas classificadas. Um crédito que concilie várias baixas pode produzir várias linhas. No resumo da API, saídas/estornos entram no numerador da taxa de sucesso; avalie esse indicador junto das categorias e dos valores.

## Variáveis de ambiente

Os valores monetários no ambiente utilizam **ponto decimal**, sem separador de milhar.

| Variável | Padrão | Finalidade |
| --- | --- | --- |
| `SUPABASE_URL` | Sem padrão | URL do projeto utilizado pelo backend |
| `SUPABASE_KEY` | Sem padrão | Chave do Supabase utilizada exclusivamente pelo backend |
| `CONCILIACAO_JANELA_DIAS` | `31` | Janela geral das regras temporais |
| `CONCILIACAO_JANELA_DIAS_CENTAVOS` | `3` | Janela da aproximação de centavos |
| `CONCILIACAO_DIAS_ALERTA_TEMPORAL` | `5` | Limite para destaque temporal no Excel |
| `CONCILIACAO_TOLERANCIA_CENTAVOS` | `1.50` | Tolerância de pequenas diferenças e somas |
| `CONCILIACAO_TOLERANCIA_DESCONTO_PIX` | `15.00` | Limite de ajuste via histórico |
| `CONCILIACAO_LIMIAR_DESMEMBRAR` | `0.05` | Valor faltante mínimo para buscar partes adicionais |
| `CONCILIACAO_TOLERANCIA_INTEGRIDADE` | `0.01` | Tolerância da verificação financeira |
| `CONCILIACAO_MAX_COMBINACOES` | `4` | Limite superior utilizado nas buscas combinatórias |
| `CONCILIACAO_TIMEOUT_COMBINACOES` | `5.0` | Timeout em segundos por busca monitorada |
| `CORS_ORIGINS` | Origens locais | Lista separada por vírgulas ou array JSON |
| `CORS_ALLOW_ORIGIN_REGEX` | `https://.*\.onrender\.com` | Expressão regular adicional para origens permitidas |

As origens locais padrão contemplam `localhost` e `127.0.0.1` nas portas `8000`, `3000`, `5173` e `5500`.

Reinicie o servidor após alterar parâmetros de ambiente.

## API

A versão declarada da API é **2.0.0**. As rotas de conciliação, histórico e analytics exigem:

```http
Authorization: Bearer TOKEN_DE_ACESSO_SUPABASE
```

| Método | Rota | Função |
| --- | --- | --- |
| `GET` | `/` | Interface web |
| `GET` | `/api/health` | Status do servidor, sem autenticação |
| `POST` | `/api/conciliar` | Processar arquivos e retornar resultados e Excel em Base64 |
| `GET` | `/api/historico` | Consultar o histórico |
| `GET` | `/api/analytics/resumo` | Métricas e séries do histórico |
| `GET` | `/api/download/{id}` | Baixar relatório armazenado |
| `PUT` | `/api/historico/{id}` | Atualizar o período |
| `PUT` | `/api/historico/{id}/anotacao` | Atualizar a anotação |
| `DELETE` | `/api/historico/{id}` | Excluir registro e tentar remover o arquivo |
| `DELETE` | `/api/historico/bulk` | Excluir vários registros e tentar remover seus arquivos |

### Enviar uma conciliação

A requisição utiliza `multipart/form-data`:

| Campo | Tipo | Obrigatório |
| --- | --- | --- |
| `argos_files` | Um ou mais arquivos | Sim |
| `banco_files` | Um ou mais arquivos | Sim |
| `banco_nome` | Texto usado como indicação do banco; padrão `generico` | Não |

Exemplo em Bash; no PowerShell, utilize `curl.exe` e adapte a continuação de linha:

```bash
curl -X POST http://localhost:8000/api/conciliar \
  -H "Authorization: Bearer SEU_TOKEN_DE_ACESSO" \
  -F "argos_files=@dados/argos.xlsx" \
  -F "banco_files=@dados/banco.xlsx" \
  -F "banco_nome=generico" \
  -o resultado.json
```

Repita os campos de arquivo para enviar múltiplos documentos. Os caminhos são ilustrativos; a pasta `dados/` não é criada pelo aplicativo.

Exemplo ilustrativo de resposta:

```json
{
  "resultados": {
    "perfeitos": 12,
    "historico": 2,
    "desmembrados": 4,
    "saidas_estornos": 1,
    "divergencias": 3
  },
  "ficheiro_base64": "CONTEUDO_BASE64_DO_XLSX",
  "id_historico": 42,
  "periodo": "01/06 a 30/06 | Bancos: BNB"
}
```

O endpoint retorna **JSON**. Para obter o Excel, decodifique `ficheiro_base64`; o frontend já realiza essa conversão.

### Filtros de analytics e edição do histórico

`GET /api/analytics/resumo` aceita os parâmetros opcionais `banco`, `dias`, `tipo_sucesso` e `busca`. Exemplo:

```text
/api/analytics/resumo?banco=BNB&dias=90&tipo_sucesso=todos
```

Corpos JSON de exemplo para as operações de histórico:

```json
{ "periodo": "Junho/2026 - Quinzena 1" }
```

```json
{ "anotacao": "Divergências encaminhadas para revisão." }
```

```json
{ "ids": [1, 2, 3] }
```

Consulte `/docs` para os esquemas interativos. Algumas descrições antigas dos endpoints ainda mencionam quatro grupos ou download direto; o comportamento implementado retorna o JSON descrito acima.

## Uso do motor em Python

O motor pode ser utilizado diretamente, sem login ou persistência no Supabase:

```python
from dotenv import load_dotenv

load_dotenv()

from conciliacao import DataCleaner, ReconciliationEngine, ExcelReporter

argos = DataCleaner.clean_argos("dados/argos.xlsx")
banco = DataCleaner.clean_bank("dados/banco.xlsx", banco_hint="generico")

motor = ReconciliationEngine(argos, banco)
relatorios = motor.execute_pipeline()

hash_arquivo = ExcelReporter.generate_report(
    relatorios,
    "Conciliacao_Final.xlsx",
)

print(f"Relatório gerado. SHA-256: {hash_arquivo}")
```

Execute esse código em um script na raiz do projeto, com os arquivos nos caminhos indicados.

O arquivo `conciliacao.py` termina com um bloco de entrada vazio; executar apenas `python conciliacao.py` não inicia uma CLI.

## Testes

A suíte existente cobre conversão monetária, normalização, regras de conciliação, anti-reutilização de registros, Excel, integridade, parâmetros, CORS, arquivos corrompidos e analytics.

Com as dependências instaladas:

```bash
python -m unittest test_conciliacao.py -v
```

Os testes HTTP usam o `TestClient`. Se a instalação não disponibilizar `httpx`, instale-o no ambiente de desenvolvimento:

```bash
python -m pip install httpx
```

Os testes de integração com documentos reais do BNB dependem de arquivos locais na pasta `bnb/` e são ignorados quando esses arquivos não existem. Esses documentos não acompanham o repositório.

Para executar a suíte, mantenha os parâmetros operacionais padrão: há testes que verificam esses valores e as origens CORS padrão. Utilize um ambiente de teste sem credenciais de produção, pois alguns testes da API podem consultar o histórico configurado.

## Publicação do serviço

Para executar em um serviço que fornece a variável `PORT`, como o ambiente de hospedagem mencionado no código:

**Instalação:**

```bash
python -m pip install -r requirements.txt
```

**Inicialização em um shell POSIX:**

```bash
python -m uvicorn app:app --host 0.0.0.0 --port "${PORT:-8000}"
```

Configure as credenciais Supabase no ambiente do serviço, atualize as constantes públicas do frontend e ajuste as origens CORS conforme os domínios utilizados.

Mantenha `index.html` e a pasta `img/` junto ao backend. Não utilize `--reload` no comando de publicação. O repositório não inclui Dockerfile ou configuração declarativa de implantação.

## Solução de problemas

| Sintoma | O que verificar |
| --- | --- |
| `Autenticação não configurada.` | Presença e validade de `SUPABASE_URL` e `SUPABASE_KEY`; reinicie o servidor |
| Login funciona, mas a API rejeita o token | Frontend e backend precisam utilizar o mesmo projeto Supabase |
| Erro de conexão/DNS no login | Internet disponível e projeto Supabase acessível |
| Histórico não aparece após conciliar | Tabela, bucket, permissões do backend e avisos de persistência no log |
| CSV ou XLS gera base vazia | Converta para `.xlsx`; os leitores tabulares atuais usam openpyxl |
| PDF não gera lançamentos | Confirme texto extraível e compatibilidade com os leitores BNB/Nature |
| Relatório Argos vazio | Verifique a coluna de valor, cabeçalhos e formato do arquivo |
| Datas do relatório antigo aparecem com ano incorreto | Utilize datas completas; datas sem ano recebem o ano corrente |
| Erro ao abrir ou salvar o Excel | Feche o arquivo de saída e confira as permissões de escrita |
| Erro ao iniciar relacionado a `img` | Confirme que a pasta de imagens está presente |
| Navegador bloqueia requisição por CORS | Ajuste `CORS_ORIGINS` e `CORS_ALLOW_ORIGIN_REGEX` para a origem utilizada |
| `TestClient` informa dependência ausente | Instale `httpx` no ambiente de testes |

## Limitações e cuidados operacionais

- A aplicação concilia arquivos exportados; não há conexão direta com internet banking ou API do Argos.
- A correspondência por valor único não limita a diferença entre datas. Revise os alertas temporais, mesmo em categorias conciliadas.
- O motor evita reutilizar registros durante o pipeline, mas não elimina automaticamente linhas duplicadas presentes nos arquivos enviados.
- O processamento e o relatório em Base64 utilizam memória do servidor. O timeout controla buscas combinatórias específicas, não a duração total da requisição.
- O leitor de planilhas bancárias filtra valores negativos antes do motor. A aba de saídas/estornos não representa necessariamente todas as movimentações de débito do extrato.
- O fallback SQLite lê uma base `historico.db` existente; a aplicação atual não cria nem grava novos históricos nessa base.
- A limpeza de uploads temporários é agendada pelo backend; relatórios históricos ficam no Supabase Storage.
- As exclusões de histórico tentam remover o arquivo correspondente. Falhas de Storage são registradas e podem deixar arquivos sem registro associado.
- O repositório não contém migrações, arquivos reais de entrada ou isolamento do histórico por usuário.

## Contribuição e documentação complementar

Para contribuir, crie uma branch, descreva a mudança e execute os testes relacionados antes de abrir um pull request. Utilize dados sintéticos para reproduzir problemas e preserve os documentos financeiros fora do controle de versão.

Documentos de contexto:

- [Especificação técnica funcional](especificacao_conciliacao_bancaria.md)
- [Diretrizes originais](diretrizes.md)
- [Plano de expansão](plano_de_expansao.md)

Esses documentos registram o planejamento inicial e podem divergir da implementação atual, especialmente em autenticação, persistência, regras, tolerâncias e quantidade de abas. Para o comportamento vigente, consulte o código e este README.

## Autoria e licença

Repositório mantido por [RainanKaneka](https://github.com/RainanKaneka).

O repositório não contém um arquivo `LICENSE`. Nenhuma licença específica é declarada neste README.
