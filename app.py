# -*- coding: utf-8 -*-
"""
app.py - Servidor de API para o Sistema de Conciliação Bancária (Carrilho Distribuidora)
==========================================================================================

Faz a ponte entre a Interface Web (index.html) e o motor de conciliação (conciliacao.py).

Fluxo da requisição POST /api/conciliar:
    1. Recebe os arquivos enviados pelo front-end (Argos + Banco).
    2. Salva-os em uma pasta temporária para o Pandas ler.
    3. Executa DataCleaner → ReconciliationEngine → ExcelReporter.
    4. Retorna o arquivo .xlsx como download direto ao navegador.
    5. Apaga a pasta temporária em background após o envio.

Como iniciar o servidor:
    pip install fastapi uvicorn python-multipart openpyxl pandas
    python app.py
    ou:
    uvicorn app:app --host 0.0.0.0 --port 8000 --reload

Acesso via navegador (front-end):
    http://localhost:8000
"""

import os
import uuid
import shutil
import logging
import base64
import sys
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import List
from fastapi.staticfiles import StaticFiles

import pandas as pd
import uvicorn
from fastapi import FastAPI, File, Form, UploadFile, BackgroundTasks, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

class PeriodoUpdate(BaseModel):
    periodo: str

# ---------------------------------------------------------------------------
# Importação do motor de conciliação (mesmo diretório)
# ---------------------------------------------------------------------------
try:
    from conciliacao import DataCleaner, ReconciliationEngine, ExcelReporter
except ImportError as exc:
    print(
        f"[ERRO FATAL] Não foi possível importar 'conciliacao.py'.\n"
        f"Certifique-se de que o arquivo 'conciliacao.py' está na mesma pasta.\n"
        f"Detalhe: {exc}"
    )
    sys.exit(1)

# ---------------------------------------------------------------------------
# Configuração de Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] %(asctime)s - %(message)s",
    datefmt="%H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Configuração do Banco de Dados SQLite (Memória do Sistema)
# ---------------------------------------------------------------------------
DB_PATH = Path(__file__).parent / "historico.db"
ARQUIVOS_ANTIGOS_DIR = Path(__file__).parent / "arquivos_antigos"
ARQUIVOS_ANTIGOS_DIR.mkdir(parents=True, exist_ok=True)

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS conciliacoes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                data_processamento DATETIME,
                perfeitos INTEGER,
                historico INTEGER,
                desmembrados INTEGER,
                divergencias INTEGER,
                taxa_sucesso REAL,
                caminho_arquivo TEXT
            )
        """)
        
        # Migração: adicionar coluna periodo se não existir
        try:
            cursor.execute("ALTER TABLE conciliacoes ADD COLUMN periodo TEXT DEFAULT 'Desconhecido'")
        except sqlite3.OperationalError:
            pass # A coluna já existe
            
        conn.commit()

init_db()

# ---------------------------------------------------------------------------
# Instância do Aplicativo FastAPI
# ---------------------------------------------------------------------------
app = FastAPI(
    title="API de Conciliação Bancária · Carrilho Distribuidora",
    description=(
        "Servidor local que recebe arquivos Argos e extratos bancários, "
        "executa o cruzamento automático e retorna a planilha de resultados."
    ),
    version="2.0.0",
    docs_url="/docs",      # Swagger UI disponível em http://localhost:8000/docs
    redoc_url="/redoc",
)

# ---------------------------------------------------------------------------
# CORS — Permite requisições do front-end rodando localmente no navegador
# ---------------------------------------------------------------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # Em produção, substitua por domínio específico
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Avisa o FastAPI para liberar o acesso público a tudo que estiver na pasta 'img'
app.mount("/img", StaticFiles(directory="img"), name="img")

# ---------------------------------------------------------------------------
# Servir o Front-end (index.html) como raiz do servidor
# ---------------------------------------------------------------------------
# Se o arquivo index.html existir na mesma pasta, ele será servido
# automaticamente em http://localhost:8000
_FRONTEND_PATH = Path(__file__).parent / "index.html"
if _FRONTEND_PATH.exists():
    from fastapi.responses import HTMLResponse

    @app.get("/", response_class=HTMLResponse, include_in_schema=False)
    async def serve_frontend():
        """Serve o front-end (index.html) na raiz do servidor."""
        return HTMLResponse(content=_FRONTEND_PATH.read_text(encoding="utf-8"))

# ---------------------------------------------------------------------------
# Pasta base para uploads temporários
# ---------------------------------------------------------------------------
BASE_TEMP_DIR = Path(__file__).parent / "temp_uploads"

# ---------------------------------------------------------------------------
# Funções Auxiliares
# ---------------------------------------------------------------------------


def _criar_pasta_temporaria() -> Path:
    """
    Cria uma pasta temporária única para cada requisição.

    Usa UUID para garantir que requisições simultâneas não colidam.

    Returns:
        Path da pasta criada.
    """
    pasta = BASE_TEMP_DIR / str(uuid.uuid4())
    pasta.mkdir(parents=True, exist_ok=True)
    logger.info(f"Pasta temporária criada: {pasta.name}")
    return pasta


def _salvar_arquivo(upload: UploadFile, pasta: Path) -> Path:
    """
    Salva um UploadFile do FastAPI fisicamente no disco.

    O Pandas exige que o arquivo esteja em disco para ler .xlsx via openpyxl.

    Args:
        upload: Arquivo recebido pelo endpoint.
        pasta:  Pasta de destino.

    Returns:
        Path completo do arquivo salvo.

    Raises:
        IOError: Se não for possível salvar o arquivo.
    """
    destino = pasta / upload.filename
    try:
        with open(destino, "wb") as f:
            conteudo = upload.file.read()
            f.write(conteudo)
        logger.info(
            f"Arquivo salvo: '{upload.filename}' "
            f"({len(conteudo) / 1024:.1f} KB)"
        )
        return destino
    except Exception as exc:
        raise IOError(
            f"Erro ao salvar o arquivo '{upload.filename}': {exc}"
        ) from exc


def _limpar_pasta_temporaria(pasta: Path) -> None:
    """
    Remove a pasta temporária e todo o seu conteúdo do disco.

    Esta função é chamada em background pelo FastAPI *após* o arquivo
    ter sido enviado com sucesso ao cliente, evitando acúmulo de lixo.

    Args:
        pasta: Path da pasta a remover.
    """
    try:
        if pasta.exists():
            shutil.rmtree(pasta)
            logger.info(f"Pasta temporária removida: {pasta.name}")
    except Exception as exc:
        # Falha na limpeza não é crítica — apenas loga o aviso
        logger.warning(f"Não foi possível remover pasta temporária: {exc}")


def _validar_extensao(filename: str) -> None:
    """
    Valida se a extensão do arquivo é suportada pelo sistema.

    Args:
        filename: Nome do arquivo a validar.

    Raises:
        HTTPException 400: Se a extensão não for .xlsx, .xls ou .csv.
    """
    ext = Path(filename).suffix.lower()
    if ext not in {".xlsx", ".xls", ".csv"}:
        raise HTTPException(
            status_code=400,
            detail=(
                f"Arquivo '{filename}' com formato '{ext}' não é suportado. "
                f"Use arquivos .xlsx, .xls ou .csv."
            ),
        )


# ---------------------------------------------------------------------------
# Endpoint Principal: POST /api/conciliar
# ---------------------------------------------------------------------------


@app.post(
    "/api/conciliar",
    summary="Executa a conciliação bancária",
    description=(
        "Recebe arquivos do sistema Argos e extratos bancários, "
        "executa as 3 regras de conciliação e retorna a planilha Excel consolidada."
    ),
    response_description="Arquivo Excel (.xlsx) com os 4 grupos de resultados.",
    tags=["Conciliação"],
)
async def conciliar(
    background_tasks: BackgroundTasks,
    argos_files: List[UploadFile] = File(..., description="Arquivos do sistema Argos (.xlsx)"),
    banco_files: List[UploadFile] = File(..., description="Extratos bancários (.xlsx)"),
    banco_nome: str = Form(default="generico", description="Nome do banco: caixa | banese | generico"),
) -> JSONResponse:
    """
    Endpoint principal de concilição bancária.

    Fluxo interno:
        1. Valida extensões de todos os arquivos recebidos.
        2. Cria pasta temporária única para a requisição.
        3. Salva os arquivos em disco.
        4. Processa cada arquivo Argos com DataCleaner.clean_argos().
        5. Processa cada arquivo de Banco com DataCleaner.clean_bank().
        6. Concatena os DataFrames de cada grupo.
        7. Executa ReconciliationEngine.execute_pipeline().
        8. Gera o arquivo Excel com ExcelReporter.generate_report().
        9. Lê o arquivo gerado, codifica em Base64 e retorna JSON.
       10. Agenda limpeza da pasta temporária em background após o envio.

    Args:
        background_tasks: Injeção do FastAPI para tarefas pós-resposta.
        argos_files:      Lista de arquivos do sistema Argos.
        banco_files:      Lista de extratos bancários.
        banco_nome:       Identificador do banco ("caixa", "banese" ou "generico").

    Returns:
        JSONResponse com ``resultados`` (contagens por grupo) e
        ``ficheiro_base64`` (string Base64 do .xlsx gerado).

    Raises:
        HTTPException 400: Para erros de formato, colunas ausentes ou dados inválidos.
        HTTPException 500: Para erros internos inesperados do servidor.
    """
    pasta_temp: Path = _criar_pasta_temporaria()

    # Agenda a limpeza da pasta ANTES de qualquer possível falha,
    # garantindo que ela sempre será removida ao final da requisição.
    background_tasks.add_task(_limpar_pasta_temporaria, pasta_temp)

    logger.info("=" * 60)
    logger.info(f"Nova requisição de conciliação recebida.")
    logger.info(f"  Arquivos Argos : {[f.filename for f in argos_files]}")
    logger.info(f"  Arquivos Banco : {[f.filename for f in banco_files]}")
    logger.info(f"  Banco Nome     : {banco_nome}")
    logger.info("=" * 60)

    try:
        # ── Etapa 1: Validar extensões ──────────────────────────────────
        for f in argos_files + banco_files:
            _validar_extensao(f.filename)

        # ── Etapa 2: Salvar arquivos em disco ───────────────────────────
        pasta_argos = pasta_temp / "argos"
        pasta_banco = pasta_temp / "banco"
        pasta_argos.mkdir()
        pasta_banco.mkdir()

        caminhos_argos = [_salvar_arquivo(f, pasta_argos) for f in argos_files]
        caminhos_banco = [_salvar_arquivo(f, pasta_banco) for f in banco_files]

        # ── Etapa 3: Processar arquivos Argos ───────────────────────────
        logger.info(f"Processando {len(caminhos_argos)} arquivo(s) Argos...")
        dfs_argos: list = []
        for caminho in caminhos_argos:
            logger.info(f"  → DataCleaner.clean_argos('{caminho.name}')")
            df = DataCleaner.clean_argos(str(caminho))
            dfs_argos.append(df)

        if not dfs_argos:
            raise ValueError("Nenhum arquivo Argos válido foi processado.")

        dfs_argos_validos = [df for df in dfs_argos if df is not None and not df.empty]
        df_argos_full = pd.concat(dfs_argos_validos, ignore_index=True) if dfs_argos_validos else pd.DataFrame()
        logger.info(
            f"Argos consolidado: {len(df_argos_full)} registro(s) no total."
        )

        # ── Etapa 4: Processar arquivos do Banco ────────────────────────
        logger.info(f"Processando {len(caminhos_banco)} arquivo(s) do banco '{banco_nome}'...")
        dfs_banco: list = []
        for caminho in caminhos_banco:
            logger.info(f"  → DataCleaner.clean_bank('{caminho.name}')")
            df = DataCleaner.clean_bank(str(caminho))
            dfs_banco.append(df)

        if not dfs_banco:
            raise ValueError("Nenhum arquivo de extrato bancário válido foi processado.")

        dfs_bancos_validos = [df for df in dfs_banco if df is not None and not df.empty]
        df_banco_full = pd.concat(dfs_bancos_validos, ignore_index=True) if dfs_bancos_validos else pd.DataFrame()
        logger.info(
            f"Banco consolidado: {len(df_banco_full)} crédito(s) no total."
        )

        # ── Etapa 5: Executar o motor de conciliação ────────────────────
        logger.info("Instanciando ReconciliationEngine e executando pipeline...")
        engine = ReconciliationEngine(df_argos_full, df_banco_full)
        relatorios = engine.execute_pipeline()

        # Resumo dos resultados no log do servidor
        for chave, df_resultado in relatorios.items():
            logger.info(f"  {chave}: {len(df_resultado)} registro(s)")

        # ── Etapa 6: Gerar arquivo Excel de saída e salvar no Arquivo Morto ──
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo = f"Conciliacao_{timestamp}.xlsx"
        caminho_saida = ARQUIVOS_ANTIGOS_DIR / nome_arquivo

        logger.info(f"Gerando Excel em: {caminho_saida.name}...")
        ExcelReporter.generate_report(relatorios, str(caminho_saida))
        logger.info("Excel gerado e salvo no Arquivo Morto com sucesso.")

        # ── Etapa 7: Ler o Excel gerado para memória e converter para Base64 ──
        with open(caminho_saida, "rb") as f:
            excel_bytes = f.read()
        base64_encoded = base64.b64encode(excel_bytes).decode('utf-8')

        # Extração direta e garantida das quantidades
        qtd_perfeitos = len(relatorios.get("1_Conciliado_Perfeito", []))
        qtd_historico = len(relatorios.get("2_Conciliado_Via_Historico", []))
        qtd_desmembrado = len(relatorios.get("3_Conciliado_Desmembrado", []))
        qtd_divergencias = len(relatorios.get("4_Divergencias_Pendentes", []))

        # Cálculo da Taxa de Sucesso
        sucesso = qtd_perfeitos + qtd_historico + qtd_desmembrado
        total = sucesso + qtd_divergencias
        taxa_sucesso = (sucesso / total * 100) if total > 0 else 0.0

        # Lógica blindada para extração do Período
        periodo_str = "Desconhecido"
        try:
            datas_todas = []
            if not df_argos_full.empty and 'Data' in df_argos_full.columns:
                datas_todas.extend(df_argos_full['Data'].dropna().tolist())
            if not df_banco_full.empty and 'Data' in df_banco_full.columns:
                datas_todas.extend(df_banco_full['Data'].dropna().tolist())
            
            if datas_todas:
                # Converte a lista de strings para datetime de forma segura
                datas_dt = pd.to_datetime(datas_todas, format='%d/%m/%Y', errors='coerce')
                datas_dt = datas_dt.dropna()
                if not datas_dt.empty:
                    data_min = datas_dt.min().strftime('%d/%m') # Ex: 16/06
                    data_max = datas_dt.max().strftime('%d/%m') # Ex: 30/06
                    periodo_str = f"{data_min} a {data_max}"
        except Exception as e:
            print(f"Aviso: Erro ao extrair período: {e}")

        # Gravar histórico no Banco de Dados
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO conciliacoes 
                (data_processamento, perfeitos, historico, desmembrados, divergencias, taxa_sucesso, caminho_arquivo, periodo)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (datetime.now(), qtd_perfeitos, qtd_historico, qtd_desmembrado, qtd_divergencias, taxa_sucesso, str(caminho_saida), periodo_str))
            novo_id = cursor.lastrowid
            conn.commit()

        # Devolve o JSON com os resultados e o ficheiro
        return {
            "resultados": {
                "perfeitos": qtd_perfeitos,
                "historico": qtd_historico,
                "desmembrados": qtd_desmembrado,
                "divergencias": qtd_divergencias
            },
            "ficheiro_base64": base64_encoded,
            "id_historico": novo_id,
            "periodo": periodo_str
        }

    except HTTPException:
        # Re-lança HTTPExceptions sem modificar (já têm status e mensagem)
        raise

    except (KeyError, ValueError) as exc:
        # Erros esperados: arquivo mal formatado, coluna ausente, etc.
        mensagem = str(exc)
        logger.error(f"Erro de validação dos dados: {mensagem}")
        raise HTTPException(
            status_code=400,
            detail=(
                f"Erro ao processar os arquivos: {mensagem}\n\n"
                f"Verifique se os arquivos estão no formato correto e tente novamente."
            ),
        ) from exc

    except FileNotFoundError as exc:
        mensagem = str(exc)
        logger.error(f"Arquivo não encontrado: {mensagem}")
        raise HTTPException(
            status_code=400,
            detail=f"Arquivo não encontrado durante o processamento: {mensagem}",
        ) from exc

    except PermissionError as exc:
        logger.error(f"Erro de permissão ao gerar o Excel: {exc}")
        raise HTTPException(
            status_code=500,
            detail=(
                "Erro de permissão ao gerar o arquivo de saída. "
                "Verifique se a pasta do servidor tem permissão de escrita."
            ),
        ) from exc

    except Exception as exc:
        # Captura qualquer outro erro inesperado
        logger.exception(f"Erro interno inesperado: {exc}")
        raise HTTPException(
            status_code=500,
            detail=(
                f"Erro interno no servidor ao processar a conciliação: {exc}\n"
                f"Consulte os logs do servidor para mais detalhes."
            ),
        ) from exc


# ---------------------------------------------------------------------------
# Endpoints de Histórico e Arquivo Morto
# ---------------------------------------------------------------------------

@app.get("/api/historico", tags=["Histórico"])
async def obter_historico() -> JSONResponse:
    """
    Retorna a lista de todas as conciliações executadas no passado,
    ordenada da mais recente para a mais antiga.
    """
    try:
        with sqlite3.connect(DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM conciliacoes ORDER BY data_processamento DESC")
            linhas = cursor.fetchall()
            
            resultados = [dict(linha) for linha in linhas]
            return JSONResponse(content=resultados)
    except Exception as exc:
        logger.error(f"Erro ao buscar histórico: {exc}")
        raise HTTPException(status_code=500, detail="Erro ao ler o histórico do banco de dados.")

@app.get("/api/download/{id}", tags=["Histórico"])
async def baixar_arquivo_antigo(id: int):
    """
    Baixa um arquivo de conciliação do Arquivo Morto pelo seu ID.
    """
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT caminho_arquivo FROM conciliacoes WHERE id = ?", (id,))
            linha = cursor.fetchone()
            
            if not linha:
                raise HTTPException(status_code=404, detail="Conciliação não encontrada.")
            
            caminho_arquivo = Path(linha[0])
            if not caminho_arquivo.exists():
                raise HTTPException(status_code=404, detail="Arquivo Excel não encontrado no disco.")
                
            return FileResponse(
                path=caminho_arquivo,
                filename=caminho_arquivo.name,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except HTTPException:
        raise
    except Exception as exc:
        logger.error(f"Erro ao baixar arquivo {id}: {exc}")
        raise HTTPException(status_code=500, detail="Erro interno ao processar o download.")

@app.delete("/api/historico/{id}", tags=["Histórico"])
async def excluir_historico(id: int):
    """
    Exclui o histórico de conciliação do banco de dados e apaga o arquivo físico.
    """
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT caminho_arquivo FROM conciliacoes WHERE id = ?", (id,))
            linha = cursor.fetchone()
            
            if not linha:
                raise HTTPException(status_code=404, detail="Conciliação não encontrada.")
            
            caminho_arquivo = Path(linha[0])
            if caminho_arquivo.exists():
                try:
                    os.remove(caminho_arquivo)
                    logger.info(f"Arquivo físico removido: {caminho_arquivo.name}")
                except OSError as exc:
                    logger.warning(f"Não foi possível remover o arquivo {caminho_arquivo.name}: {exc}")
            
            cursor.execute("DELETE FROM conciliacoes WHERE id = ?", (id,))
            conn.commit()
            
            return JSONResponse(content={"status": "success", "message": "Histórico excluído com sucesso."})
    except HTTPException:
        raise
    except Exception as exc:
        logger.error(f"Erro ao excluir histórico {id}: {exc}")
        raise HTTPException(status_code=500, detail="Erro interno ao excluir histórico.")

@app.put("/api/historico/{id}", tags=["Histórico"])
async def editar_historico(id: int, payload: PeriodoUpdate):
    """
    Edita o período de uma conciliação já gravada.
    """
    try:
        with sqlite3.connect(DB_PATH) as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE conciliacoes SET periodo = ? WHERE id = ?", (payload.periodo, id))
            if cursor.rowcount == 0:
                raise HTTPException(status_code=404, detail="Conciliação não encontrada.")
            conn.commit()
            return JSONResponse(content={"status": "success", "message": "Período atualizado."})
    except HTTPException:
        raise
    except Exception as exc:
        logger.error(f"Erro ao editar histórico {id}: {exc}")
        raise HTTPException(status_code=500, detail="Erro interno ao tentar atualizar o histórico.")

# ---------------------------------------------------------------------------
# Endpoint de Health Check
# ---------------------------------------------------------------------------


@app.get(
    "/api/health",
    summary="Verifica se o servidor está em execução",
    tags=["Sistema"],
)
async def health_check() -> JSONResponse:
    """
    Retorna o status do servidor e a versão do sistema.

    Útil para o front-end verificar se a API está disponível antes
    de habilitar o botão de upload.
    """
    return JSONResponse(
        content={
            "status": "ok",
            "sistema": "Conciliação Bancária · Carrilho Distribuidora",
            "versao": "2.0.0",
            "motor": "conciliacao.py (DataCleaner + ReconciliationEngine + ExcelReporter)",
            "bancos_suportados": ["caixa", "banese", "generico"],
        }
    )


# ---------------------------------------------------------------------------
# Ponto de Entrada — Inicia o Uvicorn automaticamente
# ---------------------------------------------------------------------------


if __name__ == "__main__":
    logger.info("=" * 60)
    logger.info("  Carrilho Distribuidora · API de Conciliação Bancária")
    logger.info("=" * 60)
    logger.info("  Servidor iniciando em: http://localhost:8000")
    logger.info("  Front-end:             http://localhost:8000")
    logger.info("  Documentação Swagger:  http://localhost:8000/docs")
    logger.info("  Health Check:          http://localhost:8000/api/health")
    logger.info("=" * 60)

    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=8000,
        reload=True,           # Reinicia automaticamente ao salvar o arquivo
        log_level="info",
    )
