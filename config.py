"""
Módulo de Configuração Central da Conciliação Bancária (Carrilho Distribuidora).

Centraliza parâmetros operacionais, janelas temporais, tolerâncias monetárias
e limites de busca algorítmica, permitindo parametrização estática ou via
variáveis de ambiente sem necessidade de alterar o código do motor.
"""

import os

def _get_float_env(key: str, default: float) -> float:
    try:
        return float(os.getenv(key, default))
    except (ValueError, TypeError):
        return default

def _get_int_env(key: str, default: int) -> int:
    try:
        return int(os.getenv(key, default))
    except (ValueError, TypeError):
        return default

def _get_str_env(key: str, default: str) -> str:
    val = os.getenv(key)
    return val if val is not None else default

def _get_list_env(key: str, default: list[str]) -> list[str]:
    val = os.getenv(key, "")
    if not val.strip():
        return list(default)
    if val.strip() == "*":
        return ["*"]
    if val.startswith("["):
        import json
        try:
            parsed = json.loads(val)
            if isinstance(parsed, list):
                return [str(x).strip() for x in parsed if str(x).strip()]
        except Exception:
            pass
    return [orig.strip() for orig in val.split(",") if orig.strip()]

# ==============================================================================
# JANELAS TEMPORAIS (DIAS)
# ==============================================================================

# Janela padrão em dias para conciliação bancária entre registros do Argos e do Banco.
# Mantida em 31 dias por decisão de negócio deliberada.
JANELA_DIAS: int = _get_int_env("CONCILIACAO_JANELA_DIAS", 31)

# Janela em dias para conciliação por aproximação de centavos (Regra 3.5).
JANELA_DIAS_CENTAVOS: int = _get_int_env("CONCILIACAO_JANELA_DIAS_CENTAVOS", 3)

# Limite de dias a partir do qual é gerado um alerta visual amarelo no relatório Excel.
DIAS_ALERTA_TEMPORAL: int = _get_int_env("CONCILIACAO_DIAS_ALERTA_TEMPORAL", 5)

# ==============================================================================
# TOLERÂNCIAS MONETÁRIAS (REAIS)
# ==============================================================================

# Tolerância em reais para conciliação por aproximação de centavos e desmembramentos.
TOLERANCIA_CENTAVOS: float = _get_float_env("CONCILIACAO_TOLERANCIA_CENTAVOS", 1.50)

# Tolerância máxima de desconto ou acréscimo identificado no texto histórico de PIX.
TOLERANCIA_DESCONTO_PIX: float = _get_float_env("CONCILIACAO_TOLERANCIA_DESCONTO_PIX", 15.00)

# Valor mínimo faltante para acionar a busca de desmembramento de notas.
LIMIAR_VALOR_FALTANTE_DESMEMBRAR: float = _get_float_env("CONCILIACAO_LIMIAR_DESMEMBRAR", 0.05)

# Tolerância de centavos para a validação matemática de integridade financeira do pipeline.
TOLERANCIA_INTEGRIDADE: float = _get_float_env("CONCILIACAO_TOLERANCIA_INTEGRIDADE", 0.01)

# ==============================================================================
# LIMITES DE PROCESSAMENTO / COMBINAÇÕES
# ==============================================================================

# Quantidade máxima de parcelas combinadas em desmembramentos (itertools.combinations).
# Evita explosão combinatória O(N^k) em bases volumosas.
MAX_COMBINACOES: int = _get_int_env("CONCILIACAO_MAX_COMBINACOES", 4)

# Timeout configurável (em segundos) para abortar buscas de combinações longas.
# Evita travamentos completos do pipeline ao rodar itertools.combinations com datasets grandes.
TIMEOUT_COMBINACOES_SEGUNDOS: float = _get_float_env("CONCILIACAO_TIMEOUT_COMBINACOES", 5.0)

# ==============================================================================
# SEGURANÇA & CORS (ORIGENS PERMITIDAS)
# ==============================================================================

# Origens locais padrão para desenvolvimento frontend / backend
DEFAULT_CORS_ORIGINS: list[str] = [
    "http://localhost:8000",
    "http://127.0.0.1:8000",
    "http://localhost:3000",
    "http://127.0.0.1:3000",
    "http://localhost:5173",
    "http://127.0.0.1:5173",
    "http://localhost:5500",
    "http://127.0.0.1:5500",
]

# Lista de origens explícitas autorizadas para comunicação cross-origin.
# Configurável via variável de ambiente CORS_ORIGINS (separada por vírgula ou JSON).
CORS_ORIGINS: list[str] = _get_list_env("CORS_ORIGINS", DEFAULT_CORS_ORIGINS)

# Expressão regular para autorizar origens dinâmicas (ex: deploys no Render *.onrender.com).
# Configurável via variável de ambiente CORS_ALLOW_ORIGIN_REGEX.
CORS_ALLOW_ORIGIN_REGEX: str = _get_str_env(
    "CORS_ALLOW_ORIGIN_REGEX",
    r"https://.*\.onrender\.com"
)
