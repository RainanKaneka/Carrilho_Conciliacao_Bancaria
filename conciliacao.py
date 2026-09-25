import argparse
import warnings
import os
import zipfile
import logging
from pathlib import Path
import pandas as pd
import openpyxl
import re
import pdfplumber
import itertools
import hashlib
import time
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

logger = logging.getLogger(__name__)


class CorruptedFileError(Exception):
    """Exceção levantada quando um arquivo está corrompido, vazio ou não pode ser aberto pelo Excel/sistema."""
    pass

import config
from config import (
    JANELA_DIAS,
    JANELA_DIAS_CENTAVOS,
    DIAS_ALERTA_TEMPORAL,
    TOLERANCIA_CENTAVOS,
    TOLERANCIA_DESCONTO_PIX,
    LIMIAR_VALOR_FALTANTE_DESMEMBRAR,
    TOLERANCIA_INTEGRIDADE,
    MAX_COMBINACOES,
    TIMEOUT_COMBINACOES_SEGUNDOS,
)

def parse_currency(val, default=None):
    """Função blindada para converter qualquer formato de dinheiro/moeda para float.

    Suporta:
    - Formato brasileiro com milhar e vírgula: "1.234,56"
    - Formato com R$: "R$ 1.500,00", "R$512,50"
    - Numéricos diretos: int, float
    - Strings nulas, vazias ou inválidas retornam default (None por padrão).
    """
    if pd.isna(val):
        return default
    if isinstance(val, (int, float)):
        return float(val)
    val = str(val).upper().replace('R$', '').strip()
    if not val:
        return default
    if '.' in val and ',' in val:
        val = val.replace('.', '').replace(',', '.')
    elif ',' in val:
        val = val.replace(',', '.')
    try:
        return float(val)
    except (ValueError, TypeError):
        return default


def safe_float(val, default=0.0):
    """Atalho blindado para parse_currency com fallback padrão 0.0."""
    return parse_currency(val, default=default)


class DataCleaner:
    @staticmethod
    def is_excel_valid(file_path: str) -> tuple[bool, str]:
        """
        Verifica se um arquivo Excel (.xlsx, .xls) pode ser aberto e lido corretamente.

        Retorna:
            tuple[bool, str]: (True, "") se válido, ou (False, "motivo da falha") se corrompido/inválido.
        """
        path = Path(file_path)
        if not path.exists():
            return False, f"Arquivo não encontrado: '{file_path}'"

        try:
            tamanho = path.stat().st_size
        except OSError as e:
            return False, f"Erro ao acessar arquivo '{path.name}': {e}"

        if tamanho == 0:
            return False, f"O arquivo '{path.name}' está vazio (0 bytes) e não pode ser processado."

        ext = path.suffix.lower()
        if ext == ".xlsx":
            if not zipfile.is_zipfile(file_path):
                return False, f"O arquivo '{path.name}' não é um arquivo Excel (.xlsx) válido ou está corrompido (estrutura ZIP inválida)."

            try:
                with zipfile.ZipFile(file_path, "r") as zf:
                    bad_file = zf.testzip()
                    if bad_file:
                        return False, f"O arquivo '{path.name}' possui blocos de dados corrompidos no arquivo ZIP ({bad_file})."
                    namelist = zf.namelist()
                    if "[Content_Types].xml" not in namelist:
                        return False, f"O arquivo '{path.name}' não possui a estrutura interna esperada de uma planilha Excel (.xlsx)."
            except zipfile.BadZipFile as e:
                return False, f"O arquivo '{path.name}' está corrompido (falha na descompressão ZIP): {e}"
            except Exception as e:
                return False, f"Erro ao verificar integridade do arquivo '{path.name}': {e}"

            wb = None
            try:
                wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
                if not wb.sheetnames:
                    return False, f"A planilha '{path.name}' não contém nenhuma aba legível."
                first_sheet = wb[wb.sheetnames[0]]
                _ = first_sheet.max_row
            except Exception as e:
                return False, f"O arquivo Excel '{path.name}' está corrompido e não pôde ser aberto: {e}"
            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

            return True, ""

        elif ext == ".xls":
            try:
                with open(file_path, "rb") as f:
                    header = f.read(8)
                if header != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                    return False, f"O arquivo '{path.name}' não possui cabeçalho binário válido do formato .xls."
                return True, ""
            except Exception as e:
                return False, f"Erro ao validar arquivo .xls '{path.name}': {e}"

        else:
            return False, f"Extensão '{ext}' não é suportada para validação de Excel."

    @staticmethod
    def validate_excel(file_path: str) -> None:
        """
        Valida que o arquivo Excel abre corretamente antes do processamento.

        Raises:
            FileNotFoundError: Se o arquivo não existir.
            CorruptedFileError: Se o arquivo estiver vazio, corrompido ou ilegível.
        """
        is_valid, motivo = DataCleaner.is_excel_valid(file_path)
        if not is_valid:
            if "não encontrado" in motivo.lower():
                raise FileNotFoundError(motivo)
            raise CorruptedFileError(motivo)

    @staticmethod
    def validate_file(file_path: str) -> None:
        """
        Valida que o arquivo (Excel, PDF ou CSV) pode ser aberto e lido antes de processar.

        Raises:
            FileNotFoundError: Se o arquivo não existir.
            CorruptedFileError: Se o arquivo estiver corrompido ou ilegível.
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"Arquivo não encontrado: '{file_path}'")
        if path.stat().st_size == 0:
            raise CorruptedFileError(f"O arquivo '{path.name}' está vazio (0 bytes) e não pode ser processado.")

        ext = path.suffix.lower()
        if ext in [".xlsx", ".xls"]:
            DataCleaner.validate_excel(file_path)
        elif ext == ".pdf":
            try:
                import pdfplumber
                with pdfplumber.open(file_path) as pdf:
                    if not pdf.pages:
                        raise CorruptedFileError(f"O arquivo PDF '{path.name}' não possui páginas válidas.")
            except CorruptedFileError:
                raise
            except Exception as e:
                raise CorruptedFileError(f"O arquivo PDF '{path.name}' está corrompido ou ilegível: {e}")
        elif ext == ".csv":
            try:
                with open(file_path, "rb") as f:
                    sample = f.read(8192)
                if b"\x00" in sample:
                    raise CorruptedFileError(f"O arquivo CSV '{path.name}' contém bytes nulos inválidos.")
            except CorruptedFileError:
                raise
            except Exception as e:
                raise CorruptedFileError(f"Erro ao validar arquivo CSV '{path.name}': {e}")

    @staticmethod
    def _detect_bank_from_content(df: pd.DataFrame, file_path: str = "") -> str:
        """Detecção inteligente do banco analisando cabeçalhos e conteúdo inicial do arquivo."""
        try:
            content = ' '.join(str(c).lower() for c in df.columns)
            content += ' ' + ' '.join(str(v).lower() for v in df.head(30).values.flatten())
            
            if "caixa econ" in content or "cef" in content or "caixa econômica" in content:
                return "CAIXA ECONOMICA"
            if "banese" in content or "banco do estado de sergipe" in content:
                return "BANESE"
            if "banco do nordeste" in content or "bnb" in content:
                return "BNB"
        except Exception:
            pass
            
        nome_arquivo = str(file_path).lower()
        if 'caixa' in nome_arquivo: 
            return 'CAIXA ECONOMICA'
        elif 'banese' in nome_arquivo: 
            return 'BANESE'
            
        return "BANCO DESCONHECIDO"

    @staticmethod
    def clean_argos(file_path: str) -> pd.DataFrame:
        DataCleaner.validate_excel(file_path)
        try:
            import pandas as pd
            df = pd.read_excel(file_path, engine='openpyxl')
            
            banco_detectado = DataCleaner._detect_bank_from_content(df, file_path)
            if banco_detectado == "BANCO DESCONHECIDO":
                banco_detectado = "ARGOS"
            
            header_idx = None
            cols_str = ' '.join(str(c).lower() for c in df.columns)
            if not ('valor' in cols_str and ('cliente' in cols_str or 'parceiro' in cols_str)):
                for idx, row in df.head(20).iterrows():
                    row_str = ' '.join(str(val).lower() for val in row.values)
                    if 'valor' in row_str and ('cliente' in row_str or 'parceiro' in row_str):
                        header_idx = idx
                        break
            
            if header_idx is not None:
                df = pd.read_excel(file_path, header=header_idx + 1, engine='openpyxl')
            
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if 'banco' in col_lower and 'baixa' not in col_lower and 'Banco' not in col_mapping.values(): 
                    col_mapping[col] = 'Banco'
                elif ('cliente' in col_lower or 'parceiro' in col_lower) and 'Cliente' not in col_mapping.values(): 
                    col_mapping[col] = 'Cliente'
                elif 'valor' in col_lower and 'Valor' not in col_mapping.values(): 
                    col_mapping[col] = 'Valor'
                elif 'data' in col_lower and 'baixa' not in col_lower and 'Data' not in col_mapping.values(): 
                    col_mapping[col] = 'Data'
                elif ('obs' in col_lower or 'hist' in col_lower) and 'Histórico' not in col_mapping.values(): 
                    col_mapping[col] = 'Histórico'
                elif 'evento descri' in col_lower and 'Tipo Evento' not in col_mapping.values():
                    col_mapping[col] = 'Tipo Evento'
            
            df = df.rename(columns=col_mapping)
            df = df.loc[:, ~df.columns.duplicated()].copy()
            cols_to_keep = [c for c in ['Banco', 'Cliente', 'Valor', 'Data', 'Histórico', 'Tipo Evento'] if c in df.columns]
            df = df[cols_to_keep]
            
            if 'Valor' not in df.columns: return pd.DataFrame()
            if 'Data' not in df.columns: df['Data'] = ''
            if 'Histórico' not in df.columns: df['Histórico'] = ''
            if 'Tipo Evento' not in df.columns: df['Tipo Evento'] = ''
            if 'Banco' not in df.columns: df['Banco'] = ''
            if 'Cliente' not in df.columns: df['Cliente'] = 'CLIENTE NÃO INFORMADO'
            
            df['Banco'] = df['Banco'].fillna(banco_detectado)
            df['Banco'] = df['Banco'].replace(r'^\s*$', banco_detectado, regex=True)
            
            df = df.dropna(subset=['Valor'])

            # NOVO PARSER DE DATA BLINDADO (Padrão BR)
            def parse_date_br(val):
                import datetime
                if pd.isna(val) or str(val).strip() == '': return ''
                if isinstance(val, datetime.datetime): return val.strftime('%d/%m/%Y')
                val_str = str(val).strip().replace('.', '')
                try:
                    partes = val_str.split('/')
                    if len(partes) == 2:
                        current_year = datetime.datetime.now().year
                        val_str = f"{val_str}/{current_year}"
                    return pd.to_datetime(val_str, dayfirst=True).strftime('%d/%m/%Y')
                except:
                    return val_str

            df['Valor'] = df['Valor'].apply(parse_currency)
            df['Data'] = df['Data'].apply(parse_date_br)
            df['Histórico'] = df['Histórico'].fillna('')
            
            return df.dropna(subset=['Valor'])
        except (CorruptedFileError, FileNotFoundError):
            raise
        except (zipfile.BadZipFile, openpyxl.utils.exceptions.InvalidFileException) as exc:
            raise CorruptedFileError(f"O arquivo '{Path(file_path).name}' está corrompido: {exc}") from exc
        except Exception as e:
            logger.error(f"Erro ao ler Argos {file_path}: {e}")
            return pd.DataFrame()

    @staticmethod
    def _read_nature_pdf(file_path: str) -> pd.DataFrame:
        data = []
        current_date = None
        
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                for i, line in enumerate(lines):
                    date_match = re.search(r'^(\d{2} [A-Z]{3} \d{4})', line)
                    if date_match:
                        current_date = date_match.group(1)
                    
                    if "Transferência recebida pelo Pix" in line or "Transferência Recebida" in line:
                        val_match = re.search(r'([\d\.]+,\d{2})$', line)
                        
                        if val_match:
                            val_str = val_match.group(1)
                            name_match = re.search(r'Pix\s+([A-Z\s]+)\s+-', line, re.IGNORECASE)
                            if not name_match:
                                name_match = re.search(r'Recebida\s+([A-Za-z\s]+)\s+-', line, re.IGNORECASE)
                                
                            if name_match:
                                name = name_match.group(1).strip()
                            else:
                                clean_name = re.sub(r'Transferência recebida pelo Pix|Transferência Recebida', '', line, flags=re.IGNORECASE).strip()
                                clean_name = clean_name.replace(val_str, '').strip()
                                clean_name = clean_name.split('-')[0].strip()
                                name = clean_name
                                
                            date_str = ""
                            if current_date:
                                try:
                                    dt = pd.to_datetime(current_date, format="%d %b %Y")
                                    date_str = dt.strftime("%d/%m/%Y")
                                except:
                                    pass
                                    
                            data.append({
                                'Data': date_str,
                                'Histórico': name,
                                'Valor': val_str,
                                'Tipo': 'C'
                            })
                        else:
                            if i + 1 < len(lines):
                                next_line = lines[i+1].strip()
                                val_match_next = re.search(r'^([\d\.]+,\d{2})$', next_line)
                                if val_match_next:
                                    val_str = val_match_next.group(1)
                                    clean_name = re.sub(r'Transferência recebida pelo Pix|Transferência Recebida', '', line, flags=re.IGNORECASE).strip()
                                    clean_name = clean_name.split('-')[0].strip()
                                    name = clean_name
                                    
                                    date_str = ""
                                    if current_date:
                                        try:
                                            dt = pd.to_datetime(current_date, format="%d %b %Y")
                                            date_str = dt.strftime("%d/%m/%Y")
                                        except:
                                            pass
                                    
                                    data.append({
                                        'Data': date_str,
                                        'Histórico': name,
                                        'Valor': val_str,
                                        'Tipo': 'C'
                                    })
        
        df = pd.DataFrame(data)
        if not df.empty:
            df['Valor'] = df['Valor'].apply(safe_float)
            df['Banco'] = 'NATURE'
            
        return df

    @staticmethod
    def _read_bnb_pdf(file_path: str) -> pd.DataFrame:
        data = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    line = line.strip()
                    match = re.search(r'^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(\d+)\s+(-?\s*[\d\.]+,\d{2})\s+([\d\.]+,\d{2})$', line)
                    if match:
                        date_str = match.group(1)
                        historico = match.group(2).strip()
                        valor_str = match.group(4).replace(' ', '')
                        
                        if '-' in valor_str:
                            tipo = 'D'
                            valor_str = valor_str.replace('-', '')
                        else:
                            tipo = 'C'
                        
                        data.append({
                            'Data': date_str,
                            'Histórico': historico,
                            'Valor': valor_str,
                            'Tipo': tipo
                        })
        df = pd.DataFrame(data)
        if not df.empty:
            df['Valor'] = df['Valor'].apply(safe_float)
            df['Banco'] = 'BNB'
            
        return df

    @staticmethod
    def clean_bank(file_path: str) -> pd.DataFrame:
        DataCleaner.validate_file(file_path)
        try:
            import pandas as pd
            
            if str(file_path).lower().endswith('.pdf'):
                text = ""
                with pdfplumber.open(file_path) as pdf:
                    if pdf.pages:
                        text = pdf.pages[0].extract_text() or ""
                text_lower = text.lower()
                if "banco do nordeste" in text_lower or "bnb" in text_lower:
                    return DataCleaner._read_bnb_pdf(file_path)
                else:
                    return DataCleaner._read_nature_pdf(file_path)
                
            df = pd.read_excel(file_path, engine='openpyxl')
            
            banco_detectado = DataCleaner._detect_bank_from_content(df, file_path)
            
            header_idx = None
            cols_str = ' '.join(str(c).lower() for c in df.columns)
            if not ('data' in cols_str and 'valor' in cols_str):
                for idx, row in df.head(20).iterrows():
                    row_str = ' '.join(str(val).lower() for val in row.values)
                    if 'data' in row_str and 'valor' in row_str:
                        header_idx = idx
                        break
            
            if header_idx is not None:
                df = pd.read_excel(file_path, header=header_idx + 1, engine='openpyxl')
            
            col_mapping = {}
            for col in df.columns:
                col_lower = str(col).lower()
                if 'data' in col_lower and 'Data' not in col_mapping.values(): 
                    col_mapping[col] = 'Data'
                elif ('histórico' in col_lower or 'historico' in col_lower or 'descrição' in col_lower or 'hitórico' in col_lower) and 'Histórico' not in col_mapping.values(): 
                    col_mapping[col] = 'Histórico'
                elif 'valor' in col_lower and 'Valor' not in col_mapping.values(): 
                    col_mapping[col] = 'Valor'
                elif 'tipo' in col_lower and 'Tipo' not in col_mapping.values():
                    col_mapping[col] = 'Tipo'
            
            df = df.rename(columns=col_mapping)
            df = df.loc[:, ~df.columns.duplicated()].copy()
            cols_to_keep = [c for c in ['Data', 'Histórico', 'Valor', 'Tipo'] if c in df.columns]
            df = df[cols_to_keep]
            
            if 'Valor' not in df.columns: return pd.DataFrame()
            if 'Histórico' not in df.columns: df['Histórico'] = ''
            if 'Tipo' not in df.columns: df['Tipo'] = ''
            
            df = df.dropna(subset=['Valor'])

            def parse_date_br(val):
                import datetime
                if pd.isna(val) or str(val).strip() == '': return ''
                if isinstance(val, datetime.datetime): return val.strftime('%d/%m/%Y')
                val_str = str(val).strip().replace('.', '')
                try:
                    partes = val_str.split('/')
                    if len(partes) == 2:
                        current_year = datetime.datetime.now().year
                        val_str = f"{val_str}/{current_year}"
                    return pd.to_datetime(val_str, dayfirst=True).strftime('%d/%m/%Y')
                except:
                    return val_str

            df['Valor'] = df['Valor'].apply(parse_currency)
            df['Data'] = df['Data'].apply(parse_date_br)
            df = df.dropna(subset=['Valor'])
            df = df[df['Valor'] > 0]
            df['Histórico'] = df['Histórico'].fillna('')
            
            df['Banco'] = banco_detectado
            
            return df
        except (CorruptedFileError, FileNotFoundError):
            raise
        except (zipfile.BadZipFile, openpyxl.utils.exceptions.InvalidFileException) as exc:
            raise CorruptedFileError(f"O arquivo '{Path(file_path).name}' está corrompido: {exc}") from exc
        except Exception as e:
            logger.error(f"Erro ao ler Banco {file_path}: {e}")
            return pd.DataFrame()



class ReconciliationEngine:
    def __init__(
        self,
        df_argos,
        df_bank,
        janela_dias: int = config.JANELA_DIAS,
        tolerancia_centavos: float = config.TOLERANCIA_CENTAVOS,
        max_combinacoes: int = config.MAX_COMBINACOES,
        tolerancia_desconto_pix: float = config.TOLERANCIA_DESCONTO_PIX,
        janela_dias_centavos: int = config.JANELA_DIAS_CENTAVOS,
        timeout_combinacoes: float = config.TIMEOUT_COMBINACOES_SEGUNDOS,
    ):
        import pandas as pd
        
        self.janela_dias = int(janela_dias) if janela_dias is not None else config.JANELA_DIAS
        self.tolerancia_centavos = float(tolerancia_centavos) if tolerancia_centavos is not None else config.TOLERANCIA_CENTAVOS
        self.max_combinacoes = int(max_combinacoes) if max_combinacoes is not None else config.MAX_COMBINACOES
        self.tolerancia_desconto_pix = float(tolerancia_desconto_pix) if tolerancia_desconto_pix is not None else config.TOLERANCIA_DESCONTO_PIX
        self.janela_dias_centavos = int(janela_dias_centavos) if janela_dias_centavos is not None else config.JANELA_DIAS_CENTAVOS
        self.timeout_combinacoes = float(timeout_combinacoes) if timeout_combinacoes is not None else config.TIMEOUT_COMBINACOES_SEGUNDOS

        self.df_argos = df_argos.copy() if isinstance(df_argos, pd.DataFrame) else pd.DataFrame()
        self.df_bank = df_bank.copy() if isinstance(df_bank, pd.DataFrame) else pd.DataFrame()

        # BLINDAGEM: Garante que todos os valores monetários são floats matemáticos válidos
        if not self.df_argos.empty and 'Valor' in self.df_argos.columns:
            self.df_argos['Valor'] = self.df_argos['Valor'].apply(safe_float)
            
        if not self.df_bank.empty and 'Valor' in self.df_bank.columns:
            self.df_bank['Valor'] = self.df_bank['Valor'].apply(safe_float)

        saidas_estornos = []
        if not self.df_argos.empty:
            if 'Tipo Evento' not in self.df_argos.columns: self.df_argos['Tipo Evento'] = ''
            is_estorno = (self.df_argos['Valor'] < 0) | (self.df_argos['Tipo Evento'].astype(str).str.lower().str.contains('estorno'))
            for _, row in self.df_argos[is_estorno].iterrows():
                saidas_estornos.append({
                    'Banco': row.get('Banco', ''),
                    'Cliente': row.get('Cliente', ''),
                    'Valor': abs(row['Valor']),
                    'Data': row.get('Data', ''),
                    'Histórico': row.get('Histórico', ''),
                    'Baixas': '',
                    'Data Baixa': '',
                    'Motivo Divergência': 'Estorno (Argos)',
                    'Regra Aplicada': 'N/A'
                })
            self.df_argos = self.df_argos[~is_estorno & (self.df_argos['Valor'] > 0)]

        if not self.df_bank.empty:
            if 'Tipo' not in self.df_bank.columns: self.df_bank['Tipo'] = ''
            is_saida = (self.df_bank['Valor'] < 0) | (self.df_bank['Tipo'].astype(str).str.upper() == 'D')
            for _, row in self.df_bank[is_saida].iterrows():
                saidas_estornos.append({
                    'Banco': row.get('Banco', ''),
                    'Cliente': row.get('Histórico', ''),
                    'Valor': abs(row['Valor']),
                    'Data': row.get('Data', ''),
                    'Histórico': row.get('Histórico', ''),
                    'Baixas': '',
                    'Data Baixa': '',
                    'Motivo Divergência': 'Saída (Banco)',
                    'Regra Aplicada': 'N/A'
                })
            self.df_bank = self.df_bank[~is_saida & (self.df_bank['Valor'] > 0)]

        self.df_saidas_estornos = pd.DataFrame(saidas_estornos)

        if not self.df_argos.empty:
            self.df_argos['ID_Argos'] = range(len(self.df_argos))
        if not self.df_bank.empty:
            self.df_bank['ID_Bank'] = range(len(self.df_bank))

    @staticmethod
    def _calcular_assinatura_digital(df_argos: pd.DataFrame, df_bank: pd.DataFrame, resultados: dict, timestamp_str: str) -> str:
        """Calcula a assinatura digital (hash SHA-256) representativa dos dados e parâmetros da conciliação."""
        hasher = hashlib.sha256()
        hasher.update(timestamp_str.encode('utf-8'))
        
        # Hash dos DataFrames de entrada
        if not df_argos.empty:
            hasher.update(df_argos.to_csv(index=False).encode('utf-8'))
        if not df_bank.empty:
            hasher.update(df_bank.to_csv(index=False).encode('utf-8'))
            
        # Hash dos totais por categoria
        for k in sorted(resultados.keys()):
            df_k = resultados[k]
            if isinstance(df_k, pd.DataFrame) and not df_k.empty:
                total_k = round(float(df_k['Valor'].sum()), 2) if 'Valor' in df_k.columns else 0.0
                qtd_k = len(df_k)
                hasher.update(f"{k}:{qtd_k}:{total_k}".encode('utf-8'))
                
        return hasher.hexdigest()

    def execute_pipeline(self):
        import pandas as pd
        import re
        import itertools
        
        resultados = {
            '1_Conciliado_Perfeito': [],
            '2_Conciliado_Via_Historico': [],
            '3_Conciliado_Desmembrado': [],
            '4_Saidas_Estornos': self.df_saidas_estornos.to_dict('records') if hasattr(self, 'df_saidas_estornos') else [],
            '5_Divergencias_Pendentes': []
        }

        if self.df_argos.empty or self.df_bank.empty:
            for k in resultados.keys():
                resultados[k] = pd.DataFrame(columns=['Banco', 'Cliente', 'Valor', 'Data', 'Histórico', 'Baixas', 'Data Baixa', 'Motivo Divergência', 'Regra Aplicada'])
            
            soma_input = self.df_argos['Valor'].sum() if not self.df_argos.empty else 0.0
            status_msg = "OK - Nenhum centavo perdido ou duplicado" if soma_input == 0 else "ERRO (Perda/Duplicação identificada)"
            df_integridade = pd.DataFrame([
                {"Métrica": "Total Input Argos", "Valor": round(soma_input, 2)},
                {"Métrica": "Total Output Conciliado", "Valor": 0.0},
                {"Métrica": "Total Output Divergências (Falta Banco)", "Valor": 0.0},
                {"Métrica": "Total Output (Conciliado + Divergências)", "Valor": 0.0},
                {"Métrica": "Diferença (Perda/Duplicação)", "Valor": round(soma_input, 2)},
                {"Métrica": "Status da Integridade", "Valor": status_msg},
            ])

            timestamp_execucao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            assinatura_sha256 = self._calcular_assinatura_digital(
                self.df_argos, self.df_bank, resultados, timestamp_execucao
            )

            df_executivo = pd.DataFrame([
                {"Métrica": "Data/Hora de Processamento", "Valor": timestamp_execucao},
                {"Métrica": "Status da Conciliação", "Valor": "Entradas Vazias"},
                {"Métrica": "Taxa de Sucesso (Registros)", "Valor": "0.00%"},
                {"Métrica": "Taxa de Sucesso (Financeira)", "Valor": "0.00%"},
                {"Métrica": "Total de Registros Argos", "Valor": len(self.df_argos)},
                {"Métrica": "Total Volume Argos (R$)", "Valor": round(soma_input, 2)},
                {"Métrica": "Total de Registros Banco", "Valor": len(self.df_bank)},
                {"Métrica": "Total Volume Banco (R$)", "Valor": 0.0},
                {"Métrica": "Total Conciliado (Registros)", "Valor": 0},
                {"Métrica": "Total Conciliado (R$)", "Valor": 0.0},
                {"Métrica": "  - Match Perfeito (Qtd)", "Valor": 0},
                {"Métrica": "  - Match Perfeito (R$)", "Valor": 0.0},
                {"Métrica": "  - Via Histórico (Qtd)", "Valor": 0},
                {"Métrica": "  - Via Histórico (R$)", "Valor": 0.0},
                {"Métrica": "  - Desmembrado (Qtd)", "Valor": 0},
                {"Métrica": "  - Desmembrado (R$)", "Valor": 0.0},
                {"Métrica": "Total Saídas / Estornos (Registros)", "Valor": 0},
                {"Métrica": "Total Saídas / Estornos (R$)", "Valor": 0.0},
                {"Métrica": "Total Divergências Pendentes (Registros)", "Valor": 0},
                {"Métrica": "Total Divergências Pendentes (R$)", "Valor": 0.0},
                {"Métrica": "Status de Integridade Financeira", "Valor": status_msg},
                {"Métrica": "Assinatura Digital (Hash SHA-256)", "Valor": assinatura_sha256},
            ])

            return {
                '0_Resumo_Executivo': df_executivo,
                '1_Conciliado_Perfeito': resultados['1_Conciliado_Perfeito'],
                '2_Conciliado_Via_Historico': resultados['2_Conciliado_Via_Historico'],
                '3_Conciliado_Desmembrado': resultados['3_Conciliado_Desmembrado'],
                '4_Saidas_Estornos': resultados['4_Saidas_Estornos'],
                '5_Divergencias_Pendentes': resultados['5_Divergencias_Pendentes'],
                '6_Resumo_Integridade': df_integridade,
            }

        argos_pendentes = self.df_argos.copy()
        bank_pendentes = self.df_bank.copy()

        def remover_matched(ids_a, ids_b):
            nonlocal argos_pendentes, bank_pendentes
            argos_pendentes = argos_pendentes[~argos_pendentes['ID_Argos'].isin(ids_a)]
            bank_pendentes = bank_pendentes[~bank_pendentes['ID_Bank'].isin(ids_b)]

        # ==========================================
        # REGRA 2 e 2.5: VIA HISTÓRICO & DESMEMBRADO GUIADO
        # PRIORIDADE MÁXIMA: A intenção humana (escrita) supera a matemática crua.
        # ==========================================
        # regex ajustada para exigir indicativo claro de dinheiro (R$, reais, pix, valor, baixa) antes ou depois
        regex = re.compile(r'(?:PIX.*?|VALOR DE|COMPROVANTE.*?|PAGO.*?|RESTANTE.*?|BAIXA.*?|DE\s*|^|[^\d])R?\$?\s*(\d+(?:\.\d{3})*,\d{1,2})(?:\s*REAIS)?', re.IGNORECASE)
        
        m_a, m_b = [], []
        
        for i, row_a in argos_pendentes.iterrows():
            if row_a['ID_Argos'] in m_a: continue
            
            col_hist = 'OBS' if 'OBS' in row_a else 'Histórico'
            texto_hist = str(row_a.get(col_hist, ''))
            # Corrige erros de digitação comuns no histórico, como "192,oo" ao invés de "192,00"
            texto_hist_corrigido = re.sub(r'(\d+),([oO]+)', lambda m: f"{m.group(1)},{m.group(2).lower().replace('o', '0')}", texto_hist)
            # Corrige datas escritas com vírgula, ex: "DIA 22,06" -> "DIA 22/06"
            texto_hist_corrigido = re.sub(r'(dia\s*\d{1,2}),(\d{1,2})', r'\1/\2', texto_hist_corrigido, flags=re.IGNORECASE)
            # Adiciona ",00" a valores redondos perto de palavras-chave, ex: "pix de 1.000" -> "1.000,00"
            texto_hist_corrigido = re.sub(r'\b(valor(?: de)?|pix|r\$|pago)\s+(\d+(?:\.\d{3})*)(?![\d.,])', r'\1 \2,00', texto_hist_corrigido, flags=re.IGNORECASE)
            texto_hist_corrigido = re.sub(r'\b(\d+(?:\.\d{3})*)(?![\d.,])\s*(reais)', r'\1,00 \2', texto_hist_corrigido, flags=re.IGNORECASE)
            match = regex.search(texto_hist_corrigido)
            
            # Bloqueio rigoroso: Só extrai números se o histórico contiver palavras relacionadas a dinheiro
            # Isso evita extrair uma data (ex: "15/06" virar R$ 15,06 e roubar depósitos)
            palavras_dinheiro = ['pix', 'valor', 'reais', 'r$', 'pago', 'restante', 'baixa', 'comprovante']
            if match and not any(p in texto_hist_corrigido.lower() for p in palavras_dinheiro):
                v_temp = parse_currency(match.group(1))
                if v_temp is None or abs(v_temp - row_a['Valor']) > self.tolerancia_desconto_pix:
                    match = None

            if match:
                v_regex = parse_currency(match.group(1))
                if v_regex is None:
                    continue

                # Extrai possível data do histórico para usar de referência
                data_referencia = row_a['Data']
                match_data = re.search(r'dia\s*(\d{1,2})[/\.](\d{1,2})', texto_hist_corrigido, re.IGNORECASE)
                if match_data:
                    dia, mes = match_data.groups()
                    try:
                        ano = pd.to_datetime(row_a['Data'], format='%d/%m/%Y').year
                        data_referencia = f"{int(dia):02d}/{int(mes):02d}/{ano}"
                    except:
                        pass
                
                candidatos = bank_pendentes[(bank_pendentes['Valor'] == v_regex) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
                if not candidatos.empty:
                    candidatos['diff_dias'] = abs((pd.to_datetime(data_referencia, format='%d/%m/%Y', errors='coerce') - 
                                                   pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                    candidatos = candidatos.sort_values(by='diff_dias')
                    
                    for j, row_b in candidatos.iterrows():
                        if pd.isna(row_b['diff_dias']) or row_b['diff_dias'] <= self.janela_dias: 
                            valor_faltante = round(v_regex - row_a['Valor'], 2)
                            comb_encontrada = []
                            
                            # Tenta Desmembrar (OTIMIZADO)
                            if valor_faltante > LIMIAR_VALOR_FALTANTE_DESMEMBRAR:
                                cliente_atual = row_a.get('Cliente', None)
                                col_cliente = 'CLIENTES' if 'CLIENTES' in argos_pendentes.columns else 'Cliente'
                                
                                # FILTRO CRÍTICO: Limita a busca a notas que "cabem" no espaço vazio, evitando loop O(N³)
                                filter_mask = (
                                    (~argos_pendentes['ID_Argos'].isin(m_a)) & 
                                    (argos_pendentes['ID_Argos'] != row_a['ID_Argos']) &
                                    (argos_pendentes['Valor'] <= valor_faltante + self.tolerancia_centavos)
                                )
                                
                                # TRAVA MESTRA: Só desmembra se for o MESMO cliente!
                                if pd.notna(cliente_atual) and cliente_atual != 'CLIENTE NÃO INFORMADO':
                                    filter_mask = filter_mask & (argos_pendentes[col_cliente] == cliente_atual)
                                    
                                # TRAVA DE DATA: Só busca notas próximas
                                date_a = pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce')
                                if pd.notna(date_a):
                                    diff_dias_notas = abs((pd.to_datetime(argos_pendentes['Data'], format='%d/%m/%Y', errors='coerce') - date_a).dt.days)
                                    filter_mask = filter_mask & (diff_dias_notas <= self.janela_dias)
                                    
                                outras_notas = argos_pendentes[filter_mask]
                                fast_notas = [(row['ID_Argos'], row['Valor'], row.to_dict()) for idx, row in outras_notas.iterrows()]
                                
                                start_time = time.time()
                                for r in range(1, min(self.max_combinacoes, len(fast_notas) + 1)):
                                    if time.time() - start_time > self.timeout_combinacoes:
                                        break
                                    for comb in itertools.combinations(fast_notas, r):
                                        if time.time() - start_time > self.timeout_combinacoes:
                                            break
                                        if abs(sum(item[1] for item in comb) - valor_faltante) <= self.tolerancia_centavos:
                                            comb_encontrada = [item[2].copy() for item in comb]
                                            break
                                    if comb_encontrada or time.time() - start_time > self.timeout_combinacoes: 
                                        break
                            
                            if comb_encontrada:
                                nota_principal = row_a.to_dict().copy()
                                nota_principal['Baixas'] = row_b['Banco']
                                nota_principal['Data Baixa'] = row_b['Data']
                                nota_principal['Regra Aplicada'] = 'Regra 2.5 (Desmembrado Guiado por Histórico)'
                                resultados['3_Conciliado_Desmembrado'].append(nota_principal)
                                m_a.append(row_a['ID_Argos'])
                                
                                for np in comb_encontrada:
                                    np['Baixas'] = row_b['Banco']
                                    np['Data Baixa'] = row_b['Data']
                                    np['Regra Aplicada'] = 'Regra 2.5 (Desmembrado Guiado por Histórico)'
                                    resultados['3_Conciliado_Desmembrado'].append(np)
                                    m_a.append(np['ID_Argos'])
                                    
                                m_b.append(row_b['ID_Bank'])
                                break
                            
                            # Se não achou peças mas está dentro do desconto de R$ 15
                            elif abs(row_a['Valor'] - v_regex) <= self.tolerancia_desconto_pix:
                                nota = row_a.to_dict().copy()
                                nota['Baixas'] = row_b['Banco']
                                nota['Data Baixa'] = row_b['Data']
                                sinal = "+" if valor_faltante > 0 else ""
                                nota['Motivo Divergência'] = f'Desconto/Acréscimo no PIX ({sinal}R$ {valor_faltante})'
                                nota['Regra Aplicada'] = 'Regra 2 (Aproximação de Valor via Histórico)'
                                resultados['2_Conciliado_Via_Historico'].append(nota)
                                m_a.append(row_a['ID_Argos'])
                                m_b.append(row_b['ID_Bank'])
                                break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 0.5: CONCILIADO POR NOME (Nome do cliente no histórico do banco)
        # ==========================================
        m_a, m_b = [], []
        col_cliente = 'CLIENTES' if 'CLIENTES' in argos_pendentes.columns else 'Cliente'
        
        for i, row_a in argos_pendentes.iterrows():
            cliente_str = str(row_a.get(col_cliente, '')).strip().upper()
            if not cliente_str or cliente_str == 'CLIENTE NÃO INFORMADO':
                continue
                
            # Extrair palavras principais do nome do cliente (com mais de 2 letras)
            palavras_nome = [p for p in cliente_str.split() if len(p) > 2]
            if len(palavras_nome) < 2:
                continue
                
            candidatos = bank_pendentes[(abs(bank_pendentes['Valor'] - row_a['Valor']) <= self.tolerancia_centavos) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
            if not candidatos.empty:
                diff_dias = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                               pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                candidatos['diff_dias'] = diff_dias
                candidatos = candidatos.sort_values(by='diff_dias')
                
                for j, row_b in candidatos.iterrows():
                    hist_banco = str(row_b.get('Histórico', '')).upper()
                    
                    # Verifica se as duas primeiras palavras do nome estão no histórico
                    if palavras_nome[0] in hist_banco and palavras_nome[1] in hist_banco:
                        if pd.isna(row_b['diff_dias']) or row_b['diff_dias'] <= self.janela_dias:
                            nota = row_a.to_dict().copy()
                            nota['Baixas'] = row_b['Banco']
                            nota['Data Baixa'] = row_b['Data']
                            
                            diff_valor = round(row_b['Valor'] - row_a['Valor'], 2)
                            if diff_valor != 0:
                                sinal = "+" if diff_valor > 0 else ""
                                nota['Motivo Divergência'] = f'Aproximação de Centavos ({sinal}R$ {diff_valor})'
                                
                            nota['Regra Aplicada'] = 'Regra 0.5 (Match por Nome no Histórico)'
                            resultados['1_Conciliado_Perfeito'].append(nota)
                            m_a.append(row_a['ID_Argos'])
                            m_b.append(row_b['ID_Bank'])
                            break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 1.1: MATCH PERFEITO ÚNICO (Sem limite de datas)
        # ==========================================
        m_a, m_b = [], []
        
        argos_val_counts = argos_pendentes['Valor'].value_counts()
        bank_val_counts = bank_pendentes['Valor'].value_counts()
        
        unique_values = [v for v, c in argos_val_counts.items() if c == 1 and bank_val_counts.get(v) == 1]
        
        for v in unique_values:
            idx_a = argos_pendentes[argos_pendentes['Valor'] == v].index[0]
            idx_b = bank_pendentes[bank_pendentes['Valor'] == v].index[0]
            
            row_a = argos_pendentes.loc[idx_a]
            row_b = bank_pendentes.loc[idx_b]
            
            nota = row_a.to_dict().copy()
            nota['Baixas'] = row_b['Banco']
            nota['Data Baixa'] = row_b['Data']
            nota['Regra Aplicada'] = 'Regra 1.1 (Valor Único em Ambos)'
            resultados['1_Conciliado_Perfeito'].append(nota)
            
            m_a.append(row_a['ID_Argos'])
            m_b.append(row_b['ID_Bank'])
                
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 1: CONCILIADO PERFEITO (Valores Exatos com concorrência)
        # ==========================================
        m_a, m_b = [], []
        col_cliente = 'CLIENTE' if 'CLIENTE' in argos_pendentes.columns else 'Cliente'
        
        for i, row_a in argos_pendentes.iterrows():
            candidatos = bank_pendentes[(bank_pendentes['Valor'] == row_a['Valor']) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
            if not candidatos.empty:
                # Calcular name_score
                cliente_str = str(row_a.get(col_cliente, '')).strip().upper()
                palavras_nome = [p for p in cliente_str.split() if len(p) > 2]
                
                def calc_name_score(hist):
                    hist = str(hist).upper()
                    return sum(1 for p in palavras_nome if p in hist)
                
                candidatos['name_score'] = candidatos['Histórico'].apply(calc_name_score)
                
                # Tenta extrair data do histórico para ajudar na Regra 1
                texto_hist = str(row_a.get('Histórico', '') if 'Histórico' in row_a else row_a.get('OBS', ''))
                data_referencia_hist = None
                match_data = re.search(r'(?:dia\s*|data:\s*)?(\d{1,2})[/\.](\d{1,2})', texto_hist, re.IGNORECASE)
                if match_data:
                    dia, mes = match_data.groups()
                    try:
                        ano = pd.to_datetime(row_a['Data'], format='%d/%m/%Y').year
                        data_referencia_hist = pd.to_datetime(f"{int(dia):02d}/{int(mes):02d}/{ano}", format='%d/%m/%Y', errors='coerce')
                    except:
                        pass
                        
                diff_dias = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                               pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                
                if pd.notna(data_referencia_hist):
                    diff_dias_hist = abs((data_referencia_hist - pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                    candidatos['diff_dias'] = pd.concat([diff_dias, diff_dias_hist], axis=1).min(axis=1)
                else:
                    candidatos['diff_dias'] = diff_dias
                    
                # Ordena priorizando quem tem match no nome, e depois a data mais próxima
                candidatos = candidatos.sort_values(by=['name_score', 'diff_dias'], ascending=[False, True])
                
                for j, row_b in candidatos.iterrows():
                    # Se o name_score for 0 e houver mais de um candidato possível com esse valor na base do banco (competição real sem pistas)
                    # Não vamos chutar cegamente para não dar baixa na pessoa errada.
                    # Mas se só sobrou 1 candidato, a gente pode dar match (pois é o único restante com esse valor).
                    if row_b['name_score'] == 0 and len(candidatos) > 1:
                        hist_b = str(row_b.get('Histórico', '')).strip().upper()
                        genericos = [
                            'PIX RECEBIDO', 'CRED PIX CHAVE', 'PIX RECEBIDO DADOS CONTA', 
                            'DEPOSITO DINH LOTERICO', 'DEPOSITO DINHEIRO ATM'
                        ]
                        if hist_b not in genericos:
                            # Pula, pois é um chute perigoso (Ex: 2 Chrystians no banco e 1 Emporio no sistema)
                            continue
                        
                    if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= self.janela_dias:
                        nota = row_a.to_dict().copy()
                        nota['Baixas'] = row_b['Banco']
                        nota['Data Baixa'] = row_b['Data']
                        nota['Regra Aplicada'] = 'Regra 1 (Match Exato com/sem Nome)'
                        resultados['1_Conciliado_Perfeito'].append(nota)
                        m_a.append(row_a['ID_Argos'])
                        m_b.append(row_b['ID_Bank'])
                        break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 3: DESMEMBRADOS (Vários Argos -> 1 Banco)
        # ==========================================
        m_a_temp, m_b_temp = [], []
        
        for j, row_b in bank_pendentes.iterrows():
            if row_b['ID_Bank'] in m_b: continue
            
            valor_banco = row_b['Valor']
            data_banco = pd.to_datetime(row_b['Data'], format='%d/%m/%Y', errors='coerce')
            
            argos_candidatos = argos_pendentes[
                (~argos_pendentes['ID_Argos'].isin(m_a + m_a_temp)) & 
                (argos_pendentes['Valor'] <= valor_banco)
            ].copy()
            
            if pd.notna(data_banco):
                argos_candidatos['diff_dias'] = abs((pd.to_datetime(argos_candidatos['Data'], format='%d/%m/%Y', errors='coerce') - data_banco).dt.days)
                argos_candidatos = argos_candidatos[argos_candidatos['diff_dias'] <= self.janela_dias]
            
            if argos_candidatos.empty: continue
            
            # Prioridade 1: Combinações do mesmo cliente
            col_cliente = 'CLIENTES' if 'CLIENTES' in argos_candidatos.columns else 'Cliente'
            comb_encontrada = None
            
            for cliente, grupo in argos_candidatos.groupby(col_cliente):
                if len(grupo) < 2: continue
                fast_grupo = [(row['ID_Argos'], row['Valor'], row.to_dict()) for idx, row in grupo.iterrows()]
                
                start_time = time.time()
                for r in range(2, min(self.max_combinacoes, len(fast_grupo) + 1)):
                    if time.time() - start_time > self.timeout_combinacoes:
                        break
                    for comb in itertools.combinations(fast_grupo, r):
                        if time.time() - start_time > self.timeout_combinacoes:
                            break
                        if abs(sum(item[1] for item in comb) - valor_banco) <= self.tolerancia_centavos:
                            comb_encontrada = comb
                            break
                    if comb_encontrada or time.time() - start_time > self.timeout_combinacoes:
                        break
                if comb_encontrada: break
                            
            if comb_encontrada:
                for item in comb_encontrada:
                    nota = item[2].copy()
                    nota['Baixas'] = row_b['Banco']
                    nota['Data Baixa'] = row_b['Data']
                    nota['Regra Aplicada'] = 'Regra 3 (Combinação de Múltiplas Notas)'
                    resultados['3_Conciliado_Desmembrado'].append(nota)
                    m_a_temp.append(item[0])
                m_b_temp.append(row_b['ID_Bank'])
                
        m_a.extend(m_a_temp)
        m_b.extend(m_b_temp)
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 3.5: CONCILIADO POR APROXIMAÇÃO DE CENTAVOS (Último Recurso)
        # ==========================================
        m_a, m_b = [], []
        if not argos_pendentes.empty and not bank_pendentes.empty:
            for i, row_a in argos_pendentes.iterrows():
                candidatos = bank_pendentes[(abs(bank_pendentes['Valor'] - row_a['Valor']) > 0) & 
                                            (abs(bank_pendentes['Valor'] - row_a['Valor']) <= self.tolerancia_centavos) & 
                                            (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
                if not candidatos.empty:
                    candidatos['diff_dias'] = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                                                   pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                    candidatos = candidatos.sort_values(by='diff_dias')
                    
                    for j, row_b in candidatos.iterrows():
                        if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= self.janela_dias_centavos:
                            nota = row_a.to_dict().copy()
                            nota['Baixas'] = row_b['Banco']
                            nota['Data Baixa'] = row_b['Data']
                            diff_valor = round(row_b['Valor'] - row_a['Valor'], 2)
                            sinal = "+" if diff_valor > 0 else ""
                            nota['Motivo Divergência'] = f'Aproximação de Centavos ({sinal}R$ {diff_valor})'
                            nota['Regra Aplicada'] = 'Regra 3.5 (Aproximação de Centavos)'
                            resultados['1_Conciliado_Perfeito'].append(nota)
                            m_a.append(row_a['ID_Argos'])
                            m_b.append(row_b['ID_Bank'])
                            break
            remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 4: DIVERGÊNCIAS PENDENTES
        # ==========================================
        for i, row_a in argos_pendentes.iterrows():
            nota = row_a.to_dict().copy()
            nota['Data Baixa'] = nota['Data']
            nota['Data'] = ''
            nota['Motivo Divergência'] = 'Falta no Banco'
            nota['Regra Aplicada'] = 'N/A'
            resultados['5_Divergencias_Pendentes'].append(nota)
            
        for i, row_b in bank_pendentes.iterrows():
            col_hist_b = 'Histórico' if 'Histórico' in row_b else 'N/A'
            col_cliente = 'CLIENTES' if 'CLIENTES' in argos_pendentes.columns else 'Cliente'
            nota = {
                'Banco': row_b['Banco'],
                col_cliente: row_b.get(col_hist_b, 'N/A'),
                'Valor': row_b['Valor'],
                'Data': row_b['Data'],
                'Motivo Divergência': 'Sobrou no Banco / Faltou no Argos',
                'Regra Aplicada': 'N/A'
            }
            resultados['5_Divergencias_Pendentes'].append(nota)

        # ==========================================
        # FORMATAÇÃO FINAL DAS COLUNAS (Alinhado com o Template da Cliente)
        # Ordem Exata: Banco | Cliente | Valor | Data (Pgto) | Baixas (Banco Baixa) | Data Baixa | Histórico | Motivo | Regra Aplicada
        # ==========================================
        ordem_colunas = ['Banco', 'Cliente', 'Valor', 'Data', 'Baixas', 'Data Baixa', 'Histórico', 'Motivo Divergência', 'Regra Aplicada']
        
        for k in resultados.keys():
            df = pd.DataFrame(resultados[k])
            if df.empty:
                df = pd.DataFrame(columns=ordem_colunas)
            else:
                if 'ID_Argos' in df.columns: df = df.drop(columns=['ID_Argos'])
                if 'ID_Bank' in df.columns: df = df.drop(columns=['ID_Bank'])
                if 'CLIENTES' in df.columns: df = df.drop(columns=['CLIENTES'])
                
                # Garante que todas as colunas existem antes de reordenar
                for col in ordem_colunas:
                    if col not in df.columns:
                        df[col] = ''
                
                df = df[ordem_colunas]
            resultados[k] = df

        # ==========================================
        # VALIDAÇÃO DE SOMA (INTEGRIDADE FINANCEIRA)
        # ==========================================
        soma_input = self.df_argos['Valor'].sum() if not self.df_argos.empty else 0.0
        soma_output = 0.0
        
        for k in ['1_Conciliado_Perfeito', '2_Conciliado_Via_Historico', '3_Conciliado_Desmembrado']:
            if not resultados[k].empty:
                soma_output += resultados[k]['Valor'].sum()
                
        soma_divergencia = 0.0
        if not resultados['5_Divergencias_Pendentes'].empty:
            df_div = resultados['5_Divergencias_Pendentes']
            soma_divergencia = df_div[df_div['Motivo Divergência'] == 'Falta no Banco']['Valor'].sum()
            soma_output += soma_divergencia
            
        diff_integridade = abs(soma_input - soma_output)
        status_msg = "OK - Nenhum centavo perdido ou duplicado"
        if diff_integridade > TOLERANCIA_INTEGRIDADE:
            msg = f"CRÍTICO: Perda de integridade financeira! Input Argos: R$ {soma_input:.2f} | Output Argos: R$ {soma_output:.2f}"
            print(msg)
            warnings.warn(msg)
            status_msg = "ERRO (Perda/Duplicação identificada)"

        df_resumo = pd.DataFrame([
            {"Métrica": "Total Input Argos", "Valor": round(soma_input, 2)},
            {"Métrica": "Total Output Conciliado", "Valor": round(soma_output - soma_divergencia, 2)},
            {"Métrica": "Total Output Divergências (Falta Banco)", "Valor": round(soma_divergencia, 2)},
            {"Métrica": "Total Output (Conciliado + Divergências)", "Valor": round(soma_output, 2)},
            {"Métrica": "Diferença (Perda/Duplicação)", "Valor": round(diff_integridade, 2)},
            {"Métrica": "Status da Integridade", "Valor": status_msg},
        ])

        # ==========================================
        # ABA RESUMO EXECUTIVO (PRIMEIRA ABA NO EXCEL)
        # ==========================================
        qtd_argos = len(self.df_argos) if not self.df_argos.empty else 0
        qtd_bank = len(self.df_bank) if not self.df_bank.empty else 0

        qtd_perfeitos = len(resultados['1_Conciliado_Perfeito'])
        qtd_historico = len(resultados['2_Conciliado_Via_Historico'])
        qtd_desmembrado = len(resultados['3_Conciliado_Desmembrado'])
        qtd_saidas_estornos = len(resultados['4_Saidas_Estornos'])
        qtd_divergencias = len(resultados['5_Divergencias_Pendentes'])

        soma_bank = round(float(self.df_bank['Valor'].sum()), 2) if not self.df_bank.empty and 'Valor' in self.df_bank.columns else 0.0
        soma_perfeitos = round(float(resultados['1_Conciliado_Perfeito']['Valor'].sum()), 2) if not resultados['1_Conciliado_Perfeito'].empty else 0.0
        soma_historico = round(float(resultados['2_Conciliado_Via_Historico']['Valor'].sum()), 2) if not resultados['2_Conciliado_Via_Historico'].empty else 0.0
        soma_desmembrado = round(float(resultados['3_Conciliado_Desmembrado']['Valor'].sum()), 2) if not resultados['3_Conciliado_Desmembrado'].empty else 0.0
        soma_saidas_estornos = round(float(resultados['4_Saidas_Estornos']['Valor'].sum()), 2) if not resultados['4_Saidas_Estornos'].empty else 0.0
        soma_divergencias = round(float(resultados['5_Divergencias_Pendentes']['Valor'].sum()), 2) if not resultados['5_Divergencias_Pendentes'].empty else 0.0

        qtd_conciliados = qtd_perfeitos + qtd_historico + qtd_desmembrado
        soma_conciliados = round(soma_perfeitos + soma_historico + soma_desmembrado, 2)

        qtd_sucesso_total = qtd_conciliados + qtd_saidas_estornos
        soma_sucesso_total = round(soma_conciliados + soma_saidas_estornos, 2)

        total_itens_processados = qtd_sucesso_total + qtd_divergencias

        if total_itens_processados > 0:
            taxa_sucesso_qtd = (qtd_sucesso_total / total_itens_processados) * 100.0
        else:
            taxa_sucesso_qtd = 100.0 if qtd_argos == 0 else 0.0

        base_financeira = soma_sucesso_total + soma_divergencias
        if base_financeira > 0:
            taxa_sucesso_valor = (soma_sucesso_total / base_financeira) * 100.0
        elif soma_input > 0:
            taxa_sucesso_valor = (soma_sucesso_total / soma_input) * 100.0
        else:
            taxa_sucesso_valor = 100.0 if qtd_argos == 0 else 0.0

        timestamp_execucao = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        assinatura_sha256 = self._calcular_assinatura_digital(
            self.df_argos, self.df_bank, resultados, timestamp_execucao
        )

        status_conciliacao = (
            "Concluída com Sucesso" if diff_integridade <= TOLERANCIA_INTEGRIDADE and qtd_divergencias == 0
            else "Concluída com Divergências Pendentes" if diff_integridade <= TOLERANCIA_INTEGRIDADE
            else "Atenção: Diferença de Integridade Detectada"
        )

        df_executivo = pd.DataFrame([
            {"Métrica": "Data/Hora de Processamento", "Valor": timestamp_execucao},
            {"Métrica": "Status da Conciliação", "Valor": status_conciliacao},
            {"Métrica": "Taxa de Sucesso (Registros)", "Valor": f"{taxa_sucesso_qtd:.2f}%"},
            {"Métrica": "Taxa de Sucesso (Financeira)", "Valor": f"{taxa_sucesso_valor:.2f}%"},
            {"Métrica": "Total de Registros Argos", "Valor": int(qtd_argos)},
            {"Métrica": "Total Volume Argos (R$)", "Valor": round(soma_input, 2)},
            {"Métrica": "Total de Registros Banco", "Valor": int(qtd_bank)},
            {"Métrica": "Total Volume Banco (R$)", "Valor": round(soma_bank, 2)},
            {"Métrica": "Total Conciliado (Registros)", "Valor": int(qtd_conciliados)},
            {"Métrica": "Total Conciliado (R$)", "Valor": round(soma_conciliados, 2)},
            {"Métrica": "  - Match Perfeito (Qtd)", "Valor": int(qtd_perfeitos)},
            {"Métrica": "  - Match Perfeito (R$)", "Valor": round(soma_perfeitos, 2)},
            {"Métrica": "  - Via Histórico (Qtd)", "Valor": int(qtd_historico)},
            {"Métrica": "  - Via Histórico (R$)", "Valor": round(soma_historico, 2)},
            {"Métrica": "  - Desmembrado (Qtd)", "Valor": int(qtd_desmembrado)},
            {"Métrica": "  - Desmembrado (R$)", "Valor": round(soma_desmembrado, 2)},
            {"Métrica": "Total Saídas / Estornos (Registros)", "Valor": int(qtd_saidas_estornos)},
            {"Métrica": "Total Saídas / Estornos (R$)", "Valor": round(soma_saidas_estornos, 2)},
            {"Métrica": "Total Divergências Pendentes (Registros)", "Valor": int(qtd_divergencias)},
            {"Métrica": "Total Divergências Pendentes (R$)", "Valor": round(soma_divergencias, 2)},
            {"Métrica": "Status de Integridade Financeira", "Valor": status_msg},
            {"Métrica": "Assinatura Digital (Hash SHA-256)", "Valor": assinatura_sha256},
        ])

        return {
            '0_Resumo_Executivo': df_executivo,
            '1_Conciliado_Perfeito': resultados['1_Conciliado_Perfeito'],
            '2_Conciliado_Via_Historico': resultados['2_Conciliado_Via_Historico'],
            '3_Conciliado_Desmembrado': resultados['3_Conciliado_Desmembrado'],
            '4_Saidas_Estornos': resultados['4_Saidas_Estornos'],
            '5_Divergencias_Pendentes': resultados['5_Divergencias_Pendentes'],
            '6_Resumo_Integridade': df_resumo,
        }




class ExcelReporter:
    LIMITE_DIAS_ALERTA_TEMPORAL = config.DIAS_ALERTA_TEMPORAL

    @staticmethod
    def generate_report(data_sheets: dict, output_path: str, dias_alerta_temporal: int = config.DIAS_ALERTA_TEMPORAL):
        formatted_sheets = {}
        
        mapa_colunas = {
            'BANCO': 'BANCO',
            'Banco': 'BANCO',
            'CLIENTES': 'CLIENTE',
            'Cliente': 'CLIENTE',
            'VALOR': 'VALOR DA BAIXA',
            'Valor': 'VALOR DA BAIXA',
            'DATA PAGAMENTO': 'DATA DO PAGAMENTO',
            'Data': 'DATA DO PAGAMENTO',
            'BAIXAS': 'BANCO DA BAIXA',
            'Baixas': 'BANCO DA BAIXA',
            'DATA BAIXA': 'DATA DA BAIXA',
            'Data Baixa': 'DATA DA BAIXA',
            'OBS': 'HISTÓRICO',
            'Histórico': 'HISTÓRICO',
            'Motivo Divergência': 'MOTIVO DIVERGÊNCIA',
            'Regra Aplicada': 'REGRA APLICADA'
        }
        
        ordem_desejada = ['BANCO', 'CLIENTE', 'VALOR DA BAIXA', 'DATA DO PAGAMENTO', 'BANCO DA BAIXA', 'DATA DA BAIXA', 'HISTÓRICO', 'MOTIVO DIVERGÊNCIA', 'REGRA APLICADA']
        
        for sheet_name, df in data_sheets.items():
            novo_nome_aba = sheet_name.replace('_', ' ')
            if 'Resumo Integridade' in novo_nome_aba or 'Resumo Executivo' in novo_nome_aba:
                formatted_sheets[novo_nome_aba] = df
                continue
                
            if not df.empty:
                df = df.rename(columns=mapa_colunas)
                colunas_finais = [c for c in ordem_desejada if c in df.columns]
                df = df[colunas_finais].copy()
                
                # Alerta temporal para diferença de datas entre pagamento e baixa
                observacoes = []
                has_pgto = 'DATA DO PAGAMENTO' in df.columns
                has_baixa = 'DATA DA BAIXA' in df.columns
                
                if has_pgto and has_baixa:
                    dt_pgto = pd.to_datetime(df['DATA DO PAGAMENTO'], dayfirst=True, errors='coerce')
                    dt_baixa = pd.to_datetime(df['DATA DA BAIXA'], dayfirst=True, errors='coerce')
                    diff_dias = (dt_pgto - dt_baixa).abs().dt.days
                    
                    for diff in diff_dias:
                        if pd.notna(diff) and diff > dias_alerta_temporal:
                            observacoes.append(f"⚠️ Alerta Temporal: Diferença de {int(diff)} dias entre pagamento e baixa")
                        else:
                            observacoes.append('')
                else:
                    observacoes = [''] * len(df)
                    
                df['OBSERVAÇÃO'] = observacoes
            else:
                df = pd.DataFrame(columns=[c for c in ordem_desejada if c != 'MOTIVO DIVERGÊNCIA'] + ['OBSERVAÇÃO'])
            
            formatted_sheets[novo_nome_aba] = df

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in formatted_sheets.items():
                if df.empty:
                    pd.DataFrame({'Aviso': ['Nenhum registo encontrado nesta categoria.']}).to_excel(writer, sheet_name=sheet_name, index=False)
                    continue
                
                for col in df.columns:
                    if 'DATA' in col.upper() and 'RESUMO' not in sheet_name.upper():
                        df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y').replace('NaT', '')

                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
                worksheet = writer.sheets[sheet_name]

                header_fill = PatternFill(start_color='266C40', end_color='266C40', fill_type='solid')
                fill_verde_claro = PatternFill(start_color='719C82', end_color='719C82', fill_type='solid')
                fill_branco = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
                fill_alerta = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
                
                font_branca_bold = Font(color='FFFFFF', bold=True)
                font_preta_bold = Font(color='000000', bold=True)
                font_alerta = Font(color='8A5300', bold=True)
                
                borda_fina = Border(
                    left=Side(border_style='thin', color='356A1C'),
                    right=Side(border_style='thin', color='356A1C'),
                    top=Side(border_style='thin', color='356A1C'),
                    bottom=Side(border_style='thin', color='356A1C')
                )
                borda_alerta = Border(
                    left=Side(border_style='thin', color='D69E2E'),
                    right=Side(border_style='thin', color='D69E2E'),
                    top=Side(border_style='thin', color='D69E2E'),
                    bottom=Side(border_style='thin', color='D69E2E')
                )

                col_indices = {str(worksheet.cell(row=1, column=i).value).upper(): i for i in range(1, worksheet.max_column + 1)}
                max_col = worksheet.max_column

                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = font_branca_bold
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                    cell.border = borda_fina

                for row in range(2, worksheet.max_row + 1):
                    is_zebra = (row % 2 == 0)

                    obs_cell_val = ""
                    if 'OBSERVAÇÃO' in col_indices:
                        obs_cell_val = str(worksheet.cell(row=row, column=col_indices['OBSERVAÇÃO']).value or '')
                    tem_alerta_temporal = "Alerta Temporal" in obs_cell_val

                    for col_name, col_idx in col_indices.items():
                        cell = worksheet.cell(row=row, column=col_idx)
                        
                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                        
                        if sheet_name == '4 Saidas Estornos' and col_name == 'VALOR DA BAIXA':
                            cell.font = Font(color='FF0000', bold=True)
                        elif tem_alerta_temporal and col_name in ['DATA DO PAGAMENTO', 'DATA DA BAIXA', 'OBSERVAÇÃO']:
                            cell.font = font_alerta
                        else:
                            cell.font = font_preta_bold
                            
                        if tem_alerta_temporal and col_name in ['DATA DO PAGAMENTO', 'DATA DA BAIXA', 'OBSERVAÇÃO']:
                            cell.border = borda_alerta
                        else:
                            cell.border = borda_fina

                        if tem_alerta_temporal and col_name in ['DATA DO PAGAMENTO', 'DATA DA BAIXA', 'OBSERVAÇÃO']:
                            cell.fill = fill_alerta
                        elif col_name in ['BANCO', 'BANCO DA BAIXA']:
                            cell.fill = fill_verde_claro
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                        else:
                            if is_zebra:
                                cell.fill = fill_verde_claro
                            else:
                                cell.fill = fill_branco

                        is_resumo_executivo = 'Resumo Executivo' in sheet_name
                        is_resumo_integridade = 'Resumo Integridade' in sheet_name

                        if col_name == 'VALOR DA BAIXA':
                            cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                            if cell.value is not None:
                                try:
                                    cell.value = float(cell.value)
                                    cell.number_format = 'R$ #,##0.00'
                                except:
                                    pass
                        elif is_resumo_integridade and col_name == 'VALOR':
                            if cell.value is not None:
                                try:
                                    cell.value = float(cell.value)
                                    cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                                    cell.number_format = 'R$ #,##0.00'
                                except:
                                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
                        elif is_resumo_executivo and col_name == 'VALOR':
                            metrica_val = str(worksheet.cell(row=row, column=col_indices.get('MÉTRICA', 1)).value or '')
                            if '(R$)' in metrica_val or 'Volume' in metrica_val:
                                try:
                                    cell.value = float(cell.value)
                                    cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                                    cell.number_format = 'R$ #,##0.00'
                                except:
                                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
                            elif '(Qtd)' in metrica_val or 'Registros' in metrica_val:
                                try:
                                    cell.value = int(cell.value)
                                    cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                                    cell.number_format = '#,##0'
                                except:
                                    cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                            elif 'Taxa de Sucesso' in metrica_val:
                                cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                            else:
                                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)

                for col_idx in range(1, max_col + 1):
                    max_length = 0
                    col_letter = get_column_letter(col_idx)
                    col_name_val = worksheet.cell(row=1, column=col_idx).value
                    if not col_name_val: continue
                    col_name = str(col_name_val).upper()
                    
                    max_length = len(col_name) + 3 
                    
                    for row_idx in range(1, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        try:
                            if cell.value:
                                linhas = str(cell.value).split('\n')
                                for linha in linhas:
                                    max_length = max(max_length, len(linha))
                        except:
                            pass
                    
                    if col_name in ['BANCO', 'BANCO DA BAIXA']:
                        worksheet.column_dimensions[col_letter].width = max(max_length, 20)
                    elif col_name == 'VALOR DA BAIXA' or col_name == 'VALOR':
                        worksheet.column_dimensions[col_letter].width = max(max_length + 6, 20)
                    elif 'DATA' in col_name:
                        worksheet.column_dimensions[col_letter].width = max(max_length + 2, 23)
                    elif col_name == 'OBSERVAÇÃO':
                        worksheet.column_dimensions[col_letter].width = max(max_length + 2, 40)
                    elif col_name == 'MÉTRICA':
                        worksheet.column_dimensions[col_letter].width = max(max_length + 2, 45)
                    else:
                        worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)

        with open(output_path, "rb") as f:
            file_hash = hashlib.sha256(f.read()).hexdigest()
        return file_hash


class AnalyticsService:
    """
    Serviço de análise de dados históricos e inteligência de conciliação (Dashboard Analytics).
    Permite filtragem por banco, período e tipo de conciliação, gerando métricas agregadas e séries temporais.
    """

    @staticmethod
    def extrair_bancos_de_periodo(periodo_str: str) -> list:
        """
        Extrai a lista de bancos de uma string de período no padrão '... | Bancos: B1, B2'.
        """
        if not periodo_str:
            return []
        periodo_s = str(periodo_str)
        if "| Bancos:" not in periodo_s:
            return []
        try:
            raw_bancos = periodo_s.split("| Bancos:")[1].strip()
            return [b.strip() for b in raw_bancos.split(",") if b.strip()]
        except Exception:
            return []

    @staticmethod
    def filtrar_registros(
        registros: list,
        banco: str = None,
        dias: int = None,
        tipo_sucesso: str = None,
        busca: str = None,
    ) -> list:
        """
        Filtra uma lista de registros de histórico com base nos critérios fornecidos.
        """
        if not registros:
            return []

        resultado = list(registros)

        # 1. Filtro por Banco
        if banco and str(banco).strip().lower() not in ["todos", "todos os bancos", ""]:
            banco_filtro = str(banco).strip().upper()
            filtrados = []
            for r in resultado:
                periodo_val = str(r.get("periodo", "") or "").upper()
                banco_val = str(r.get("banco", "") or "").upper()
                bancos_lista = [b.upper() for b in AnalyticsService.extrair_bancos_de_periodo(periodo_val)]

                if banco_filtro in bancos_lista or banco_filtro in periodo_val or banco_filtro in banco_val:
                    filtrados.append(r)
                elif banco_filtro in ["OUTROS", "GERAL", "DESCONHECIDO"] and not bancos_lista:
                    filtrados.append(r)
            resultado = filtrados

        # 2. Filtro por Período em Dias
        if dias is not None and dias > 0:
            from datetime import datetime, timedelta
            limite = datetime.now() - timedelta(days=dias)
            filtrados = []
            for r in resultado:
                dt_val = r.get("data_processamento")
                if dt_val:
                    try:
                        if isinstance(dt_val, datetime):
                            dt = dt_val
                        else:
                            clean_dt_str = str(dt_val).replace("T", " ").split(".")[0]
                            dt = datetime.fromisoformat(clean_dt_str)
                        if dt >= limite:
                            filtrados.append(r)
                    except Exception:
                        filtrados.append(r)
            resultado = filtrados

        # 3. Filtro por Tipo de Sucesso / Desempenho
        if tipo_sucesso and str(tipo_sucesso).strip().lower() not in ["todos", ""]:
            tipo = str(tipo_sucesso).strip().lower()
            filtrados = []
            for r in resultado:
                taxa = float(r.get("taxa_sucesso", 0.0) or 0.0)
                divs = int(r.get("divergencias", 0) or 0)

                if tipo in ["alta", "alto", ">=80"]:
                    if taxa >= 80.0:
                        filtrados.append(r)
                elif tipo in ["media", "medio", "50-79"]:
                    if 50.0 <= taxa < 80.0:
                        filtrados.append(r)
                elif tipo in ["baixa", "baixo", "<50"]:
                    if taxa < 50.0:
                        filtrados.append(r)
                elif tipo in ["perfeito", "100"]:
                    if taxa >= 99.99 and divs == 0:
                        filtrados.append(r)
                elif tipo in ["com_divergencias", "divergencias"]:
                    if divs > 0:
                        filtrados.append(r)
                else:
                    filtrados.append(r)
            resultado = filtrados

        # 4. Filtro por Busca Textual Livre
        if busca and str(busca).strip():
            termo = str(busca).strip().lower()
            filtrados = []
            for r in resultado:
                periodo_str = str(r.get("periodo", "") or "").lower()
                anotacao_str = str(r.get("anotacao", "") or "").lower()
                id_str = str(r.get("id", "") or "").lower()
                if termo in periodo_str or termo in anotacao_str or termo == id_str:
                    filtrados.append(r)
            resultado = filtrados

        return resultado

    @staticmethod
    def calcular_metricas(registros: list) -> dict:
        """
        Calcula as métricas globais consolidadas e KPIs a partir dos registros.
        """
        if not registros:
            return {
                "total_conciliacoes": 0,
                "taxa_media_sucesso": 0.0,
                "taxa_media_aritmetica": 0.0,
                "total_transacoes": 0,
                "total_conciliados": 0,
                "total_divergencias": 0,
                "distribuicao_tipos": {
                    "perfeitos": 0,
                    "historico": 0,
                    "desmembrados": 0,
                    "saidas_estornos": 0,
                    "divergencias": 0,
                },
                "percentuais_tipos": {
                    "perfeitos": 0.0,
                    "historico": 0.0,
                    "desmembrados": 0.0,
                    "saidas_estornos": 0.0,
                    "divergencias": 0.0,
                },
                "bancos_detectados": [],
            }

        total_conciliacoes = len(registros)
        total_perfeitos = sum(int(r.get("perfeitos", 0) or 0) for r in registros)
        total_historico = sum(int(r.get("historico", 0) or 0) for r in registros)
        total_desmembrados = sum(int(r.get("desmembrados", 0) or 0) for r in registros)
        total_saidas_estornos = sum(int(r.get("saidas_estornos", 0) or 0) for r in registros)
        total_divergencias = sum(int(r.get("divergencias", 0) or 0) for r in registros)

        total_conciliados = (
            total_perfeitos + total_historico + total_desmembrados + total_saidas_estornos
        )
        total_transacoes = total_conciliados + total_divergencias

        # Taxa média ponderada
        taxa_media_ponderada = (
            round((total_conciliados / total_transacoes) * 100, 2)
            if total_transacoes > 0
            else 0.0
        )

        # Taxa média aritmética simples
        taxas = [float(r.get("taxa_sucesso", 0.0) or 0.0) for r in registros]
        taxa_media_aritmetica = round(sum(taxas) / len(taxas), 2) if taxas else 0.0

        # Percentuais de cada tipo
        pct = lambda val: round((val / total_transacoes * 100), 2) if total_transacoes > 0 else 0.0

        percentuais = {
            "perfeitos": pct(total_perfeitos),
            "historico": pct(total_historico),
            "desmembrados": pct(total_desmembrados),
            "saidas_estornos": pct(total_saidas_estornos),
            "divergencias": pct(total_divergencias),
        }

        # Bancos detectados
        bancos_set = set()
        for r in registros:
            p = str(r.get("periodo", "") or "")
            for b in AnalyticsService.extrair_bancos_de_periodo(p):
                bancos_set.add(b)
            if r.get("banco"):
                bancos_set.add(str(r.get("banco")).strip())

        return {
            "total_conciliacoes": total_conciliacoes,
            "taxa_media_sucesso": taxa_media_ponderada,
            "taxa_media_aritmetica": taxa_media_aritmetica,
            "total_transacoes": total_transacoes,
            "total_conciliados": total_conciliados,
            "total_divergencias": total_divergencias,
            "distribuicao_tipos": {
                "perfeitos": total_perfeitos,
                "historico": total_historico,
                "desmembrados": total_desmembrados,
                "saidas_estornos": total_saidas_estornos,
                "divergencias": total_divergencias,
            },
            "percentuais_tipos": percentuais,
            "bancos_detectados": sorted(list(bancos_set)),
        }

    @staticmethod
    def calcular_series_temporal(registros: list, max_pontos: int = 15) -> list:
        """
        Prepara a série temporal cronológica para plotagem em gráficos de linha e barras.
        """
        if not registros:
            return []

        def _get_dt(r):
            dt = r.get("data_processamento", "")
            return str(dt)

        ordenados = sorted(registros, key=_get_dt)
        if max_pontos and len(ordenados) > max_pontos:
            ordenados = ordenados[-max_pontos:]

        series = []
        for r in ordenados:
            periodo_raw = str(r.get("periodo", "") or "Desconhecido")
            periodo_curto = periodo_raw.split(" | ")[0]
            bancos = AnalyticsService.extrair_bancos_de_periodo(periodo_raw)

            perfeitos = int(r.get("perfeitos", 0) or 0)
            historico = int(r.get("historico", 0) or 0)
            desmembrados = int(r.get("desmembrados", 0) or 0)
            saidas = int(r.get("saidas_estornos", 0) or 0)
            divergencias = int(r.get("divergencias", 0) or 0)
            conciliados = perfeitos + historico + desmembrados + saidas

            series.append({
                "id": r.get("id"),
                "data_processamento": str(r.get("data_processamento", "")),
                "periodo_label": periodo_curto,
                "periodo_completo": periodo_raw,
                "bancos": bancos,
                "taxa_sucesso": round(float(r.get("taxa_sucesso", 0.0) or 0.0), 2),
                "conciliados": conciliados,
                "divergencias": divergencias,
                "total": conciliados + divergencias,
                "perfeitos": perfeitos,
                "historico": historico,
                "desmembrados": desmembrados,
                "saidas_estornos": saidas,
            })

        return series

    @staticmethod
    def gerar_dashboard_analytics(
        registros: list,
        banco: str = None,
        dias: int = None,
        tipo_sucesso: str = None,
        busca: str = None,
    ) -> dict:
        """
        Gera o pacote consolidado de inteligência para o Dashboard Analytics.
        """
        todos = list(registros or [])
        filtrados = AnalyticsService.filtrar_registros(
            registros=todos,
            banco=banco,
            dias=dias,
            tipo_sucesso=tipo_sucesso,
            busca=busca,
        )

        metricas = AnalyticsService.calcular_metricas(filtrados)
        series = AnalyticsService.calcular_series_temporal(filtrados)

        todos_bancos = set()
        for r in todos:
            p = str(r.get("periodo", "") or "")
            for b in AnalyticsService.extrair_bancos_de_periodo(p):
                todos_bancos.add(b)
            if r.get("banco"):
                todos_bancos.add(str(r.get("banco")).strip())

        return {
            "filtros_aplicados": {
                "banco": banco,
                "dias": dias,
                "tipo_sucesso": tipo_sucesso,
                "busca": busca,
            },
            "total_geral_historico": len(todos),
            "total_filtrado": len(filtrados),
            "total_registros_brutos": len(todos),
            "total_registros_filtrados": len(filtrados),
            "metricas": metricas,
            "distribuicao": metricas["distribuicao_tipos"],
            "percentuais": metricas["percentuais_tipos"],
            "series_temporal": series,
            "bancos_disponiveis": sorted(list(todos_bancos)),
        }


if __name__ == "__main__":
    pass