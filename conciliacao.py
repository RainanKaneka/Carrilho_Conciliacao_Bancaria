import argparse
import warnings
import pandas as pd
import re
import itertools
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

def parse_currency(val):
    """Função blindada para converter qualquer formato de dinheiro para número"""
    if pd.isna(val): return None
    if isinstance(val, (int, float)): return float(val)
    val = str(val).upper().replace('R$', '').strip()
    if '.' in val and ',' in val:
        val = val.replace('.', '').replace(',', '.')
    elif ',' in val:
        val = val.replace(',', '.')
    try:
        return float(val)
    except:
        return None


class DataCleaner:
    @staticmethod
    def clean_argos(file_path: str) -> pd.DataFrame:
        try:
            import pandas as pd
            df = pd.read_excel(file_path, engine='openpyxl')
            
            header_idx = None
            for idx, row in df.iterrows():
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
            
            df = df.rename(columns=col_mapping)
            df = df.loc[:, ~df.columns.duplicated()].copy()
            cols_to_keep = [c for c in ['Banco', 'Cliente', 'Valor', 'Data', 'Histórico'] if c in df.columns]
            df = df[cols_to_keep]
            
            if 'Valor' not in df.columns: return pd.DataFrame()
            if 'Data' not in df.columns: df['Data'] = ''
            if 'Histórico' not in df.columns: df['Histórico'] = ''
            if 'Cliente' not in df.columns: df['Cliente'] = 'CLIENTE NÃO INFORMADO'
            
            df = df.dropna(subset=['Valor'])
            
            def parse_currency(val):
                if pd.isna(val): return None
                if isinstance(val, (int, float)): return float(val)
                val = str(val).upper().replace('R$', '').strip()
                if '.' in val and ',' in val:
                    val = val.replace('.', '').replace(',', '.')
                elif ',' in val:
                    val = val.replace(',', '.')
                try: return float(val)
                except: return None

            # NOVO PARSER DE DATA BLINDADO (Padrão BR)
            def parse_date_br(val):
                if pd.isna(val) or str(val).strip() == '': return ''
                try: return pd.to_datetime(val, dayfirst=True).strftime('%d/%m/%Y')
                except: return str(val)

            df['Valor'] = df['Valor'].apply(parse_currency)
            df['Data'] = df['Data'].apply(parse_date_br)
            df['Histórico'] = df['Histórico'].fillna('')
            
            return df.dropna(subset=['Valor'])
        except Exception as e:
            print(f"Erro ao ler Argos {file_path}: {e}")
            return pd.DataFrame()

    @staticmethod
    def clean_bank(file_path: str) -> pd.DataFrame:
        try:
            import pandas as pd
            df = pd.read_excel(file_path, engine='openpyxl')
            
            header_idx = None
            for idx, row in df.iterrows():
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
            
            df = df.rename(columns=col_mapping)
            df = df.loc[:, ~df.columns.duplicated()].copy()
            cols_to_keep = [c for c in ['Data', 'Histórico', 'Valor'] if c in df.columns]
            df = df[cols_to_keep]
            
            if 'Valor' not in df.columns: return pd.DataFrame()
            if 'Histórico' not in df.columns: df['Histórico'] = ''
            
            df = df.dropna(subset=['Valor'])
            
            def parse_currency(val):
                if pd.isna(val): return None
                if isinstance(val, (int, float)): return float(val)
                val = str(val).upper().replace('R$', '').strip()
                if '.' in val and ',' in val:
                    val = val.replace('.', '').replace(',', '.')
                elif ',' in val:
                    val = val.replace(',', '.')
                try: return float(val)
                except: return None

            def parse_date_br(val):
                if pd.isna(val) or str(val).strip() == '': return ''
                try: return pd.to_datetime(val, dayfirst=True).strftime('%d/%m/%Y')
                except: return str(val)

            df['Valor'] = df['Valor'].apply(parse_currency)
            df['Data'] = df['Data'].apply(parse_date_br)
            df = df.dropna(subset=['Valor'])
            df = df[df['Valor'] > 0]
            df['Histórico'] = df['Histórico'].fillna('')
            
            nome_arquivo = str(file_path).lower()
            banco_nome = "BANCO DESCONHECIDO"
            if 'caixa' in nome_arquivo: banco_nome = 'CAIXA ECONOMICA'
            elif 'banese' in nome_arquivo: banco_nome = 'BANESE'
            df['Banco'] = banco_nome
            
            return df
        except Exception as e:
            print(f"Erro ao ler Banco {file_path}: {e}")
            return pd.DataFrame()



class ReconciliationEngine:
    def __init__(self, df_argos, df_bank):
        import pandas as pd
        
        self.df_argos = df_argos.copy() if isinstance(df_argos, pd.DataFrame) else pd.DataFrame()
        self.df_bank = df_bank.copy() if isinstance(df_bank, pd.DataFrame) else pd.DataFrame()

        # BLINDAGEM: Garante que todos os valores monetários são floats matemáticos válidos
        def safe_float(val):
            if pd.isna(val): return 0.0
            if isinstance(val, (int, float)): return float(val)
            val = str(val).upper().replace('R$', '').strip()
            if '.' in val and ',' in val:
                val = val.replace('.', '').replace(',', '.')
            elif ',' in val:
                val = val.replace(',', '.')
            try:
                return float(val)
            except:
                return 0.0

        if not self.df_argos.empty and 'Valor' in self.df_argos.columns:
            self.df_argos['Valor'] = self.df_argos['Valor'].apply(safe_float)
            
        if not self.df_bank.empty and 'Valor' in self.df_bank.columns:
            self.df_bank['Valor'] = self.df_bank['Valor'].apply(safe_float)
            # Filtra zeros ou erros de conversão
            self.df_bank = self.df_bank[self.df_bank['Valor'] > 0]
            self.df_argos = self.df_argos[self.df_argos['Valor'] > 0]

        if not self.df_argos.empty:
            self.df_argos['ID_Argos'] = range(len(self.df_argos))
        if not self.df_bank.empty:
            self.df_bank['ID_Bank'] = range(len(self.df_bank))

    def execute_pipeline(self):
        import pandas as pd
        import re
        import itertools
        
        resultados = {
            '1_Conciliado_Perfeito': [],
            '2_Conciliado_Via_Historico': [],
            '3_Conciliado_Desmembrado': [],
            '4_Divergencias_Pendentes': []
        }

        if self.df_argos.empty or self.df_bank.empty:
            for k in resultados.keys():
                resultados[k] = pd.DataFrame(columns=['Banco', 'Cliente', 'Valor', 'Data', 'Histórico', 'Baixas', 'Data Baixa', 'Motivo Divergência'])
            return resultados

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
        regex = re.compile(r'(?:PIX.*?|VALOR DE|COMPROVANTE.*?|DE\s*)\s*R?\$?\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)', re.IGNORECASE)
        m_a, m_b = [], []
        
        for i, row_a in argos_pendentes.iterrows():
            if row_a['ID_Argos'] in m_a: continue
            
            col_hist = 'OBS' if 'OBS' in row_a else 'Histórico'
            match = regex.search(str(row_a.get(col_hist, '')))
            
            if match:
                val_str = match.group(1).replace('.', '').replace(',', '.')
                try: 
                    v_regex = float(val_str)
                except: 
                    continue
                
                candidatos = bank_pendentes[(bank_pendentes['Valor'] == v_regex) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
                if not candidatos.empty:
                    candidatos['diff_dias'] = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                                                   pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                    candidatos = candidatos.sort_values(by='diff_dias')
                    
                    for j, row_b in candidatos.iterrows():
                        if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= 5: 
                            valor_faltante = round(v_regex - row_a['Valor'], 2)
                            comb_encontrada = []
                            
                            # Tenta Desmembrar (OTIMIZADO)
                            if valor_faltante > 0.05:
                                # FILTRO CRÍTICO: Limita a busca a notas que "cabem" no espaço vazio, evitando loop O(N³)
                                outras_notas = argos_pendentes[
                                    (~argos_pendentes['ID_Argos'].isin(m_a)) & 
                                    (argos_pendentes['ID_Argos'] != row_a['ID_Argos']) &
                                    (argos_pendentes['Valor'] <= valor_faltante + 0.05)
                                ]
                                fast_notas = [(row['ID_Argos'], row['Valor'], row.to_dict()) for idx, row in outras_notas.iterrows()]
                                
                                for r in range(1, min(4, len(fast_notas) + 1)):
                                    for comb in itertools.combinations(fast_notas, r):
                                        if abs(sum(item[1] for item in comb) - valor_faltante) < 0.05:
                                            comb_encontrada = [item[2].copy() for item in comb]
                                            break
                                    if comb_encontrada: break
                            
                            if comb_encontrada:
                                nota_principal = row_a.to_dict().copy()
                                nota_principal['Baixas'] = row_b['Banco']
                                nota_principal['Data Baixa'] = row_b['Data']
                                resultados['3_Conciliado_Desmembrado'].append(nota_principal)
                                m_a.append(row_a['ID_Argos'])
                                
                                for np in comb_encontrada:
                                    np['Baixas'] = row_b['Banco']
                                    np['Data Baixa'] = row_b['Data']
                                    resultados['3_Conciliado_Desmembrado'].append(np)
                                    m_a.append(np['ID_Argos'])
                                    
                                m_b.append(row_b['ID_Bank'])
                                break
                            
                            # Se não achou peças mas está dentro do desconto de R$ 15
                            elif abs(row_a['Valor'] - v_regex) <= 15.0:
                                nota = row_a.to_dict().copy()
                                nota['Baixas'] = row_b['Banco']
                                nota['Data Baixa'] = row_b['Data']
                                sinal = "+" if valor_faltante > 0 else ""
                                nota['Motivo Divergência'] = f'Desconto/Acréscimo no PIX ({sinal}R$ {valor_faltante})'
                                resultados['2_Conciliado_Via_Historico'].append(nota)
                                m_a.append(row_a['ID_Argos'])
                                m_b.append(row_b['ID_Bank'])
                                break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 1: CONCILIADO PERFEITO (Valores Exatos)
        # ==========================================
        m_a, m_b = [], []
        for i, row_a in argos_pendentes.iterrows():
            candidatos = bank_pendentes[(bank_pendentes['Valor'] == row_a['Valor']) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
            if not candidatos.empty:
                candidatos['diff_dias'] = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                                               pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                candidatos = candidatos.sort_values(by='diff_dias')
                
                for j, row_b in candidatos.iterrows():
                    if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= 3:
                        nota = row_a.to_dict().copy()
                        nota['Baixas'] = row_b['Banco']
                        nota['Data Baixa'] = row_b['Data']
                        resultados['1_Conciliado_Perfeito'].append(nota)
                        m_a.append(row_a['ID_Argos'])
                        m_b.append(row_b['ID_Bank'])
                        break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 3: DESMEMBRADOS (Força Bruta Restante)
        # ==========================================
        m_a, m_b = [], []
        col_cliente = 'CLIENTES' if 'CLIENTES' in argos_pendentes.columns else 'Cliente'
        
        for cliente, grupo in argos_pendentes.groupby(col_cliente):
            # OTIMIZAÇÃO: Ignora grupos massivos para evitar gargalo
            if len(grupo) < 2 or len(grupo) > 25: continue
            fast_grupo = [(row['ID_Argos'], row['Valor'], row.to_dict()) for idx, row in grupo.iterrows()]
            
            for r in range(2, min(4, len(grupo) + 1)):
                for comb in itertools.combinations(fast_grupo, r):
                    soma_argos = sum(item[1] for item in comb)
                    ids_comb = [item[0] for item in comb]
                    if any(id_a in m_a for id_a in ids_comb): continue
                        
                    candidatos = bank_pendentes[(bank_pendentes['Valor'].between(soma_argos - 0.05, soma_argos + 0.05)) & (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
                    if not candidatos.empty:
                        candidatos['diff_dias'] = abs((pd.to_datetime(grupo.iloc[0]['Data'], format='%d/%m/%Y', errors='coerce') - 
                                                       pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                        candidatos = candidatos.sort_values(by='diff_dias')
                        row_b = candidatos.iloc[0]
                        
                        # CORREÇÃO: Cadeado de Data adicionado!
                        if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= 5:
                            for item in comb:
                                nota = item[2].copy()
                                nota['Baixas'] = row_b['Banco']
                                nota['Data Baixa'] = row_b['Data']
                                resultados['3_Conciliado_Desmembrado'].append(nota)
                                m_a.append(nota['ID_Argos'])
                            m_b.append(row_b['ID_Bank'])
                            break
        remover_matched(m_a, m_b)

        # ==========================================
        # REGRA 3.5: CONCILIADO POR APROXIMAÇÃO DE CENTAVOS (Último Recurso)
        # ==========================================
        m_a, m_b = [], []
        if not argos_pendentes.empty and not bank_pendentes.empty:
            for i, row_a in argos_pendentes.iterrows():
                candidatos = bank_pendentes[(abs(bank_pendentes['Valor'] - row_a['Valor']) > 0) & 
                                            (abs(bank_pendentes['Valor'] - row_a['Valor']) <= 1.50) & 
                                            (~bank_pendentes['ID_Bank'].isin(m_b))].copy()
                if not candidatos.empty:
                    candidatos['diff_dias'] = abs((pd.to_datetime(row_a['Data'], format='%d/%m/%Y', errors='coerce') - 
                                                   pd.to_datetime(candidatos['Data'], format='%d/%m/%Y', errors='coerce')).dt.days)
                    candidatos = candidatos.sort_values(by='diff_dias')
                    
                    for j, row_b in candidatos.iterrows():
                        if pd.notna(row_b['diff_dias']) and row_b['diff_dias'] <= 3:
                            nota = row_a.to_dict().copy()
                            nota['Baixas'] = row_b['Banco']
                            nota['Data Baixa'] = row_b['Data']
                            diff_valor = round(row_b['Valor'] - row_a['Valor'], 2)
                            sinal = "+" if diff_valor > 0 else ""
                            nota['Motivo Divergência'] = f'Aproximação de Centavos ({sinal}R$ {diff_valor})'
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
            nota['Motivo Divergência'] = 'Falta no Banco'
            resultados['4_Divergencias_Pendentes'].append(nota)
            
        for i, row_b in bank_pendentes.iterrows():
            col_hist_b = 'Histórico' if 'Histórico' in row_b else 'N/A'
            nota = {
                'Banco': row_b['Banco'],
                col_cliente: row_b.get(col_hist_b, 'N/A'),
                'Valor': row_b['Valor'],
                'Data': row_b['Data'],
                'Motivo Divergência': 'Sobrou no Banco / Faltou no Argos'
            }
            resultados['4_Divergencias_Pendentes'].append(nota)

        # ==========================================
        # FORMATAÇÃO FINAL DAS COLUNAS (Alinhado com o Template da Cliente)
        # Ordem Exata: Banco | Cliente | Valor | Data (Pgto) | Baixas (Banco Baixa) | Data Baixa | Histórico | Motivo
        # ==========================================
        ordem_colunas = ['Banco', 'Cliente', 'Valor', 'Data', 'Baixas', 'Data Baixa', 'Histórico', 'Motivo Divergência']
        
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

        return resultados




class ExcelReporter:
    @staticmethod
    def generate_report(data_sheets: dict, output_path: str):
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
            'Motivo Divergência': 'MOTIVO DIVERGÊNCIA'
        }
        
        ordem_desejada = ['BANCO', 'CLIENTE', 'VALOR DA BAIXA', 'DATA DO PAGAMENTO', 'BANCO DA BAIXA', 'DATA DA BAIXA', 'HISTÓRICO', 'MOTIVO DIVERGÊNCIA']
        
        for sheet_name, df in data_sheets.items():
            novo_nome_aba = sheet_name.replace('_', ' ')
            if not df.empty:
                df = df.rename(columns=mapa_colunas)
                colunas_finais = [c for c in ordem_desejada if c in df.columns]
                df = df[colunas_finais]
            else:
                df = pd.DataFrame(columns=[c for c in ordem_desejada if c != 'MOTIVO DIVERGÊNCIA'])
            
            formatted_sheets[novo_nome_aba] = df

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in formatted_sheets.items():
                if df.empty:
                    pd.DataFrame({'Aviso': ['Nenhum registo encontrado nesta categoria.']}).to_excel(writer, sheet_name=sheet_name, index=False)
                    continue
                
                for col in df.columns:
                    if 'DATA' in col.upper():
                        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').replace('NaT', '')

                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4)
                worksheet = writer.sheets[sheet_name]

                header_fill = PatternFill(start_color='266C40', end_color='266C40', fill_type='solid')
                fill_verde_claro = PatternFill(start_color='719C82', end_color='719C82', fill_type='solid')
                fill_branco = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
                
                font_branca_bold = Font(color='FFFFFF', bold=True)
                font_preta_bold = Font(color='000000', bold=True)
                
                borda_fina = Border(
                    left=Side(border_style='thin', color='356A1C'),
                    right=Side(border_style='thin', color='356A1C'),
                    top=Side(border_style='thin', color='356A1C'),
                    bottom=Side(border_style='thin', color='356A1C')
                )

                col_indices = {str(worksheet.cell(row=5, column=i).value).upper(): i for i in range(1, worksheet.max_column + 1)}
                max_col = worksheet.max_column

                for r in range(1, 5):
                    for c in range(1, max_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill_verde_claro
                
                worksheet.row_dimensions[1].height = 12
                worksheet.row_dimensions[2].height = 55
                worksheet.row_dimensions[3].height = 12
                worksheet.row_dimensions[4].height = 12

                try:
                    from openpyxl.drawing.image import Image as ExcelImage
                    from PIL import Image as PILImage
                    import glob
                    imagens = glob.glob("img/*.*")
                    if imagens:
                        pil_img = PILImage.open(imagens[0])
                        largura_original, altura_original = pil_img.size
                        
                        nova_altura = 65
                        nova_largura = int((nova_altura / altura_original) * largura_original)
                        
                        img = ExcelImage(imagens[0])
                        img.width = nova_largura
                        img.height = nova_altura
                        
                        if max_col <= 7:
                            coluna_alvo = 4 
                        else:
                            coluna_alvo = 5 
                            
                        col_letra_alvo = get_column_letter(max(1, coluna_alvo))
                        worksheet.add_image(img, f'{col_letra_alvo}2')
                except Exception as e:
                    pass

                for cell in worksheet[5]:
                    cell.fill = header_fill
                    cell.font = font_branca_bold
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                    cell.border = borda_fina

                for row in range(6, worksheet.max_row + 1):
                    is_zebra = (row % 2 == 0)

                    for col_name, col_idx in col_indices.items():
                        cell = worksheet.cell(row=row, column=col_idx)
                        
                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                        cell.font = font_preta_bold
                        cell.border = borda_fina

                        if col_name in ['BANCO', 'BANCO DA BAIXA']:
                            cell.fill = fill_verde_claro
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                        else:
                            if is_zebra:
                                cell.fill = fill_verde_claro
                            else:
                                cell.fill = fill_branco

                        if col_name == 'VALOR DA BAIXA':
                            cell.alignment = Alignment(horizontal='right', vertical='center', wrap_text=False)
                            if cell.value is not None:
                                try:
                                    cell.value = float(cell.value)
                                    cell.number_format = 'R$ #,##0.00'
                                except:
                                    pass

                for col_idx in range(1, max_col + 1):
                    max_length = 0
                    col_letter = get_column_letter(col_idx)
                    col_name_val = worksheet.cell(row=5, column=col_idx).value
                    if not col_name_val: continue
                    col_name = str(col_name_val).upper()
                    
                    max_length = len(col_name) + 3 
                    
                    for row_idx in range(5, worksheet.max_row + 1):
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
                    elif col_name == 'VALOR DA BAIXA':
                        worksheet.column_dimensions[col_letter].width = max(max_length + 6, 20)
                    elif 'DATA' in col_name:
                        worksheet.column_dimensions[col_letter].width = max(max_length + 2, 23)
                    else:
                        worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)


if __name__ == "__main__":
    pass