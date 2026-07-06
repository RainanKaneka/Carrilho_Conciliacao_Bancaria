import pandas as pd
import re
import os
from datetime import timedelta
import itertools
import argparse
import warnings

# Bibliotecas para pintar e formatar o Excel (Bordas adicionadas!)
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings('ignore')

class DataCleaner:
    @staticmethod
    def clean_argos(filepath: str) -> pd.DataFrame:
        df = pd.read_excel(filepath, header=None)
        header_idx = df[df.apply(lambda row: row.astype(str).str.contains('Parceiro Descrição', case=False).any(), axis=1)].index
        if len(header_idx) > 0:
            df.columns = df.iloc[header_idx[0]]
            df = df.iloc[header_idx[0]+1:].reset_index(drop=True)
        else:
            raise ValueError(f"Não foi possível encontrar a tabela principal no ficheiro do Argos: {os.path.basename(filepath)}")

        colunas_manter = ['Parceiro Descrição', 'Data', 'Valor', 'Evento Descrição', 'Histórico']
        colunas_disponiveis = [c for c in colunas_manter if c in df.columns]
        df = df[colunas_disponiveis].copy()

        df['Data'] = pd.to_datetime(df['Data'], errors='coerce', dayfirst=True).dt.normalize()
        df = df.dropna(subset=['Data'])

        def parse_valor(v):
            if pd.isna(v): return 0.0
            if isinstance(v, (int, float)): return float(v)
            v = str(v).replace('R$', '').strip()
            if ',' in v and '.' in v:
                v = v.replace('.', '').replace(',', '.')
            elif ',' in v:
                v = v.replace(',', '.')
            try: return float(v)
            except: return 0.0

        df['Valor'] = df['Valor'].apply(parse_valor)
        df['Histórico'] = df['Histórico'].fillna('').astype(str).str.strip().str.upper()
        df['Origem_Argos'] = os.path.basename(filepath)
        return df

    @staticmethod
    def clean_bank(filepath: str) -> pd.DataFrame:
        df_raw = pd.read_excel(filepath, header=None)
        
        if len(df_raw.columns) <= 3:
            records = []
            for val in df_raw.iloc[:, 0].dropna().astype(str):
                match = re.search(r'^(\d{2}/\d{2}/\d{4})\s+(.*?)\s+(\d{1,3}(?:\.\d{3})*,\d{2})\s*\+', val)
                if match:
                    records.append({
                        'data_banco': pd.to_datetime(match.group(1), format='%d/%m/%Y'),
                        'descricao_banco': match.group(2).strip(),
                        'valor_banco': float(match.group(3).replace('.', '').replace(',', '.'))
                    })
            if records:
                df = pd.DataFrame(records)
                df['Origem_Banco'] = os.path.basename(filepath)
                return df

        header_idx = -1
        for i, row in df_raw.iterrows():
            row_str = " ".join(row.dropna().astype(str).str.lower())
            if ('data' in row_str) and ('valor' in row_str or 'cred' in row_str or 'créd' in row_str or 'lancamento' in row_str or 'lançamento' in row_str):
                header_idx = i; break

        if header_idx != -1:
            df_raw.columns = df_raw.iloc[header_idx]
            df = df_raw.iloc[header_idx+1:].reset_index(drop=True)
        else:
            df = df_raw.copy()

        df = df.loc[:, df.columns.notna()]
        col_map = {}
        for col in df.columns:
            c = str(col).lower().strip()
            if 'data' in c and 'data_banco' not in col_map.values(): 
                col_map[col] = 'data_banco'
            elif ('valor' in c or 'cred' in c or 'créd' in c or 'lançamento' in c or 'lancamento' in c) and 'valor_banco' not in col_map.values(): 
                col_map[col] = 'valor_banco'
            elif ('hist' in c or 'hitórico' in c or 'descrição' in c or 'detalhe' in c) and 'descricao_banco' not in col_map.values(): 
                col_map[col] = 'descricao_banco'

        df = df.rename(columns=col_map)
        
        if 'data_banco' not in df.columns or 'valor_banco' not in df.columns:
            raise ValueError(f"ERRO: Não encontrei as colunas 'Data' e 'Valor' no extrato: {os.path.basename(filepath)}")
        
        colunas_finais = ['data_banco', 'valor_banco']
        if 'descricao_banco' in df.columns: colunas_finais.append('descricao_banco')
        df = df[colunas_finais].copy()
        
        if 'descricao_banco' not in df.columns:
            df['descricao_banco'] = "Sem descrição no arquivo"
        
        df['data_banco'] = pd.to_datetime(df['data_banco'], errors='coerce').dt.normalize()
        df = df.dropna(subset=['data_banco'])
        df['valor_banco'] = df['valor_banco'].astype(str).str.replace('R$', '', regex=False).str.strip()
        df['valor_banco'] = pd.to_numeric(df['valor_banco'].str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce')
        df = df[df['valor_banco'] > 0].copy()
        df['Origem_Banco'] = os.path.basename(filepath)
        return df

class ReconciliationEngine:
    def __init__(self, df_argos: pd.DataFrame, df_bank: pd.DataFrame):
        self.df_argos = df_argos.copy()
        self.df_bank = df_bank.copy()
        self.df_argos['status'] = 'pendente'
        self.df_bank['status'] = 'pendente'
        self.res_perfeito = []
        self.res_historico = []
        self.res_desmembrado = []
        self.palavras_bloqueio = ['RESTANTE', 'OUTRA BAIXA', 'OUTRA NOTA', 'PARTE', 'DESMEMBRADO']

    def match_exact(self):
        for idx_a, row_a in self.df_argos[self.df_argos['status'] == 'pendente'].iterrows():
            if any(palavra in row_a['Histórico'] for palavra in self.palavras_bloqueio):
                continue

            v_argos = row_a['Valor']
            d_argos = row_a['Data']
            candidatos = self.df_bank[
                (self.df_bank['status'] == 'pendente') &
                (self.df_bank['valor_banco'] == v_argos) &
                (self.df_bank['data_banco'] >= d_argos - timedelta(days=3)) &
                (self.df_bank['data_banco'] <= d_argos + timedelta(days=3))
            ]
            if not candidatos.empty:
                idx_b = candidatos.index[0]
                self.df_argos.at[idx_a, 'status'] = 'conciliado'
                self.df_bank.at[idx_b, 'status'] = 'conciliado'
                match_data = row_a.to_dict()
                match_data['data_banco_real'] = candidatos.iloc[0]['data_banco']
                match_data['Texto_Original_Banco'] = candidatos.iloc[0]['descricao_banco']
                match_data['Banco_Real'] = candidatos.iloc[0]['Origem_Banco']
                self.res_perfeito.append(match_data)

    def match_by_history_regex(self):
        regex = re.compile(r'(?:valor de|valor|pix de)?\s*([0-9]{1,3}(?:\.[0-9]{3})*\,[0-9]{2})', re.IGNORECASE)
        for idx_a, row_a in self.df_argos[self.df_argos['status'] == 'pendente'].iterrows():
            if any(palavra in row_a['Histórico'] for palavra in self.palavras_bloqueio):
                continue

            match = regex.search(row_a['Histórico'])
            if match:
                try: v_regex = float(match.group(1).replace('.', '').replace(',', '.'))
                except: continue

                # --- NOVA TRAVA DE SEGURANÇA ---
                # Se a diferença entre a nota e o texto for muito grande (> 15 reais),
                # não é um erro de centavos, é um PIX desmembrado que o vendedor anotou na observação!
                # Ignoramos na Regra 2 e deixamos a Regra 3 somar as notas.
                if abs(row_a['Valor'] - v_regex) > 15.0:
                    continue
                # -------------------------------

                d_argos = row_a['Data']
                candidatos = self.df_bank[
                    (self.df_bank['status'] == 'pendente') &
                    (self.df_bank['valor_banco'] == v_regex) &
                    (self.df_bank['data_banco'] >= d_argos - timedelta(days=3)) &
                    (self.df_bank['data_banco'] <= d_argos + timedelta(days=3))
                ]
                if not candidatos.empty:
                    idx_b = candidatos.index[0]
                    self.df_argos.at[idx_a, 'status'] = 'conciliado'
                    self.df_bank.at[idx_b, 'status'] = 'conciliado'
                    match_data = row_a.to_dict()
                    match_data['valor_real_banco'] = v_regex
                    match_data['data_banco_real'] = candidatos.iloc[0]['data_banco']
                    match_data['Texto_Original_Banco'] = candidatos.iloc[0]['descricao_banco']
                    match_data['Banco_Real'] = candidatos.iloc[0]['Origem_Banco']
                    self.res_historico.append(match_data)

    def match_split_sums(self):
        pendentes_banco = self.df_bank[self.df_bank['status'] == 'pendente']
        for idx_b, row_b in pendentes_banco.iterrows():
            v_banco = row_b['valor_banco']
            d_banco = row_b['data_banco']
            
            pendentes_argos = self.df_argos[self.df_argos['status'] == 'pendente']
            
            candidatos_argos = pendentes_argos[
                (pendentes_argos['Data'] >= d_banco - timedelta(days=3)) &
                (pendentes_argos['Data'] <= d_banco + timedelta(days=3)) &
                (pendentes_argos['Valor'] < v_banco)
            ]
            
            if candidatos_argos.empty: continue

            achou = False
            for parceiro in candidatos_argos['Parceiro Descrição'].unique():
                candidatos_parceiro = candidatos_argos[candidatos_argos['Parceiro Descrição'] == parceiro]
                if len(candidatos_parceiro) >= 2:
                    candidatos_list = list(candidatos_parceiro['Valor'].to_dict().items())
                    for r in range(2, min(5, len(candidatos_list) + 1)):
                        for comb in itertools.combinations(candidatos_list, r):
                            soma = sum(item[1] for item in comb)
                            if round(soma, 2) == round(v_banco, 2):
                                achou = [item[0] for item in comb]
                                break
                        if achou: break
                if achou: break
            
            if not achou:
                candidatos_list = list(candidatos_argos['Valor'].to_dict().items())
                for r in range(2, 4):
                    if r == 3 and len(candidatos_list) > 60:
                        continue
                    for comb in itertools.combinations(candidatos_list, r):
                        soma = sum(item[1] for item in comb)
                        if round(soma, 2) == round(v_banco, 2):
                            achou = [item[0] for item in comb]
                            break
                    if achou: break

            if achou:
                self.df_bank.at[idx_b, 'status'] = 'conciliado'
                for idx_a in achou:
                    self.df_argos.at[idx_a, 'status'] = 'conciliado'
                    match_data = self.df_argos.loc[idx_a].to_dict()
                    match_data['vinculado_a_valor_banco'] = v_banco
                    match_data['data_banco_real'] = row_b['data_banco']
                    match_data['Texto_Original_Banco'] = row_b['descricao_banco']
                    match_data['Banco_Real'] = row_b['Origem_Banco']
                    self.res_desmembrado.append(match_data)

    def formatar_aba_sucesso(self, dados):
        if not dados:
            return pd.DataFrame(columns=['BANCO', 'CLIENTES', 'VALOR', 'DATA PAGAMENTO', 'BAIXAS', 'DATA BAIXA', 'OBS'])
        
        df = pd.DataFrame(dados)
        
        def padronizar_banco(nome):
            if pd.isna(nome) or not nome: return ""
            n = str(nome).lower()
            if 'caixa' in n: return 'CAIXA ECONOMICA'
            if 'banese' in n: return 'BANESE'
            return nome

        df['BANCO'] = df.get('Banco_Real', pd.Series(dtype='str')).apply(padronizar_banco)
        df['CLIENTES'] = df['Parceiro Descrição']
        df['VALOR'] = df['Valor']
        df['DATA PAGAMENTO'] = df['data_banco_real']
        df['BAIXAS'] = df.get('Origem_Argos', pd.Series(dtype='str')).apply(padronizar_banco)
        df['DATA BAIXA'] = df['Data']
        df['OBS'] = df['Histórico']

        return df[['BANCO', 'CLIENTES', 'VALOR', 'DATA PAGAMENTO', 'BAIXAS', 'DATA BAIXA', 'OBS']]

    def execute_pipeline(self):
        self.match_exact()
        self.match_by_history_regex()
        self.match_split_sums()

        df_1 = self.formatar_aba_sucesso(self.res_perfeito)
        df_2 = self.formatar_aba_sucesso(self.res_historico)
        df_3 = self.formatar_aba_sucesso(self.res_desmembrado)

        max_data_banco = self.df_bank['data_banco'].max() if not self.df_bank.empty else pd.Timestamp.max

        div_argos = self.df_argos[(self.df_argos['status'] == 'pendente') & (self.df_argos['Data'] <= max_data_banco)].copy()
        div_argos['Origem_Divergencia'] = 'Falta no Banco'

        div_banco = self.df_bank[self.df_bank['status'] == 'pendente'].copy()
        div_banco_formatado = pd.DataFrame({
            'Data': div_banco['data_banco'],
            'Valor': div_banco['valor_banco'],
            'Evento Descrição': div_banco['descricao_banco'],
            'Origem_Banco': div_banco['Origem_Banco'],
            'Origem_Divergencia': 'Sobrou no Banco / Faltou no Argos'
        })

        df_div_bruto = pd.concat([div_argos, div_banco_formatado], ignore_index=True)
        
        if not df_div_bruto.empty:
            def padronizar_banco(nome):
                if pd.isna(nome) or not nome: return ""
                n = str(nome).lower()
                if 'caixa' in n: return 'CAIXA ECONOMICA'
                if 'banese' in n: return 'BANESE'
                return nome

            df_div_bruto['BANCO'] = df_div_bruto.get('Origem_Banco', pd.Series(dtype='str')).apply(padronizar_banco)
            df_div_bruto['BAIXAS'] = df_div_bruto.get('Origem_Argos', pd.Series(dtype='str')).apply(padronizar_banco)
            df_div_bruto['CLIENTES'] = df_div_bruto.apply(lambda r: r['Parceiro Descrição'] if pd.notna(r.get('Parceiro Descrição')) else r.get('Evento Descrição', ''), axis=1)
            df_div_bruto['VALOR'] = df_div_bruto['Valor']
            
            df_div_bruto['DATA PAGAMENTO'] = df_div_bruto.apply(lambda r: r['Data'] if r['Origem_Divergencia'] == 'Sobrou no Banco / Faltou no Argos' else pd.NaT, axis=1)
            df_div_bruto['DATA BAIXA'] = df_div_bruto.apply(lambda r: r['Data'] if r['Origem_Divergencia'] == 'Falta no Banco' else pd.NaT, axis=1)
            
            df_div_bruto['OBS'] = df_div_bruto.apply(lambda r: r['Histórico'] if pd.notna(r.get('Histórico')) else r.get('Evento Descrição', ''), axis=1)
            df_div_bruto['MOTIVO DIVERGÊNCIA'] = df_div_bruto['Origem_Divergencia']

            df_div_final = df_div_bruto[['BANCO', 'CLIENTES', 'VALOR', 'DATA PAGAMENTO', 'BAIXAS', 'DATA BAIXA', 'OBS', 'MOTIVO DIVERGÊNCIA']]
        else:
            df_div_final = pd.DataFrame(columns=['BANCO', 'CLIENTES', 'VALOR', 'DATA PAGAMENTO', 'BAIXAS', 'DATA BAIXA', 'OBS', 'MOTIVO DIVERGÊNCIA'])

        return {
            '1_Conciliado_Perfeito': df_1,
            '2_Conciliado_Via_Historico': df_2,
            '3_Conciliado_Desmembrado': df_3,
            '4_Divergencias_Pendentes': df_div_final
        }

class ExcelReporter:
    @staticmethod
    def generate_report(data_sheets: dict, output_path: str):
        # --- NOVO: Tratamento de Nomes das Abas e Colunas antes de exportar ---
        formatted_sheets = {}
        
        # Dicionário tradutor (De -> Para)
        mapa_colunas = {
            'BANCO': 'BANCO',
            'CLIENTES': 'CLIENTE',
            'VALOR': 'VALOR DA BAIXA',
            'DATA PAGAMENTO': 'DATA DO PAGAMENTO',
            'BAIXAS': 'BANCO DA BAIXA',
            'DATA BAIXA': 'DATA DA BAIXA',
            'OBS': 'HISTÓRICO'
        }
        
        for sheet_name, df in data_sheets.items():
            # Remove os "_" dos nomes das abas (ex: "1_Conciliado_Perfeito" -> "1 Conciliado Perfeito")
            novo_nome_aba = sheet_name.replace('_', ' ')
            
            if not df.empty:
                # Aplica o dicionário tradutor para renomear as colunas
                df = df.rename(columns=mapa_colunas)
            
            formatted_sheets[novo_nome_aba] = df

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for sheet_name, df in formatted_sheets.items():
                if df.empty:
                    pd.DataFrame({'Aviso': ['Nenhum registo encontrado.']}).to_excel(writer, sheet_name=sheet_name, index=False)
                    continue
                
                for col in df.columns:
                    if 'DATA' in col.upper():
                        df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y').replace('NaT', '')

                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=4)
                worksheet = writer.sheets[sheet_name]

                # --- PALETA DE CORES INSTITUCIONAIS E FONTES ---
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

                # 0. Fundo verde nas 4 primeiras linhas e Altura exata
                for r in range(1, 5):
                    for c in range(1, max_col + 1):
                        worksheet.cell(row=r, column=c).fill = fill_verde_claro
                
                worksheet.row_dimensions[1].height = 12
                worksheet.row_dimensions[2].height = 55
                worksheet.row_dimensions[3].height = 12
                worksheet.row_dimensions[4].height = 12

                # --- INSERIR LOGOMARCA ---
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
                            
                        col_letra_alvo = get_column_letter(coluna_alvo)
                        worksheet.add_image(img, f'{col_letra_alvo}2')
                except Exception as e:
                    print(f"Aviso: Não foi possível inserir a imagem. Erro: {e}")

                # 1. Formatação Visual de Cabeçalhos (Linha 5)
                for cell in worksheet[5]:
                    cell.fill = header_fill
                    cell.font = font_branca_bold
                    # GARANTIA: Sem quebra de linha nos títulos (wrap_text=False)
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
                    cell.border = borda_fina

                # 2. Formatação das Células de Dados (Começam na linha 6)
                for row in range(6, worksheet.max_row + 1):
                    is_zebra = (row % 2 == 0)

                    for col_name, col_idx in col_indices.items():
                        cell = worksheet.cell(row=row, column=col_idx)
                        
                        cell.alignment = Alignment(vertical='center', wrap_text=True)
                        cell.font = font_preta_bold
                        cell.border = borda_fina

                        # Atualizamos as referências para os NOVOS nomes das colunas!
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

                # 3. Auto-Fit: Ajustar a largura das colunas dinamicamente
                for col_idx in range(1, max_col + 1):
                    max_length = 0
                    col_letter = get_column_letter(col_idx)
                    col_name_val = worksheet.cell(row=5, column=col_idx).value
                    if not col_name_val: continue
                    col_name = str(col_name_val).upper()
                    
                    # Garante um respiro largo para os títulos (evitando a quebra de linha)
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
                    
                    # Definimos larguras mínimas baseadas nos novos textos
                    if col_name in ['BANCO', 'BANCO DA BAIXA']:
                        worksheet.column_dimensions[col_letter].width = max(max_length, 20)
                    elif col_name == 'VALOR DA BAIXA':
                        worksheet.column_dimensions[col_letter].width = max(max_length + 6, 20)
                    elif 'DATA' in col_name:
                        worksheet.column_dimensions[col_letter].width = max(max_length + 2, 23)
                    else:
                        worksheet.column_dimensions[col_letter].width = min(max_length + 2, 50)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Conciliacao Bancaria Global")
    parser.add_argument('--argos', nargs='+', required=True, help="Ficheiros do Argos")
    parser.add_argument('--bancos', nargs='+', required=True, help="Extratos dos Bancos")
    args = parser.parse_args()

    print("\n[1/3] Limpando e unificando arquivos do Argos...")
    dfs_argos = [DataCleaner.clean_argos(arq) for arq in args.argos]
    df_argos_full = pd.concat(dfs_argos, ignore_index=True) if dfs_argos else pd.DataFrame()
    print(f"      -> {len(df_argos_full)} registros lidos de {len(args.argos)} arquivo(s).")

    print("[2/3] Inteligência Artificial limpando e unificando extratos bancarios...")
    dfs_bancos = [DataCleaner.clean_bank(arq) for arq in args.bancos]
    df_bancos_full = pd.concat(dfs_bancos, ignore_index=True) if dfs_bancos else pd.DataFrame()
    print(f"      -> {len(df_bancos_full)} depositos validos lidos de {len(args.bancos)} arquivo(s).")

    print("[3/3] Cruzando todas as informacoes (Motor Global)...")
    engine = ReconciliationEngine(df_argos_full, df_bancos_full)
    relatorios = engine.execute_pipeline()

    print(f"      -> Conciliados Perfeitos   : {len(relatorios['1_Conciliado_Perfeito'])}")
    print(f"      -> Lidos via Histórico     : {len(relatorios['2_Conciliado_Via_Historico'])}")
    print(f"      -> Valores Desmembrados    : {len(relatorios['3_Conciliado_Desmembrado'])}")
    print(f"      -> DIVERGENCIAS PENDENTES  : {len(relatorios['4_Divergencias_Pendentes'])}")

    output = "conciliacao_global_final.xlsx"
    ExcelReporter.generate_report(relatorios, output)
    print(f"\n[OK] Concluido! Planilha salva como: {output}\n")