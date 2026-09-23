# -*- coding: utf-8 -*-
"""
test_conciliacao.py - Suite de Testes Unitários para o Sistema de Conciliação Bancária
==========================================================================================

Suite de testes atualizada para a API real e vigente do conciliacao.py.
Cobre:
    - parse_currency: função global de conversão e saneamento de valores monetários.
    - DataCleaner: limpeza e mapeamento de planilhas do Argos e Extratos Bancários.
    - ReconciliationEngine:
        * Inicialização e segregação de estornos/saídas
        * Estrutura de saída do execute_pipeline (5 abas padronizadas)
        * Regra 1 / 1.1 / 0.5 / 3.5: Matches Perfeitos, Nome e Aproximação de Centavos
        * Regra 2 / 2.5: Conciliação via Histórico (Regex) e Desmembramento Guiado
        * Regra 3: Conciliação Desmembrada (Subset Sum multi-notas)
        * Regra 4: Divergências Pendentes (Falta no Banco / Sobrou no Banco)
        * Integridade e ausência de duplicações de registros
    - ExcelReporter: geração física da planilha final .xlsx multi-abas formatada.

Execução:
    python -m unittest test_conciliacao.py -v
"""

import os
import tempfile
import unittest
import pandas as pd
import openpyxl

from conciliacao import (
    parse_currency,
    safe_float,
    DataCleaner,
    ReconciliationEngine,
    ExcelReporter,
)


class TestParseCurrency(unittest.TestCase):
    """Testa a função utilitária global parse_currency e safe_float."""

    def test_formato_br_com_milhar(self):
        """Converte string no formato brasileiro com separador de milhar."""
        self.assertAlmostEqual(parse_currency("1.234,56"), 1234.56, places=2)

    def test_formato_br_simples(self):
        """Converte string com vírgula decimal simples."""
        self.assertAlmostEqual(parse_currency("200,00"), 200.00, places=2)

    def test_formato_com_simbolo_rs(self):
        """Remove o prefixo R$ e espaços antes da conversão."""
        self.assertAlmostEqual(parse_currency("R$ 1.500,00"), 1500.00, places=2)
        self.assertAlmostEqual(parse_currency("R$512,50"), 512.50, places=2)

    def test_tipo_numerico_direto(self):
        """Float ou int passado diretamente retorna float sem erro."""
        self.assertAlmostEqual(parse_currency(512.50), 512.50, places=2)
        self.assertAlmostEqual(parse_currency(800), 800.00, places=2)

    def test_valores_nulos_e_invalidos(self):
        """Entradas nulas, NaN ou strings não numéricas retornam None com segurança."""
        self.assertIsNone(parse_currency(None))
        self.assertIsNone(parse_currency(float("nan")))
        self.assertIsNone(parse_currency(pd.NA))
        self.assertIsNone(parse_currency("abc"))
        self.assertIsNone(parse_currency(""))
        self.assertIsNone(parse_currency("   "))

    def test_parametro_default_configuravel(self):
        """Valida se o parâmetro default é retornado em vez de None."""
        self.assertEqual(parse_currency(None, default=0.0), 0.0)
        self.assertEqual(parse_currency("invalido", default=0.0), 0.0)
        self.assertEqual(parse_currency("", default=-1.0), -1.0)
        self.assertAlmostEqual(parse_currency("150,00", default=0.0), 150.0, places=2)

    def test_safe_float_utilitario(self):
        """Valida que safe_float converte valores válidos e faz fallback para 0.0 em inválidos."""
        self.assertAlmostEqual(safe_float("1.250,50"), 1250.50, places=2)
        self.assertAlmostEqual(safe_float("R$ 300,00"), 300.00, places=2)
        self.assertEqual(safe_float(None), 0.0)
        self.assertEqual(safe_float(float("nan")), 0.0)
        self.assertEqual(safe_float("texto aleatorio"), 0.0)
        self.assertEqual(safe_float(""), 0.0)


class TestDataCleaner(unittest.TestCase):
    """Testa a limpeza de dados e detecção de arquivos do Argos e Bancos."""

    def setUp(self):
        self.temp_files = []

    def tearDown(self):
        for path in self.temp_files:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except OSError:
                    pass

    def _criar_excel_temporario(self, df: pd.DataFrame, prefixo: str) -> str:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", prefix=prefixo, delete=False) as tf:
            caminho = tf.name
        df.to_excel(caminho, index=False, engine="openpyxl")
        self.temp_files.append(caminho)
        return caminho

    def test_clean_argos_mapeamento_e_formatacao(self):
        """Verifica se clean_argos normaliza colunas, datas e valores corretamente."""
        df_raw = pd.DataFrame({
            "Parceiro Descricao": ["JUNIO SILVA", "JOAO LIMA"],
            "Data": ["16/06/2026", "22/06/2026"],
            "Valor": ["R$ 512,50", "199,17"],
            "Evento Descricao": ["Vendas Atacado", "Vendas Varejo"],
            "Historico": ["pix no valor de 512,50", "pix no valor de 200,00"],
            "Banco": ["CAIXA", "CAIXA"],
        })
        caminho = self._criar_excel_temporario(df_raw, "argos_teste_caixa_")
        df_limpo = DataCleaner.clean_argos(caminho)

        self.assertFalse(df_limpo.empty)
        for col_esperada in ["Banco", "Cliente", "Valor", "Data", "Histórico", "Tipo Evento"]:
            self.assertIn(col_esperada, df_limpo.columns)

        self.assertAlmostEqual(df_limpo["Valor"].iloc[0], 512.50, places=2)
        self.assertAlmostEqual(df_limpo["Valor"].iloc[1], 199.17, places=2)
        self.assertEqual(df_limpo["Cliente"].iloc[0], "JUNIO SILVA")
        self.assertEqual(df_limpo["Data"].iloc[0], "16/06/2026")

    def test_clean_argos_data_sem_ano_usa_ano_atual(self):
        """Verifica se datas no formato DD/MM (sem ano) recebem o ano corrente dinamicamente."""
        from datetime import datetime
        ano_atual = datetime.now().year
        
        df_raw = pd.DataFrame({
            "Parceiro Descricao": ["TESTE ANO"],
            "Data": ["16/06"], # sem ano
            "Valor": ["R$ 100,00"],
            "Evento Descricao": ["Venda"],
            "Historico": [""],
            "Banco": ["CAIXA"],
        })
        caminho = self._criar_excel_temporario(df_raw, "argos_teste_ano_")
        df_limpo = DataCleaner.clean_argos(caminho)

        self.assertFalse(df_limpo.empty)
        # Esperado que a data seja 16/06/ANO_ATUAL
        data_esperada = f"16/06/{ano_atual}"
        self.assertEqual(df_limpo["Data"].iloc[0], data_esperada)

    def test_clean_argos_sem_coluna_valor_retorna_vazio(self):
        """Arquivo do Argos sem nenhuma coluna identificável como Valor retorna DataFrame vazio."""
        df_sem_valor = pd.DataFrame({
            "Cliente": ["Teste"],
            "Data": ["16/06/2026"],
            "Observacoes": ["Sem valor"],
        })
        caminho = self._criar_excel_temporario(df_sem_valor, "argos_invalido_")
        df_limpo = DataCleaner.clean_argos(caminho)
        self.assertTrue(df_limpo.empty)

    def test_clean_argos_deteccao_banco_pelo_conteudo_ou_arquivo(self):
        """Detecta o nome do banco analisando o conteúdo (cabeçalhos) com fallback para o nome."""
        df_raw_caixa = pd.DataFrame({
            "Cliente": ["Cliente A"],
            "Valor": [100.0],
            "Data": ["16/06/2026"],
            "Histórico": ["Extrato da Caixa Economica Federal"],
        })
        caminho_caixa = self._criar_excel_temporario(df_raw_caixa, "extrato_generico_1_")
        df_caixa = DataCleaner.clean_argos(caminho_caixa)
        self.assertEqual(df_caixa["Banco"].iloc[0], "CAIXA ECONOMICA")

        df_raw_banese = pd.DataFrame({
            "Cliente": ["Cliente B"],
            "Valor": [200.0],
            "Data": ["16/06/2026"],
            "Histórico": ["Extrato do BANESE S.A."],
        })
        caminho_banese = self._criar_excel_temporario(df_raw_banese, "extrato_generico_2_")
        df_banese = DataCleaner.clean_argos(caminho_banese)
        self.assertEqual(df_banese["Banco"].iloc[0], "BANESE")

    def test_clean_bank_filtra_debitos_e_normaliza(self):
        """clean_bank descarta valores menores ou iguais a zero (débitos) e padroniza colunas."""
        df_raw = pd.DataFrame({
            "Data": ["16/06/2026", "17/06/2026", "18/06/2026"],
            "Valor": ["512,50", "-50,00", "0,00"],
            "Histórico": ["CREDITO PIX", "DEBITO TARIFA", "SALDO"],
            "Tipo": ["C", "D", "C"],
        })
        caminho = self._criar_excel_temporario(df_raw, "banco_caixa_")
        df_limpo = DataCleaner.clean_bank(caminho)

        self.assertEqual(len(df_limpo), 1)
        self.assertAlmostEqual(df_limpo["Valor"].iloc[0], 512.50, places=2)
        self.assertEqual(df_limpo["Banco"].iloc[0], "CAIXA ECONOMICA")


class TestReconciliationEngineSetup(unittest.TestCase):
    """Testa a inicialização do motor e a segregação de estornos e saídas."""

    def test_segrega_estornos_do_argos(self):
        """Registros com valor negativo ou contendo 'estorno' no Tipo Evento são segregados."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA"],
            "Cliente": ["Cliente Normal", "Cliente Estorno"],
            "Valor": [500.0, -150.0],
            "Data": ["16/06/2026", "16/06/2026"],
            "Histórico": ["Venda", "Devolução"],
            "Tipo Evento": ["Venda", "Estorno de Cupom"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO"],
            "Valor": [500.0],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)

        # O registro de estorno deve sair da base ativa do Argos
        self.assertEqual(len(engine.df_argos), 1)
        self.assertEqual(engine.df_argos["Cliente"].iloc[0], "Cliente Normal")

        # E deve estar registrado em df_saidas_estornos
        self.assertEqual(len(engine.df_saidas_estornos), 1)
        self.assertEqual(engine.df_saidas_estornos["Motivo Divergência"].iloc[0], "Estorno (Argos)")
        self.assertAlmostEqual(engine.df_saidas_estornos["Valor"].iloc[0], 150.0, places=2)

    def test_segrega_saidas_do_banco(self):
        """Registros com valor negativo ou Tipo 'D' no extrato bancário vão para Saídas/Estornos."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["Cliente Normal"],
            "Valor": [500.0],
            "Data": ["16/06/2026"],
            "Histórico": ["Venda"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA"],
            "Histórico": ["CREDITO PIX", "TARIFA BANCARIA"],
            "Valor": [500.0, -25.0],
            "Data": ["16/06/2026", "16/06/2026"],
            "Tipo": ["C", "D"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)

        # A saída deve sair da base ativa do Banco
        self.assertEqual(len(engine.df_bank), 1)
        self.assertEqual(engine.df_bank["Histórico"].iloc[0], "CREDITO PIX")

        # E deve estar em df_saidas_estornos
        self.assertEqual(len(engine.df_saidas_estornos), 1)
        self.assertEqual(engine.df_saidas_estornos["Motivo Divergência"].iloc[0], "Saída (Banco)")


class TestReconciliationEnginePipeline(unittest.TestCase):
    """Testa o contrato do pipeline completo de conciliação."""

    def setUp(self):
        self.colunas_obrigatorias = [
            "Banco", "Cliente", "Valor", "Data",
            "Baixas", "Data Baixa", "Histórico", "Motivo Divergência"
        ]
        self.chaves_obrigatorias = {
            "1_Conciliado_Perfeito",
            "2_Conciliado_Via_Historico",
            "3_Conciliado_Desmembrado",
            "4_Saidas_Estornos",
            "5_Divergencias_Pendentes",
            "6_Resumo_Integridade",
        }

    def test_pipeline_retorna_todas_as_cinco_abas_padronizadas(self):
        """execute_pipeline deve retornar exatamente as 5 abas padronizadas com colunas corretas."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["CLIENTE A"],
            "Valor": [100.0],
            "Data": ["16/06/2026"],
            "Histórico": [""],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO"],
            "Valor": [100.0],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        resultado = engine.execute_pipeline()

        self.assertEqual(set(resultado.keys()), self.chaves_obrigatorias)
        for chave, df in resultado.items():
            self.assertIsInstance(df, pd.DataFrame, f"A chave '{chave}' deve conter um DataFrame.")
            if chave == "6_Resumo_Integridade":
                self.assertEqual(list(df.columns), ["Métrica", "Valor"])
            else:
                self.assertEqual(list(df.columns), self.colunas_obrigatorias)

    def test_pipeline_com_entradas_vazias_nao_lanca_excecao(self):
        """Pipeline executado com DataFrames vazios deve retornar as 5 abas vazias sem exceção."""
        engine = ReconciliationEngine(pd.DataFrame(), pd.DataFrame())
        resultado = engine.execute_pipeline()

        self.assertEqual(set(resultado.keys()), self.chaves_obrigatorias)
        for chave, df in resultado.items():
            if chave == "6_Resumo_Integridade":
                self.assertFalse(df.empty)
                self.assertEqual(list(df.columns), ["Métrica", "Valor"])
            else:
                self.assertTrue(df.empty)
                self.assertEqual(set(df.columns), set(self.colunas_obrigatorias))


class TestRegrasConciliacao(unittest.TestCase):
    """Testa individualmente as regras de negócio integradas no pipeline de conciliação."""

    def test_regra_1_match_perfeito_unico(self):
        """Mesmo valor e data no Argos e Banco conciliam em 1_Conciliado_Perfeito."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["JUNIO SILVA"],
            "Valor": [512.50],
            "Data": ["16/06/2026"],
            "Histórico": ["Venda normal atacado"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO PIX JUNIO"],
            "Valor": [512.50],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_perfeito = res["1_Conciliado_Perfeito"]
        self.assertEqual(len(df_perfeito), 1)
        self.assertAlmostEqual(df_perfeito["Valor"].iloc[0], 512.50, places=2)
        self.assertEqual(df_perfeito["Baixas"].iloc[0], "CAIXA")
        self.assertEqual(df_perfeito["Data Baixa"].iloc[0], "16/06/2026")
        self.assertTrue(res["5_Divergencias_Pendentes"].empty)

    def test_regra_0_5_match_por_nome(self):
        """Quando a descrição do extrato bancário contém o nome do cliente, concilia na Regra 0.5."""
        # Colocamos um segundo registro com mesmo valor para haver concorrência que exigiria o nome
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["MARIA FERREIRA SOUZA"],
            "Valor": [350.00],
            "Data": ["10/06/2026"],
            "Histórico": ["Venda balcão"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["TRANSFERENCIA PIX MARIA FERREIRA"],
            "Valor": [350.00],
            "Data": ["11/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_perfeito = res["1_Conciliado_Perfeito"]
        self.assertEqual(len(df_perfeito), 1)
        self.assertEqual(df_perfeito["Cliente"].iloc[0], "MARIA FERREIRA SOUZA")
        self.assertEqual(df_perfeito["Baixas"].iloc[0], "CAIXA")
        self.assertEqual(df_perfeito["Data Baixa"].iloc[0], "11/06/2026")

    def test_regra_3_5_aproximacao_de_centavos(self):
        """Diferenças de até R$ 1.50 com datas próximas (até 3 dias) são conciliadas com observação."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["POSTO CENTRAL"],
            "Valor": [500.00],
            "Data": ["15/06/2026"],
            "Histórico": [""],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO PIX"],
            "Valor": [500.80],  # Diferença de 80 centavos
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_perfeito = res["1_Conciliado_Perfeito"]
        self.assertEqual(len(df_perfeito), 1)
        self.assertIn("Aproximação de Centavos", str(df_perfeito["Motivo Divergência"].iloc[0]))

    def test_regra_2_conciliado_via_historico_regex(self):
        """Valor nominal no Argos de 199.17 com histórico citando 200,00 concilia com crédito bancário de 200.00."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["JOAO DE OLIVEIRA LIMA JUNIOR"],
            "Valor": [199.17],
            "Data": ["22/06/2026"],
            "Histórico": ["pix no valor de 200,00 dia 22/06"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO PIX"],
            "Valor": [200.00],
            "Data": ["22/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_hist = res["2_Conciliado_Via_Historico"]
        self.assertEqual(len(df_hist), 1)
        self.assertAlmostEqual(df_hist["Valor"].iloc[0], 199.17, places=2)
        self.assertEqual(df_hist["Data Baixa"].iloc[0], "22/06/2026")
        self.assertEqual(df_hist["Baixas"].iloc[0], "CAIXA")
        self.assertTrue(res["5_Divergencias_Pendentes"].empty)

    def test_regra_3_conciliado_desmembrado_subset_sum(self):
        """Múltiplas baixas do Argos do mesmo cliente (445.96 + 354.04) somando 800.00 conciliam com crédito de 800.00."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA"],
            "Cliente": ["LATICINIO SERTANEJO", "LATICINIO SERTANEJO"],
            "Valor": [445.96, 354.04],
            "Data": ["16/06/2026", "16/06/2026"],
            "Histórico": ["comprovante de 800 reais", "comprovante de 800 reais"],
            "Tipo Evento": ["Venda", "Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO PIX"],
            "Valor": [800.00],
            "Data": ["17/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_desm = res["3_Conciliado_Desmembrado"]
        self.assertEqual(len(df_desm), 2)
        soma_partes = round(df_desm["Valor"].sum(), 2)
        self.assertAlmostEqual(soma_partes, 800.00, places=2)
        self.assertTrue(res["5_Divergencias_Pendentes"].empty)

    def test_regra_4_divergencias_pendentes_identificadas(self):
        """Registros sem par são direcionados para 5_Divergencias_Pendentes com os motivos respectivos."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["CLIENTE SEM BANCO"],
            "Valor": [999.00],
            "Data": ["01/06/2026"],
            "Histórico": [""],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO NAO RECONHECIDO"],
            "Valor": [777.00],
            "Data": ["01/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_div = res["5_Divergencias_Pendentes"]
        self.assertEqual(len(df_div), 2)

        motivos = df_div["Motivo Divergência"].tolist()
        self.assertIn("Falta no Banco", motivos)
        self.assertTrue(any("Sobrou no Banco" in m for m in motivos))


class TestIntegridadeEAntiDuplicidade(unittest.TestCase):
    """Testa integridade financeira e garantia de que nenhum registro é duplicado ou perdido."""

    def test_nenhum_registro_argos_duplicado(self):
        """A soma dos registros conciliados e divergentes deve ser exatamente igual ao total de registros válidos do Argos."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA", "CAIXA"],
            "Cliente": ["Cliente 1", "Cliente 2", "Cliente 3"],
            "Valor": [100.00, 200.00, 300.00],
            "Data": ["16/06/2026", "16/06/2026", "16/06/2026"],
            "Histórico": ["", "pix no valor de 210,00", ""],
            "Tipo Evento": ["Venda", "Venda", "Venda"],
        })
        # Banco tem apenas o par do Cliente 1
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["PIX CLIENTE 1"],
            "Valor": [100.00],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        # Cliente 1 deve estar no Perfeito
        self.assertEqual(len(res["1_Conciliado_Perfeito"]), 1)
        # Clientes 2 e 3 devem sobrar no Argos (Falta no Banco)
        div_argos = res["5_Divergencias_Pendentes"][res["5_Divergencias_Pendentes"]["Motivo Divergência"] == "Falta no Banco"]
        self.assertEqual(len(div_argos), 2)

        # Total de registros originados do Argos no resultado
        total_argos_no_resultado = (
            len(res["1_Conciliado_Perfeito"]) +
            len(res["2_Conciliado_Via_Historico"]) +
            len(res["3_Conciliado_Desmembrado"]) +
            len(div_argos)
        )
        self.assertEqual(total_argos_no_resultado, len(df_argos))

    def test_credito_bancario_nao_e_reutilizado(self):
        """Um crédito bancário de valor V concilia com apenas 1 registro do Argos de mesmo valor V."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA"],
            "Cliente": ["Cliente A", "Cliente B"],
            "Valor": [150.00, 150.00],
            "Data": ["16/06/2026", "16/06/2026"],
            "Histórico": ["", ""],
            "Tipo Evento": ["Venda", "Venda"],
        })
        # Apenas 1 crédito de 150.00 no banco
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["PIX"],
            "Valor": [150.00],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        # Exatamente 1 deve ser conciliado
        self.assertEqual(len(res["1_Conciliado_Perfeito"]), 1)
        # O outro deve sobrar como divergência
        div_argos = res["5_Divergencias_Pendentes"][res["5_Divergencias_Pendentes"]["Motivo Divergência"] == "Falta no Banco"]
        self.assertEqual(len(div_argos), 1)


class TestExcelReporter(unittest.TestCase):
    """Testa a geração de relatório Excel multi-abas."""

    def setUp(self):
        with tempfile.NamedTemporaryFile(suffix=".xlsx", prefix="relatorio_teste_", delete=False) as tf:
            self.output_path = tf.name

    def tearDown(self):
        if os.path.exists(self.output_path):
            try:
                os.remove(self.output_path)
            except OSError:
                pass

    def test_generate_report_cria_arquivo_valido_com_todas_as_abas(self):
        """ExcelReporter.generate_report cria arquivo .xlsx que pode ser lido pelo openpyxl."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["CLIENTE A"],
            "Valor": [100.00],
            "Data": ["16/06/2026"],
            "Histórico": ["TESTE"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO"],
            "Valor": [100.00],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        resultados = engine.execute_pipeline()

        ExcelReporter.generate_report(resultados, self.output_path)

        self.assertTrue(os.path.exists(self.output_path))
        self.assertGreater(os.path.getsize(self.output_path), 0)

        # Abre com openpyxl para certificar integridade
        wb = openpyxl.load_workbook(self.output_path)
        nomes_abas = wb.sheetnames

        # Nomes esperados após substituição de underscore por espaço:
        for chave in resultados.keys():
            nome_esperado = chave.replace("_", " ")
            self.assertIn(nome_esperado, nomes_abas)

    def test_generate_report_aplica_alerta_visual_quando_diferenca_maior_que_limite(self):
        """Diferença > 5 dias entre Data do Pagamento e Data da Baixa gera texto e formatação de alerta visual no Excel."""
        resultados = {
            "1_Conciliado_Perfeito": pd.DataFrame([{
                "Banco": "CAIXA",
                "Cliente": "CLIENTE DISTANTE",
                "Valor": 500.00,
                "Data": "01/06/2026",
                "Baixas": "CAIXA",
                "Data Baixa": "20/06/2026", # 19 dias de diferença (> 5)
                "Histórico": "CRED PIX",
                "Motivo Divergência": "",
            }]),
            "2_Conciliado_Via_Historico": pd.DataFrame(),
            "3_Conciliado_Desmembrado": pd.DataFrame(),
            "4_Saidas_Estornos": pd.DataFrame(),
            "5_Divergencias_Pendentes": pd.DataFrame(),
        }

        ExcelReporter.generate_report(resultados, self.output_path, dias_alerta_temporal=5)

        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["1 Conciliado Perfeito"]

        # Identifica colunas pelo cabeçalho
        colunas = {str(ws.cell(row=1, column=c).value).upper(): c for c in range(1, ws.max_column + 1)}
        self.assertIn("OBSERVAÇÃO", colunas)
        self.assertIn("DATA DO PAGAMENTO", colunas)
        self.assertIn("DATA DA BAIXA", colunas)

        col_obs = colunas["OBSERVAÇÃO"]
        col_pgto = colunas["DATA DO PAGAMENTO"]
        col_baixa = colunas["DATA DA BAIXA"]

        cell_obs = ws.cell(row=2, column=col_obs)
        cell_pgto = ws.cell(row=2, column=col_pgto)
        cell_baixa = ws.cell(row=2, column=col_baixa)

        # 1. Verifica texto na coluna OBSERVAÇÃO
        self.assertIn("Alerta Temporal: Diferença de 19 dias", str(cell_obs.value))

        # 2. Verifica cores de preenchimento (FFF2CC)
        cor_obs = str(cell_obs.fill.start_color.rgb).upper()
        cor_pgto = str(cell_pgto.fill.start_color.rgb).upper()
        cor_baixa = str(cell_baixa.fill.start_color.rgb).upper()

        self.assertTrue(cor_obs.endswith("FFF2CC"), f"Cor inesperada para OBSERVAÇÃO: {cor_obs}")
        self.assertTrue(cor_pgto.endswith("FFF2CC"), f"Cor inesperada para DATA DO PAGAMENTO: {cor_pgto}")
        self.assertTrue(cor_baixa.endswith("FFF2CC"), f"Cor inesperada para DATA DA BAIXA: {cor_baixa}")

        # 3. Verifica fonte âmbar em negrito (8A5300)
        self.assertTrue(cell_obs.font.bold)
        self.assertTrue(str(cell_obs.font.color.rgb).upper().endswith("8A5300"))

    def test_generate_report_sem_alerta_visual_quando_dentro_do_limite(self):
        """Diferença <= 5 dias não dispara texto nem formatação de alerta visual."""
        resultados = {
            "1_Conciliado_Perfeito": pd.DataFrame([{
                "Banco": "CAIXA",
                "Cliente": "CLIENTE PROXIMO",
                "Valor": 300.00,
                "Data": "16/06/2026",
                "Baixas": "CAIXA",
                "Data Baixa": "18/06/2026", # 2 dias de diferença (<= 5)
                "Histórico": "CRED PIX",
                "Motivo Divergência": "",
            }]),
            "2_Conciliado_Via_Historico": pd.DataFrame(),
            "3_Conciliado_Desmembrado": pd.DataFrame(),
            "4_Saidas_Estornos": pd.DataFrame(),
            "5_Divergencias_Pendentes": pd.DataFrame(),
        }

        ExcelReporter.generate_report(resultados, self.output_path, dias_alerta_temporal=5)

        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["1 Conciliado Perfeito"]

        colunas = {str(ws.cell(row=1, column=c).value).upper(): c for c in range(1, ws.max_column + 1)}
        cell_obs = ws.cell(row=2, column=colunas["OBSERVAÇÃO"])

        # OBSERVAÇÃO deve estar vazia
        self.assertTrue(cell_obs.value is None or str(cell_obs.value).strip() == "")
        # Cor NÃO deve ser o alerta FFF2CC
        cor_obs = str(cell_obs.fill.start_color.rgb or "").upper()
        self.assertFalse(cor_obs.endswith("FFF2CC"))

    def test_generate_report_override_configuravel_limite_dias(self):
        """Valida que o parâmetro dias_alerta_temporal é configurável."""
        resultados = {
            "1_Conciliado_Perfeito": pd.DataFrame([{
                "Banco": "BANESE",
                "Cliente": "CLIENTE TESTE",
                "Valor": 250.00,
                "Data": "10/06/2026",
                "Baixas": "BANESE",
                "Data Baixa": "18/06/2026", # 8 dias de diferença
                "Histórico": "PIX",
                "Motivo Divergência": "",
            }]),
            "2_Conciliado_Via_Historico": pd.DataFrame(),
            "3_Conciliado_Desmembrado": pd.DataFrame(),
            "4_Saidas_Estornos": pd.DataFrame(),
            "5_Divergencias_Pendentes": pd.DataFrame(),
        }

        # Com limite = 10 dias, 8 dias NÃO gera alerta
        ExcelReporter.generate_report(resultados, self.output_path, dias_alerta_temporal=10)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["1 Conciliado Perfeito"]
        colunas = {str(ws.cell(row=1, column=c).value).upper(): c for c in range(1, ws.max_column + 1)}
        val_obs = ws.cell(row=2, column=colunas["OBSERVAÇÃO"]).value
        self.assertTrue(val_obs is None or str(val_obs).strip() == "")

        # Com limite = 5 dias, 8 dias GERA alerta
        ExcelReporter.generate_report(resultados, self.output_path, dias_alerta_temporal=5)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["1 Conciliado Perfeito"]
        val_obs = ws.cell(row=2, column=colunas["OBSERVAÇÃO"]).value
        self.assertIn("Alerta Temporal: Diferença de 8 dias", str(val_obs))

    def test_execute_pipeline_validacao_soma(self):
        """Valida se a integridade de soma do pipeline é mantida."""
        import warnings
        
        df_argos = pd.DataFrame([
            {"Data": "10/06/2026", "Cliente": "CLI A", "Valor": 150.00, "Tipo Evento": "Pix"},
            {"Data": "11/06/2026", "Cliente": "CLI B", "Valor": 250.00, "Tipo Evento": "Pix"},
            {"Data": "12/06/2026", "Cliente": "CLI C", "Valor": 50.00, "Tipo Evento": "Estorno"}, # Vai p/ saidas
        ])
        df_bank = pd.DataFrame([
            {"Data": "10/06/2026", "Histórico": "PIX CLI A", "Valor": 150.00, "Banco": "BANESE"}, # Perfeito
            {"Data": "15/06/2026", "Histórico": "DOC QUALQUER", "Valor": 80.00, "Banco": "BANESE"}, # Sobra banco
        ])
        
        conciliador = ReconciliationEngine(df_argos, df_bank)
        
        with warnings.catch_warnings(record=True) as w:
            warnings.simplefilter("always")
            resultados = conciliador.execute_pipeline()
            
            # Verifica se NÃO houve aviso CRÍTICO de integridade financeira
            avisos_criticos = [str(av.message) for av in w if "CRÍTICO: Perda de integridade" in str(av.message)]
            self.assertEqual(len(avisos_criticos), 0, "Ocorreu uma perda de integridade inesperada!")

        soma_argos_inicial = 150.0 + 250.0 # O estorno de 50.0 é retirado do argos pendente e vai pra saidas_estornos
        soma_resultados = 0.0
        for k in ['1_Conciliado_Perfeito', '2_Conciliado_Via_Historico', '3_Conciliado_Desmembrado']:
            if not resultados[k].empty:
                soma_resultados += resultados[k]['Valor'].sum()
                
        if not resultados['5_Divergencias_Pendentes'].empty:
            df_div = resultados['5_Divergencias_Pendentes']
            soma_resultados += df_div[df_div['Motivo Divergência'] == 'Falta no Banco']['Valor'].sum()
            
        self.assertAlmostEqual(soma_argos_inicial, soma_resultados, places=2)


class TestConfig(unittest.TestCase):
    """Testa o módulo de configuração central config.py e suas constantes."""

    def test_valores_padrao_constantes(self):
        """Verifica se os valores padrão das constantes de negócio estão corretos."""
        import config
        self.assertEqual(config.JANELA_DIAS, 31)
        self.assertEqual(config.JANELA_DIAS_CENTAVOS, 3)
        self.assertEqual(config.DIAS_ALERTA_TEMPORAL, 5)
        self.assertAlmostEqual(config.TOLERANCIA_CENTAVOS, 1.50, places=2)
        self.assertAlmostEqual(config.TOLERANCIA_DESCONTO_PIX, 15.00, places=2)
        self.assertAlmostEqual(config.LIMIAR_VALOR_FALTANTE_DESMEMBRAR, 0.05, places=2)
        self.assertAlmostEqual(config.TOLERANCIA_INTEGRIDADE, 0.01, places=2)
        self.assertEqual(config.MAX_COMBINACOES, 4)

    def test_helpers_variaveis_ambiente(self):
        """Valida que os helpers _get_int_env e _get_float_env respeitam fallbacks e convertem valores."""
        from config import _get_int_env, _get_float_env

        os.environ["__TEST_INT_VAR"] = "42"
        os.environ["__TEST_FLOAT_VAR"] = "99.9"
        try:
            self.assertEqual(_get_int_env("__TEST_INT_VAR", 10), 42)
            self.assertAlmostEqual(_get_float_env("__TEST_FLOAT_VAR", 1.0), 99.9, places=2)
            self.assertEqual(_get_int_env("__TEST_VAR_INEXISTENTE", 10), 10)
            self.assertAlmostEqual(_get_float_env("__TEST_VAR_INEXISTENTE", 1.5), 1.5, places=2)
        finally:
            os.environ.pop("__TEST_INT_VAR", None)
            os.environ.pop("__TEST_FLOAT_VAR", None)


class TestEngineConfigurability(unittest.TestCase):
    """Testa a possibilidade de customização e override das regras no ReconciliationEngine."""

    def test_override_janela_dias_bloqueia_match_fora_da_janela(self):
        """Se a janela for configurada para 3 dias, transações com 10 dias de diferença na Regra 1 não conciliam."""
        # Dois registros com o mesmo valor para concorrer e acionar a validação de data da Regra 1
        df_argos = pd.DataFrame([
            {"Data": "01/06/2026", "Cliente": "CLI A", "Valor": 100.00, "Tipo Evento": "Pix"},
            {"Data": "01/06/2026", "Cliente": "CLI B", "Valor": 100.00, "Tipo Evento": "Pix"},
        ])
        df_bank = pd.DataFrame([
            {"Data": "11/06/2026", "Histórico": "PIX RECEBIDO", "Valor": 100.00, "Banco": "CAIXA"},
            {"Data": "11/06/2026", "Histórico": "PIX RECEBIDO", "Valor": 100.00, "Banco": "CAIXA"},
        ])

        # Com a janela padrão de 31 dias, deve conciliar em 1_Conciliado_Perfeito
        engine_padrao = ReconciliationEngine(df_argos, df_bank)
        res_padrao = engine_padrao.execute_pipeline()
        self.assertEqual(len(res_padrao["1_Conciliado_Perfeito"]), 2)

        # Com override da janela_dias=3, não deve conciliar na Regra 1 e deve virar divergência
        engine_restrito = ReconciliationEngine(df_argos, df_bank, janela_dias=3)
        res_restrito = engine_restrito.execute_pipeline()
        self.assertEqual(len(res_restrito["1_Conciliado_Perfeito"]), 0)
        self.assertEqual(len(res_restrito["5_Divergencias_Pendentes"]), 4)

    def test_override_tolerancia_centavos(self):
        """Com tolerância de R$ 0.10, uma diferença de R$ 0.50 não concilia por aproximação."""
        df_argos = pd.DataFrame([{
            "Data": "01/06/2026", "Cliente": "CLI X", "Valor": 100.00, "Tipo Evento": "Pix"
        }])
        df_bank = pd.DataFrame([{
            "Data": "02/06/2026", "Histórico": "PIX", "Valor": 100.50, "Banco": "CAIXA"
        }])

        # Tolerância padrão (1.50) concilia na regra 3.5 (aproximação de centavos)
        engine_padrao = ReconciliationEngine(df_argos, df_bank)
        res_padrao = engine_padrao.execute_pipeline()
        self.assertEqual(len(res_padrao["1_Conciliado_Perfeito"]), 1)

        # Tolerância customizada de 0.10 rejeita a aproximação
        engine_restrito = ReconciliationEngine(df_argos, df_bank, tolerancia_centavos=0.10)
        res_restrito = engine_restrito.execute_pipeline()
        self.assertEqual(len(res_restrito["1_Conciliado_Perfeito"]), 0)
        self.assertEqual(len(res_restrito["5_Divergencias_Pendentes"]), 2)


class TestSecurityCORS(unittest.TestCase):
    """Testa a blindagem de segurança de CORS da API FastAPI."""

    def test_cors_config_valores(self):
        """Verifica se as origens padrão e regex do Render estão definidas no config."""
        import config
        self.assertIn("http://localhost:8000", config.DEFAULT_CORS_ORIGINS)
        self.assertIn("http://localhost:3000", config.DEFAULT_CORS_ORIGINS)
        self.assertIn("http://localhost:5173", config.DEFAULT_CORS_ORIGINS)
        self.assertEqual(config.CORS_ALLOW_ORIGIN_REGEX, r"https://.*\.onrender\.com")

    def test_cors_env_parsing(self):
        """Valida parsing de origens via variável de ambiente CORS_ORIGINS."""
        from config import _get_list_env

        # String separada por vírgula
        os.environ["__TEST_CORS"] = "https://app1.com, https://app2.com"
        try:
            origins = _get_list_env("__TEST_CORS", [])
            self.assertEqual(origins, ["https://app1.com", "https://app2.com"])
        finally:
            os.environ.pop("__TEST_CORS", None)

    def test_cors_middleware_permite_origens_autorizadas_e_bloqueia_nao_autorizadas(self):
        """Valida que o middleware CORS autoriza localhost e Render, mas bloqueia origens externas desconhecidas."""
        from fastapi.testclient import TestClient
        from app import app

        client = TestClient(app)

        # 1. Localhost autorizado
        res_local = client.get("/", headers={"Origin": "http://localhost:8000"})
        self.assertEqual(res_local.headers.get("access-control-allow-origin"), "http://localhost:8000")

        # 2. Render autorizado via regex
        res_render = client.get("/", headers={"Origin": "https://carrilho-distribuidora.onrender.com"})
        self.assertEqual(res_render.headers.get("access-control-allow-origin"), "https://carrilho-distribuidora.onrender.com")

        # 3. Origem não autorizada (maliciosa) deve ser BLOQUEADA (sem cabeçalho Access-Control-Allow-Origin)
        res_blocked = client.get("/", headers={"Origin": "https://site-malicioso.com"})
        self.assertIsNone(res_blocked.headers.get("access-control-allow-origin"))


if __name__ == "__main__":
    unittest.main(verbosity=2)


