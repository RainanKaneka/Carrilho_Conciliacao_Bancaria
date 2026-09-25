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
import re
import pandas as pd
import openpyxl

from conciliacao import (
    parse_currency,
    safe_float,
    DataCleaner,
    ReconciliationEngine,
    ExcelReporter,
    CorruptedFileError,
    AnalyticsService,
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
            "Baixas", "Data Baixa", "Histórico", "Motivo Divergência", "Regra Aplicada"
        ]
        self.chaves_obrigatorias = {
            "0_Resumo_Executivo",
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
            if chave in ["0_Resumo_Executivo", "6_Resumo_Integridade"]:
                self.assertEqual(list(df.columns), ["Métrica", "Valor"])
            else:
                self.assertEqual(list(df.columns), self.colunas_obrigatorias)

    def test_pipeline_com_entradas_vazias_nao_lanca_excecao(self):
        """Pipeline executado com DataFrames vazios deve retornar as 5 abas vazias sem exceção."""
        engine = ReconciliationEngine(pd.DataFrame(), pd.DataFrame())
        resultado = engine.execute_pipeline()

        self.assertEqual(set(resultado.keys()), self.chaves_obrigatorias)
        for chave, df in resultado.items():
            if chave in ["0_Resumo_Executivo", "6_Resumo_Integridade"]:
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

        # Certifica que a primeira aba é o Resumo Executivo
        self.assertEqual(nomes_abas[0], "0 Resumo Executivo")

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


class TestResumoExecutivo(unittest.TestCase):
    """Testa a geração, formatação, integridade e assinatura digital da aba Resumo Executivo (Fase 2.2)."""

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.output_path = os.path.join(self.temp_dir.name, "teste_executivo.xlsx")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_resumo_executivo_primeira_aba_no_dict_e_no_excel(self):
        """0_Resumo_Executivo deve ser a PRIMEIRA chave retornada e a PRIMEIRA aba criada no Excel."""
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
        res = engine.execute_pipeline()

        # 1. Primeira chave do dicionário
        primeira_chave = list(res.keys())[0]
        self.assertEqual(primeira_chave, "0_Resumo_Executivo")

        # 2. Primeira aba do Excel
        file_hash = ExcelReporter.generate_report(res, self.output_path)
        self.assertTrue(os.path.exists(self.output_path))
        self.assertIsInstance(file_hash, str)
        self.assertEqual(len(file_hash), 64)

        wb = openpyxl.load_workbook(self.output_path)
        self.assertEqual(wb.sheetnames[0], "0 Resumo Executivo")

    def test_resumo_executivo_contem_todos_os_campos_obrigatorios(self):
        """Resumo Executivo deve conter totais, taxa de sucesso, data/hora e assinatura digital."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA", "CAIXA"],
            "Cliente": ["CLIENTE 1", "CLIENTE 2"],
            "Valor": [200.00, 300.00],
            "Data": ["10/06/2026", "11/06/2026"],
            "Histórico": ["Venda 1", "Venda 2"],
            "Tipo Evento": ["Venda", "Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO 1"],
            "Valor": [200.00],
            "Data": ["10/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_exec = res["0_Resumo_Executivo"]
        self.assertEqual(list(df_exec.columns), ["Métrica", "Valor"])

        metricas = dict(zip(df_exec["Métrica"], df_exec["Valor"]))

        # 1. Data e Hora
        self.assertIn("Data/Hora de Processamento", metricas)
        data_hora = str(metricas["Data/Hora de Processamento"])
        self.assertTrue(re.match(r"^\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}$", data_hora))

        # 2. Status da Conciliação
        self.assertIn("Status da Conciliação", metricas)

        # 3. Taxa de Sucesso
        self.assertIn("Taxa de Sucesso (Registros)", metricas)
        self.assertIn("Taxa de Sucesso (Financeira)", metricas)
        self.assertTrue(str(metricas["Taxa de Sucesso (Registros)"]).endswith("%"))
        self.assertTrue(str(metricas["Taxa de Sucesso (Financeira)"]).endswith("%"))

        # 4. Totais
        self.assertIn("Total de Registros Argos", metricas)
        self.assertEqual(metricas["Total de Registros Argos"], 2)
        self.assertIn("Total Volume Argos (R$)", metricas)
        self.assertAlmostEqual(metricas["Total Volume Argos (R$)"], 500.00, places=2)
        self.assertIn("Total de Registros Banco", metricas)
        self.assertEqual(metricas["Total de Registros Banco"], 1)
        self.assertIn("Total Volume Banco (R$)", metricas)
        self.assertAlmostEqual(metricas["Total Volume Banco (R$)"], 200.00, places=2)

        self.assertIn("Total Conciliado (Registros)", metricas)
        self.assertEqual(metricas["Total Conciliado (Registros)"], 1)
        self.assertIn("Total Conciliado (R$)", metricas)
        self.assertAlmostEqual(metricas["Total Conciliado (R$)"], 200.00, places=2)

        self.assertIn("Total Divergências Pendentes (Registros)", metricas)
        self.assertEqual(metricas["Total Divergências Pendentes (Registros)"], 1)
        self.assertIn("Total Divergências Pendentes (R$)", metricas)
        self.assertAlmostEqual(metricas["Total Divergências Pendentes (R$)"], 300.00, places=2)

        # 5. Assinatura Digital (Hash SHA-256)
        self.assertIn("Assinatura Digital (Hash SHA-256)", metricas)
        hash_sha256 = str(metricas["Assinatura Digital (Hash SHA-256)"])
        self.assertEqual(len(hash_sha256), 64)
        self.assertTrue(re.match(r"^[0-9a-f]{64}$", hash_sha256))

    def test_taxa_sucesso_calculo_correto(self):
        """Valida que a taxa de sucesso percentual é calculada com exatidão matemática."""
        # 3 conciliados perfeitos + 1 divergência = 3/4 = 75.00%
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"] * 4,
            "Cliente": ["CLIENTE 1", "CLIENTE 2", "CLIENTE 3", "CLIENTE 4"],
            "Valor": [100.0, 200.0, 300.0, 400.0],
            "Data": ["10/06/2026", "11/06/2026", "12/06/2026", "13/06/2026"],
            "Histórico": ["Venda 1", "Venda 2", "Venda 3", "Venda 4"],
            "Tipo Evento": ["Venda"] * 4,
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"] * 3,
            "Histórico": ["CREDITO 1", "CREDITO 2", "CREDITO 3"],
            "Valor": [100.0, 200.0, 300.0],
            "Data": ["10/06/2026", "11/06/2026", "12/06/2026"],
            "Tipo": ["C"] * 3,
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        df_exec = res["0_Resumo_Executivo"]
        metricas = dict(zip(df_exec["Métrica"], df_exec["Valor"]))

        self.assertEqual(metricas["Taxa de Sucesso (Registros)"], "75.00%")
        self.assertEqual(metricas["Taxa de Sucesso (Financeira)"], "60.00%")

    def test_assinatura_digital_sha256_consistente_e_sensivel_a_alteracoes(self):
        """A assinatura digital é determinística para mesmos dados e sensível (efeito avalanche) para qualquer alteração."""
        df_argos_1 = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["CLIENTE A"],
            "Valor": [100.00],
            "Data": ["16/06/2026"],
            "Histórico": ["TESTE"],
        })
        df_bank_1 = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO"],
            "Valor": [100.00],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        resultados_dummy = {
            "1_Conciliado_Perfeito": df_argos_1.copy(),
            "5_Divergencias_Pendentes": pd.DataFrame(),
        }

        # Mesmo timestamp e dados idênticos geram o mesmo hash
        fixed_ts = "24/09/2026 12:00:00"
        hash_1 = ReconciliationEngine._calcular_assinatura_digital(df_argos_1, df_bank_1, resultados_dummy, fixed_ts)
        hash_2 = ReconciliationEngine._calcular_assinatura_digital(df_argos_1.copy(), df_bank_1.copy(), resultados_dummy, fixed_ts)
        self.assertEqual(hash_1, hash_2)
        self.assertEqual(len(hash_1), 64)

        # Alterando apenas 1 centavo no valor de entrada
        df_argos_modificado = df_argos_1.copy()
        df_argos_modificado.loc[0, "Valor"] = 100.01
        hash_alterado = ReconciliationEngine._calcular_assinatura_digital(df_argos_modificado, df_bank_1, resultados_dummy, fixed_ts)
        self.assertNotEqual(hash_1, hash_alterado)

    def test_resumo_executivo_formatacao_visual_excel(self):
        """Verifica estilos do Excel (cabeçalho verde, colunas formatadas, largura da assinatura digital)."""
        df_argos = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Cliente": ["CLIENTE A"],
            "Valor": [1500.50],
            "Data": ["16/06/2026"],
            "Histórico": ["TESTE"],
            "Tipo Evento": ["Venda"],
        })
        df_bank = pd.DataFrame({
            "Banco": ["CAIXA"],
            "Histórico": ["CREDITO"],
            "Valor": [1500.50],
            "Data": ["16/06/2026"],
            "Tipo": ["C"],
        })

        engine = ReconciliationEngine(df_argos, df_bank)
        res = engine.execute_pipeline()

        ExcelReporter.generate_report(res, self.output_path)

        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["0 Resumo Executivo"]

        # Cabeçalho
        self.assertEqual(ws.cell(row=1, column=1).value, "Métrica")
        self.assertEqual(ws.cell(row=1, column=2).value, "Valor")
        self.assertEqual(ws.cell(row=1, column=1).fill.start_color.rgb, "00266C40")
        self.assertTrue(ws.cell(row=1, column=1).font.bold)

        # Verifica que a coluna VALOR tem largura suficiente para comportar o hash SHA-256 (64 chars)
        largura_col_b = ws.column_dimensions["B"].width
        self.assertGreaterEqual(largura_col_b, 65)

    def test_resumo_executivo_com_entradas_vazias(self):
        """Resumo Executivo deve ser gerado com segurança mesmo com entradas vazias."""
        engine = ReconciliationEngine(pd.DataFrame(), pd.DataFrame())
        res = engine.execute_pipeline()

        self.assertIn("0_Resumo_Executivo", res)
        df_exec = res["0_Resumo_Executivo"]
        self.assertFalse(df_exec.empty)
        metricas = dict(zip(df_exec["Métrica"], df_exec["Valor"]))

        self.assertEqual(metricas["Status da Conciliação"], "Entradas Vazias")
        self.assertEqual(metricas["Taxa de Sucesso (Registros)"], "0.00%")
        self.assertEqual(metricas["Total de Registros Argos"], 0)
        self.assertEqual(len(metricas["Assinatura Digital (Hash SHA-256)"]), 64)


class TestTimeoutCombinacoes(unittest.TestCase):
    def test_timeout_combinacoes_aborta_busca(self):
        """Um timeout de 0.0 segundos deve abortar imediatamente o itertools.combinations e pular a Regra 3."""
        df_argos = pd.DataFrame([
            {"Banco": "BANCO", "Cliente": "CLI_1", "Valor": 10.00, "Data": "01/01/2026"},
            {"Banco": "BANCO", "Cliente": "CLI_1", "Valor": 20.00, "Data": "01/01/2026"},
            {"Banco": "BANCO", "Cliente": "CLI_1", "Valor": 30.00, "Data": "01/01/2026"},
        ])
        df_bank = pd.DataFrame([
            {"Banco": "BANCO", "Data": "01/01/2026", "Valor": 60.00, "Histórico": "DEPOSITO"}
        ])
        
        # Com timeout normal (5.0s), isso iria conciliar perfeitamente na Regra 3
        # max_combinacoes=4 significa que testa 1, 2 e 3 combinações (já que range vai até max_combinacoes)
        engine_normal = ReconciliationEngine(df_argos, df_bank, max_combinacoes=4, timeout_combinacoes=5.0)
        res_normal = engine_normal.execute_pipeline()
        self.assertEqual(len(res_normal["3_Conciliado_Desmembrado"]), 3)

        # Com timeout de -1.0s, o loop aborta garantidamente na primeira iteração (resolvendo resolução de timer)
        engine_timeout = ReconciliationEngine(df_argos, df_bank, max_combinacoes=4, timeout_combinacoes=-1.0)
        res_timeout = engine_timeout.execute_pipeline()
        self.assertEqual(len(res_timeout["3_Conciliado_Desmembrado"]), 0)
        self.assertEqual(len(res_timeout["5_Divergencias_Pendentes"]), 4)


class TestTratamentoArquivosCorrompidos(unittest.TestCase):
    """Testa a validação e tratamento robusto de arquivos Excel/banco corrompidos (Fase 2.5)."""

    def setUp(self):
        self.temp_files = []

    def tearDown(self):
        for f in self.temp_files:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass

    def _criar_arquivo_temp(self, suffix=".xlsx", content=b""):
        with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tf:
            tf.write(content)
            self.temp_files.append(tf.name)
            return tf.name

    def test_validate_excel_arquivo_inexistente(self):
        """Arquivo inexistente deve levantar FileNotFoundError."""
        caminho_inexistente = "arquivo_totalmente_inexistente_12345.xlsx"
        with self.assertRaises(FileNotFoundError):
            DataCleaner.validate_excel(caminho_inexistente)

        is_valid, motivo = DataCleaner.is_excel_valid(caminho_inexistente)
        self.assertFalse(is_valid)
        self.assertIn("não encontrado", motivo.lower())

    def test_validate_excel_arquivo_vazio_zero_bytes(self):
        """Arquivo com tamanho 0 bytes deve levantar CorruptedFileError."""
        arquivo_vazio = self._criar_arquivo_temp(suffix=".xlsx", content=b"")
        with self.assertRaises(CorruptedFileError) as ctx:
            DataCleaner.validate_excel(arquivo_vazio)
        self.assertIn("0 bytes", str(ctx.exception).lower())

        is_valid, motivo = DataCleaner.is_excel_valid(arquivo_vazio)
        self.assertFalse(is_valid)
        self.assertIn("vazio", motivo.lower())

    def test_validate_excel_arquivo_texto_falso_xlsx(self):
        """Arquivo de texto puro com extensão .xlsx deve ser rejeitado como ZIP inválido."""
        fake_xlsx = self._criar_arquivo_temp(suffix=".xlsx", content=b"Isto eh um texto simples fingindo ser excel.")
        with self.assertRaises(CorruptedFileError) as ctx:
            DataCleaner.validate_excel(fake_xlsx)
        self.assertTrue("inválid" in str(ctx.exception).lower() or "corrompido" in str(ctx.exception).lower())

        is_valid, motivo = DataCleaner.is_excel_valid(fake_xlsx)
        self.assertFalse(is_valid)

    def test_validate_excel_arquivo_zip_sem_estrutura_xlsx(self):
        """Arquivo ZIP válido mas sem a estrutura interna de planilha Excel ([Content_Types].xml) deve ser rejeitado."""
        import zipfile
        fake_zip = self._criar_arquivo_temp(suffix=".xlsx", content=b"")
        with zipfile.ZipFile(fake_zip, "w") as zf:
            zf.writestr("arquivo_aleatorio.txt", "conteúdo não excel")

        with self.assertRaises(CorruptedFileError) as ctx:
            DataCleaner.validate_excel(fake_zip)
        self.assertIn("estrutura interna", str(ctx.exception).lower())

        is_valid, motivo = DataCleaner.is_excel_valid(fake_zip)
        self.assertFalse(is_valid)

    def test_validate_excel_arquivo_zip_truncado(self):
        """Arquivo Excel com pacote ZIP truncado/corrompido deve levantar CorruptedFileError."""
        valido = self._criar_arquivo_temp(suffix=".xlsx")
        wb = openpyxl.Workbook()
        wb.active.title = "Dados"
        wb.active.append(["A", "B", "C"])
        wb.save(valido)

        with open(valido, "rb") as f:
            data = f.read()

        truncado = self._criar_arquivo_temp(suffix=".xlsx", content=data[: len(data) // 2])

        with self.assertRaises(CorruptedFileError):
            DataCleaner.validate_excel(truncado)

        is_valid, _ = DataCleaner.is_excel_valid(truncado)
        self.assertFalse(is_valid)

    def test_validate_excel_arquivo_valido(self):
        """Arquivo Excel legítimo passa na validação sem erros."""
        valido = self._criar_arquivo_temp(suffix=".xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Planilha1"
        ws.append(["Header1", "Header2"])
        ws.append(["Valor1", "Valor2"])
        wb.save(valido)

        # Não deve levantar exceção
        DataCleaner.validate_excel(valido)
        is_valid, motivo = DataCleaner.is_excel_valid(valido)
        self.assertTrue(is_valid)
        self.assertEqual(motivo, "")

    def test_clean_argos_rejeita_arquivo_corrompido(self):
        """DataCleaner.clean_argos levanta CorruptedFileError ao receber arquivo corrompido."""
        corrupto = self._criar_arquivo_temp(suffix=".xlsx", content=b"conteudo_lixo_123")
        with self.assertRaises(CorruptedFileError):
            DataCleaner.clean_argos(corrupto)

    def test_clean_bank_rejeita_arquivo_corrompido(self):
        """DataCleaner.clean_bank levanta CorruptedFileError ao receber arquivo corrompido."""
        corrupto = self._criar_arquivo_temp(suffix=".xlsx", content=b"conteudo_lixo_456")
        with self.assertRaises(CorruptedFileError):
            DataCleaner.clean_bank(corrupto)

    def test_validate_file_pdf_corrompido(self):
        """DataCleaner.validate_file levanta CorruptedFileError ao receber arquivo PDF corrompido."""
        fake_pdf = self._criar_arquivo_temp(suffix=".pdf", content=b"nao eh um pdf valido %PDF lixo")
        with self.assertRaises(CorruptedFileError):
            DataCleaner.validate_file(fake_pdf)

    def test_api_conciliar_retorna_400_quando_arquivo_corrompido(self):
        """A API FastAPI POST /api/conciliar retorna HTTP 400 com mensagem detalhada se um arquivo enviado estiver corrompido."""
        from fastapi.testclient import TestClient
        from app import app, get_current_user
        import io

        app.dependency_overrides[get_current_user] = lambda: {"id": "test_user_corrupted"}
        client = TestClient(app)

        try:
            # Cria um arquivo válido de banco em memória
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Data", "Histórico", "Valor"])
            ws.append(["01/06/2026", "CREDITO", 100.0])
            banco_bytes = io.BytesIO()
            wb.save(banco_bytes)
            banco_bytes.seek(0)

            # Arquivo Argos totalmente corrompido (bytes inválidos)
            argos_corrupt_bytes = io.BytesIO(b"ISTO_NAO_EH_EXCEL")

            response = client.post(
                "/api/conciliar",
                files=[
                    ("argos_files", ("argos_corrompido.xlsx", argos_corrupt_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
                    ("banco_files", ("banco_valido.xlsx", banco_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
                ],
                data={"banco_nome": "generico"},
            )

            self.assertEqual(response.status_code, 400)
            data_resp = response.json()
            self.assertIn("detail", data_resp)
            self.assertTrue(
                "corrompido" in data_resp["detail"].lower()
                or "integridade" in data_resp["detail"].lower()
            )
        finally:
            app.dependency_overrides.clear()


class TestAnalyticsService(unittest.TestCase):
    """
    Testes unitários e de integração para AnalyticsService (Fase 3.1)
    e para o endpoint GET /api/analytics/resumo da API FastAPI.
    """

    def setUp(self):
        from datetime import datetime, timedelta
        hoje = datetime.now()
        dez_dias_atras = hoje - timedelta(days=10)
        sessenta_dias_atras = hoje - timedelta(days=60)

        self.registros_amostra = [
            {
                "id": "rec-1",
                "data_processamento": hoje.isoformat(),
                "periodo": "01/09/2026 a 15/09/2026 | Bancos: ITAU, BRADESCO",
                "taxa_sucesso": 100.0,
                "perfeitos": 8,
                "historico": 2,
                "desmembrados": 0,
                "saidas_estornos": 1,
                "divergencias": 0,
                "anotacao": "Conciliação quinzenal Itaú e Bradesco sem erros",
            },
            {
                "id": "rec-2",
                "data_processamento": dez_dias_atras.isoformat(),
                "periodo": "01/08/2026 a 31/08/2026 | Bancos: SANTANDER",
                "taxa_sucesso": 75.0,
                "perfeitos": 5,
                "historico": 1,
                "desmembrados": 0,
                "saidas_estornos": 0,
                "divergencias": 2,
                "anotacao": "Auditoria de rotina Santander com pequenas divergencias",
            },
            {
                "id": "rec-3",
                "data_processamento": sessenta_dias_atras.isoformat(),
                "periodo": "01/07/2026 a 31/07/2026 | Bancos: ITAU",
                "taxa_sucesso": 40.0,
                "perfeitos": 2,
                "historico": 0,
                "desmembrados": 0,
                "saidas_estornos": 0,
                "divergencias": 3,
                "anotacao": "Lote antigo com pendencias",
            },
        ]

    def test_extrair_bancos_de_periodo(self):
        """Valida extração correta de bancos a partir da string do período."""
        p1 = "01/01/2026 a 31/01/2026 | Bancos: ITAU, BRADESCO, SANTANDER"
        bancos = AnalyticsService.extrair_bancos_de_periodo(p1)
        self.assertEqual(bancos, ["ITAU", "BRADESCO", "SANTANDER"])

        # Sem sufixo de bancos ou vazio
        self.assertEqual(AnalyticsService.extrair_bancos_de_periodo("01/01/2026 a 31/01/2026"), [])
        self.assertEqual(AnalyticsService.extrair_bancos_de_periodo(""), [])
        self.assertEqual(AnalyticsService.extrair_bancos_de_periodo(None), [])

    def test_filtrar_registros_por_banco(self):
        """Valida filtragem precisa por nome do banco (insensível a maiúsculas)."""
        # Filtrar por ITAU (deve retornar rec-1 e rec-3)
        res_itau = AnalyticsService.filtrar_registros(self.registros_amostra, banco="ITAU")
        self.assertEqual(len(res_itau), 2)
        self.assertEqual({r["id"] for r in res_itau}, {"rec-1", "rec-3"})

        # Filtrar por SANTANDER (deve retornar rec-2)
        res_sant = AnalyticsService.filtrar_registros(self.registros_amostra, banco="santander")
        self.assertEqual(len(res_sant), 1)
        self.assertEqual(res_sant[0]["id"], "rec-2")

        # Filtrar por 'todos' ou vazio (deve retornar todos os 3)
        self.assertEqual(len(AnalyticsService.filtrar_registros(self.registros_amostra, banco="todos")), 3)
        self.assertEqual(len(AnalyticsService.filtrar_registros(self.registros_amostra, banco="")), 3)

        # Filtrar por banco inexistente
        self.assertEqual(len(AnalyticsService.filtrar_registros(self.registros_amostra, banco="NUBANK")), 0)

    def test_filtrar_registros_por_periodo_dias(self):
        """Valida corte cronológico por número de dias a partir de hoje."""
        # Últimos 15 dias: deve trazer rec-1 (hoje) e rec-2 (10 dias atrás), excluindo rec-3 (60 dias atrás)
        res_15d = AnalyticsService.filtrar_registros(self.registros_amostra, dias=15)
        self.assertEqual(len(res_15d), 2)
        self.assertEqual({r["id"] for r in res_15d}, {"rec-1", "rec-2"})

        # Últimos 5 dias: apenas rec-1 (hoje)
        res_5d = AnalyticsService.filtrar_registros(self.registros_amostra, dias=5)
        self.assertEqual(len(res_5d), 1)
        self.assertEqual(res_5d[0]["id"], "rec-1")

        # Últimos 90 dias: traz todos os 3
        res_90d = AnalyticsService.filtrar_registros(self.registros_amostra, dias=90)
        self.assertEqual(len(res_90d), 3)

    def test_filtrar_registros_por_tipo_sucesso(self):
        """Valida filtro por faixas de performance e existência de divergências."""
        # Alta (>= 80%): rec-1 (100%)
        res_alta = AnalyticsService.filtrar_registros(self.registros_amostra, tipo_sucesso="alta")
        self.assertEqual(len(res_alta), 1)
        self.assertEqual(res_alta[0]["id"], "rec-1")

        # Média (50% a 79%): rec-2 (75%)
        res_media = AnalyticsService.filtrar_registros(self.registros_amostra, tipo_sucesso="media")
        self.assertEqual(len(res_media), 1)
        self.assertEqual(res_media[0]["id"], "rec-2")

        # Baixa (< 50%): rec-3 (40%)
        res_baixa = AnalyticsService.filtrar_registros(self.registros_amostra, tipo_sucesso="baixa")
        self.assertEqual(len(res_baixa), 1)
        self.assertEqual(res_baixa[0]["id"], "rec-3")

        # Perfeito (100% sem divergências): rec-1
        res_perfeito = AnalyticsService.filtrar_registros(self.registros_amostra, tipo_sucesso="perfeito")
        self.assertEqual(len(res_perfeito), 1)
        self.assertEqual(res_perfeito[0]["id"], "rec-1")

        # Com divergências: rec-2 (2 divs) e rec-3 (3 divs)
        res_div = AnalyticsService.filtrar_registros(self.registros_amostra, tipo_sucesso="divergencias")
        self.assertEqual(len(res_div), 2)
        self.assertEqual({r["id"] for r in res_div}, {"rec-2", "rec-3"})

    def test_filtrar_registros_por_busca_texto(self):
        """Valida busca textual por ID, período ou anotação."""
        res_anotacao = AnalyticsService.filtrar_registros(self.registros_amostra, busca="auditoria")
        self.assertEqual(len(res_anotacao), 1)
        self.assertEqual(res_anotacao[0]["id"], "rec-2")

        res_id = AnalyticsService.filtrar_registros(self.registros_amostra, busca="rec-3")
        self.assertEqual(len(res_id), 1)
        self.assertEqual(res_id[0]["id"], "rec-3")

    def test_calcular_metricas_e_resiliencia_vazio(self):
        """Calcula métricas agregadas consolidadas e testa resiliência para lista vazia."""
        metricas = AnalyticsService.calcular_metricas(self.registros_amostra)

        # Total conciliados: (8+2+0+1) + (5+1+0+0) + (2+0+0+0) = 11 + 6 + 2 = 19
        # Total divergências: 0 + 2 + 3 = 5
        # Total transações: 19 + 5 = 24
        # Taxa média ponderada: 19 / 24 * 100 = 79.17%
        self.assertEqual(metricas["total_conciliacoes"], 3)
        self.assertEqual(metricas["total_conciliados"], 19)
        self.assertEqual(metricas["total_divergencias"], 5)
        self.assertEqual(metricas["total_transacoes"], 24)
        self.assertAlmostEqual(metricas["taxa_media_sucesso"], 79.17, places=1)
        self.assertIn("ITAU", metricas["bancos_detectados"])
        self.assertIn("BRADESCO", metricas["bancos_detectados"])
        self.assertIn("SANTANDER", metricas["bancos_detectados"])

        # Caso lista vazia (não deve dar divisão por zero)
        vazio = AnalyticsService.calcular_metricas([])
        self.assertEqual(vazio["total_conciliacoes"], 0)
        self.assertEqual(vazio["taxa_media_sucesso"], 0.0)
        self.assertEqual(vazio["total_transacoes"], 0)
        self.assertEqual(vazio["bancos_detectados"], [])

    def test_calcular_series_temporal(self):
        """Testa geração da série temporal ordenada e formatação de pontos."""
        series = AnalyticsService.calcular_series_temporal(self.registros_amostra, max_pontos=2)
        # Deve respeitar max_pontos limitando aos 2 mais recentes (rec-2 e rec-1)
        self.assertEqual(len(series), 2)
        self.assertEqual(series[0]["id"], "rec-2")
        self.assertEqual(series[1]["id"], "rec-1")
        self.assertIn("periodo_label", series[0])
        self.assertIn("taxa_sucesso", series[0])

        # Lista vazia retorna lista vazia
        self.assertEqual(AnalyticsService.calcular_series_temporal([]), [])

    def test_gerar_dashboard_analytics_completo(self):
        """Valida que o pacote gerado por gerar_dashboard_analytics contém toda a estrutura esperada pelo frontend."""
        dashboard = AnalyticsService.gerar_dashboard_analytics(
            registros=self.registros_amostra,
            banco="ITAU",
            dias=30,
        )
        self.assertIn("filtros_aplicados", dashboard)
        self.assertEqual(dashboard["filtros_aplicados"]["banco"], "ITAU")
        self.assertEqual(dashboard["filtros_aplicados"]["dias"], 30)
        self.assertIn("metricas", dashboard)
        self.assertIn("series_temporal", dashboard)
        self.assertIn("bancos_disponiveis", dashboard)
        self.assertEqual(dashboard["total_registros_brutos"], 3)
        self.assertEqual(dashboard["total_registros_filtrados"], 1)  # apenas rec-1 atende ITAU + 30 dias

    def test_api_endpoint_analytics_resumo(self):
        """Testa a integração via HTTP com o endpoint GET /api/analytics/resumo."""
        from fastapi.testclient import TestClient
        from app import app, get_current_user

        app.dependency_overrides[get_current_user] = lambda: {"id": "test_analytics_user"}
        client = TestClient(app)

        try:
            # 1. Chamada geral sem parâmetros
            resp = client.get("/api/analytics/resumo")
            self.assertEqual(resp.status_code, 200)
            data = resp.json()
            self.assertIn("metricas", data)
            self.assertIn("series_temporal", data)
            self.assertIn("bancos_disponiveis", data)
            self.assertIn("filtros_aplicados", data)
            self.assertIn("total_geral_historico", data)

            # 2. Chamada com filtros específicos (banco e dias)
            resp_filtro = client.get("/api/analytics/resumo?banco=ITAU&dias=90&tipo_sucesso=todos")
            self.assertEqual(resp_filtro.status_code, 200)
            data_filtro = resp_filtro.json()
            self.assertEqual(data_filtro["filtros_aplicados"]["banco"], "ITAU")
            self.assertEqual(data_filtro["filtros_aplicados"]["dias"], 90)
        finally:
            app.dependency_overrides.clear()


if __name__ == "__main__":
    unittest.main(verbosity=2)


