# -*- coding: utf-8 -*-
"""
test_conciliacao.py - Suite de Testes Unitarios para o Sistema de Conciliacao Bancaria
==========================================================================================

Testa as tres regras de negocio do ReconciliationEngine com dados mockados
extraidos das imagens reais do problema, alem de testar os componentes auxiliares
DataCleaner, as funcoes de conversao e a expressao regular.

Execucao:
    python -m unittest test_conciliacao.py -v
    python -m unittest discover -v

Estrutura de testes:
    TestConversorValorBR       -> funcao _converter_valor_br
    TestEncontraColuna         -> funcao _encontrar_coluna
    TestRegexValorHistorico    -> constante REGEX_VALOR_HISTORICO
    TestDataCleanerArgos       -> DataCleaner.clean_argos
    TestDataCleanerBanco       -> DataCleaner.clean_bank
    TestReconciliationEngine   -> ReconciliationEngine (todas as regras)
"""

import unittest
import pandas as pd
from datetime import datetime

from conciliacao import (
    DataCleaner,
    ReconciliationEngine,
    ExcelReporter,
    _converter_valor_br,
    _encontrar_coluna,
    REGEX_VALOR_HISTORICO,
)


class TestConversorValorBR(unittest.TestCase):
    """Testa a funcao auxiliar de conversao de valores monetarios brasileiros."""

    def test_formato_br_com_milhar(self):
        """Converte string com separador de milhar: 1.234,56 -> 1234.56."""
        self.assertAlmostEqual(_converter_valor_br("1.234,56"), 1234.56, places=2)

    def test_formato_br_simples(self):
        """Converte string simples: 200,00 -> 200.0."""
        self.assertAlmostEqual(_converter_valor_br("200,00"), 200.00, places=2)

    def test_formato_float_direto(self):
        """Float passado diretamente retorna sem modificacao."""
        self.assertAlmostEqual(_converter_valor_br(512.50), 512.50, places=2)

    def test_valor_inteiro(self):
        """Inteiro eh convertido para float."""
        self.assertAlmostEqual(_converter_valor_br(800), 800.00, places=2)

    def test_valor_com_simbolo_rs(self):
        """Remove o simbolo R$ antes da conversao."""
        self.assertAlmostEqual(_converter_valor_br("R$ 1.500,00"), 1500.00, places=2)

    def test_valor_none_retorna_none(self):
        """None como entrada retorna None."""
        self.assertIsNone(_converter_valor_br(None))

    def test_valor_nan_retorna_none(self):
        """NaN como entrada retorna None."""
        self.assertIsNone(_converter_valor_br(float("nan")))

    def test_string_invalida_retorna_none(self):
        """String nao numerica retorna None sem lancar excecao."""
        self.assertIsNone(_converter_valor_br("abc"))
        self.assertIsNone(_converter_valor_br("N/A"))


class TestEncontraColuna(unittest.TestCase):
    """Testa a busca case-insensitive de colunas por lista de candidatos."""

    def test_encontra_coluna_exata(self):
        """Encontra coluna com nome exatamente igual ao candidato."""
        resultado = _encontrar_coluna(["Data", "Valor", "Historico"], ["Data"])
        self.assertEqual(resultado, "Data")

    def test_encontra_coluna_case_insensitive(self):
        """Encontra coluna mesmo com diferenca de caixa."""
        resultado = _encontrar_coluna(["data", "valor"], ["DATA"])
        self.assertEqual(resultado, "data")

    def test_retorna_primeiro_candidato_disponivel(self):
        """Retorna o primeiro candidato encontrado na lista."""
        resultado = _encontrar_coluna(["Historico", "Credito"], ["X", "Credito", "Valor"])
        self.assertEqual(resultado, "Credito")

    def test_retorna_none_quando_nenhum_candidato_encontrado(self):
        """Retorna None se nenhum candidato estiver presente no DataFrame."""
        resultado = _encontrar_coluna(["Data", "Valor"], ["Credito", "VL CR"])
        self.assertIsNone(resultado)


class TestRegexValorHistorico(unittest.TestCase):
    """Testa a expressao regular de extracao de valores do campo Historico."""

    def test_captura_pix_no_valor_de(self):
        """Captura valor precedido por "pix no valor de"."""
        m = REGEX_VALOR_HISTORICO.search("pix no valor de 200,00 dia 22/06")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "200,00")

    def test_captura_valor_de(self):
        """Captura valor precedido por "valor de" com milhar."""
        m = REGEX_VALOR_HISTORICO.search("comprovante no valor de 1.500,00")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "1.500,00")

    def test_captura_pix_de(self):
        """Captura valor precedido por "pix de"."""
        m = REGEX_VALOR_HISTORICO.search("recebido pix de 512,50 confirmado")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "512,50")

    def test_captura_valor_isolado_sem_prefixo(self):
        """Captura valor no formato BR mesmo sem prefixo descritivo."""
        m = REGEX_VALOR_HISTORICO.search("transferencia 354,04 ok")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "354,04")

    def test_captura_case_insensitive(self):
        """Regex funciona com MAIUSCULAS (case-insensitive)."""
        m = REGEX_VALOR_HISTORICO.search("PIX NO VALOR DE 200,00 DIA 22/06")
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), "200,00")

    def test_nao_captura_sem_decimal_br(self):
        """Valor sem virgula decimal (ex: 800 reais) nao deve ser capturado.
        Essa regra evita falsos positivos na Regra 3 (valores desmembrados).
        """
        m = REGEX_VALOR_HISTORICO.search("comprovante no valor de 800 reais")
        self.assertIsNone(m, "Nao deveria capturar valor sem virgula decimal.")

    def test_nao_captura_historico_vazio(self):
        """Historico vazio retorna None sem lancar excecao."""
        m = REGEX_VALOR_HISTORICO.search("")
        self.assertIsNone(m)

    def test_nao_captura_texto_sem_numero(self):
        """Texto sem nenhum numero retorna None."""
        m = REGEX_VALOR_HISTORICO.search("comprovante enviado sem valor informado")
        self.assertIsNone(m)


class TestDataCleanerArgos(unittest.TestCase):
    """Testa a limpeza e normalizacao do DataFrame do Argos."""

    def _criar_df_argos_bruto(self) -> pd.DataFrame:
        """Cria DataFrame simulando a exportacao bruta do sistema Argos."""
        return pd.DataFrame({
            "Ordem": [1, 2],
            "Conta": ["001", "002"],
            "Evento": ["EV01", "EV02"],
            "Tipo Evento": ["TE1", "TE2"],
            "Parceiro Descricao": ["  junio silva  ", "  joao lima  "],
            "Data": ["16/06/2026", "22/06/2026"],
            "Valor": ["512,50", "199,17"],
            "Evento Descricao": ["Inclusao Vendas", "Inclusao Vendas"],
            "Historico": ["pix no valor de 512,50", "pix no valor de 200,00"],
        })

    def test_remove_colunas_desnecessarias(self):
        """Remove as colunas de metadados internos do Argos."""
        df_raw = self._criar_df_argos_bruto()
        df_limpo = DataCleaner.clean_argos(df_raw)
        for coluna in ["Ordem", "Conta", "Evento", "Tipo Evento"]:
            self.assertNotIn(coluna, df_limpo.columns,
                             f"Coluna de metadado '{coluna}' nao foi removida.")

    def test_converte_data_formato_br(self):
        """Converte a coluna Data de DD/MM/AAAA para datetime64.
        Usa is_datetime64_any_dtype para ser agnóstico à resolucao (ns vs us),
        garantindo compatibilidade com pandas 1.x e 2.x.
        """
        df_raw = self._criar_df_argos_bruto()
        df_limpo = DataCleaner.clean_argos(df_raw)
        self.assertTrue(
            pd.api.types.is_datetime64_any_dtype(df_limpo["Data"]),
            f"Coluna 'Data' deveria ser datetime64, mas e: {df_limpo['Data'].dtype}"
        )
        self.assertEqual(df_limpo["Data"].iloc[0], pd.Timestamp("2026-06-16"))

    def test_converte_valor_para_float(self):
        """Converte a coluna Valor de string BR para float64."""
        df_raw = self._criar_df_argos_bruto()
        df_limpo = DataCleaner.clean_argos(df_raw)
        self.assertEqual(df_limpo["Valor"].dtype, "float64")
        self.assertAlmostEqual(df_limpo["Valor"].iloc[0], 512.50, places=2)
        self.assertAlmostEqual(df_limpo["Valor"].iloc[1], 199.17, places=2)

    def test_normaliza_texto_para_uppercase_e_strip(self):
        """Normaliza coluna de texto: remove espacos e converte para UPPER."""
        df_raw = self._criar_df_argos_bruto()
        df_limpo = DataCleaner.clean_argos(df_raw)
        self.assertEqual(df_limpo["Parceiro Descricao"].iloc[0], "JUNIO SILVA")
        self.assertEqual(df_limpo["Parceiro Descricao"].iloc[1], "JOAO LIMA")

    def test_levanta_keyerror_coluna_obrigatoria_ausente(self):
        """Lanca KeyError com mensagem em portugues se Valor estiver ausente."""
        df_sem_valor = pd.DataFrame({
            "Parceiro Descricao": ["teste"],
            "Data": ["16/06/2026"],
            "Historico": ["ok"],
        })
        with self.assertRaises(KeyError):
            DataCleaner.clean_argos(df_sem_valor)

    def test_remove_linhas_com_data_invalida(self):
        """Linhas com data invalida sao removidas silenciosamente com aviso."""
        df_com_data_invalida = pd.DataFrame({
            "Parceiro Descricao": ["ok", "invalido"],
            "Data": ["16/06/2026", "nao-e-data"],
            "Valor": ["100,00", "200,00"],
            "Historico": ["", ""],
        })
        df_limpo = DataCleaner.clean_argos(df_com_data_invalida)
        self.assertEqual(len(df_limpo), 1)
        self.assertEqual(df_limpo["Valor"].iloc[0], 100.00)


class TestDataCleanerBanco(unittest.TestCase):
    """Testa a normalizacao dos extratos bancarios."""

    def _criar_df_banco_bruto(self) -> pd.DataFrame:
        """Cria DataFrame simulando exportacao bruta de extrato bancario."""
        return pd.DataFrame({
            "Data": ["16/06/2026", "22/06/2026", "17/06/2026", "15/06/2026"],
            "Credito": [512.50, 200.00, 800.00, -50.00],
            "Historico": ["CREDITO PIX", "CREDITO PIX", "CREDITO PIX", "DEBITO TED"],
        })

    def test_filtra_apenas_creditos(self):
        """Registros com valor negativo (debitos) sao descartados."""
        df_raw = self._criar_df_banco_bruto()
        df_limpo = DataCleaner.clean_bank(df_raw, bank_name="caixa")
        self.assertTrue(
            (df_limpo["valor_banco"] > 0).all(),
            "Registros com valor <= 0 nao foram filtrados."
        )
        self.assertEqual(len(df_limpo), 3, "Deveriam restar exatamente 3 creditos.")

    def test_renomeia_colunas_para_padrao_interno(self):
        """Colunas do banco sao renomeadas para o padrao interno."""
        df_raw = self._criar_df_banco_bruto()
        df_limpo = DataCleaner.clean_bank(df_raw, bank_name="caixa")
        for coluna_esperada in ["data_banco", "valor_banco", "descricao_banco"]:
            self.assertIn(coluna_esperada, df_limpo.columns,
                          f"Coluna '{coluna_esperada}' ausente no DataFrame normalizado.")

    def test_converte_data_banco_para_datetime(self):
        """Converte data_banco para datetime64.
        Usa is_datetime64_any_dtype para ser agnóstico à resolucao (ns vs us),
        garantindo compatibilidade com pandas 1.x e 2.x.
        """
        df_raw = self._criar_df_banco_bruto()
        df_limpo = DataCleaner.clean_bank(df_raw, bank_name="caixa")
        self.assertTrue(
            pd.api.types.is_datetime64_any_dtype(df_limpo["data_banco"]),
            f"Coluna 'data_banco' deveria ser datetime64, mas e: {df_limpo['data_banco'].dtype}"
        )

    def test_converte_valor_banco_para_float(self):
        """Converte valor_banco para float64."""
        df_raw = self._criar_df_banco_bruto()
        df_limpo = DataCleaner.clean_bank(df_raw, bank_name="caixa")
        self.assertEqual(df_limpo["valor_banco"].dtype, "float64")

    def test_levanta_keyerror_coluna_nao_identificada(self):
        """Lanca KeyError se coluna essencial nao puder ser identificada."""
        df_colunas_erradas = pd.DataFrame({
            "ColunaNaoExistente": ["2026-06-16"],
            "OutraColunaNaoExistente": [100.0],
            "MaisUmaColunaNaoExistente": ["PIX"],
        })
        with self.assertRaises(KeyError):
            DataCleaner.clean_bank(df_colunas_erradas, bank_name="caixa")


class TestReconciliationEngine(unittest.TestCase):
    """
    Suite principal de testes do motor de conciliacao.

    Simula os dados exatos extraidos das imagens reais do problema, conforme
    definido na especificacao tecnica funcional.

    Cenarios representados no setUp:
        - JUNIO SILVA:         Regra 1 - valor 512.50 bate exatamente em 16/06.
        - JOAO LIMA JR:        Regra 2 - valor 199.17 no Argos, historico diz 200,00
                               que bate com credito bancario de 200.00 em 22/06.
        - LATICINIO SERTANEJO: Regra 3 - dois lancamentos (445.96 + 354.04)
                               somados equivalem ao credito unico de 800.00 em 17/06.
    """

    def setUp(self):
        """Configuracao de dados mockados baseados nas imagens reais do problema."""
        self.argos_base = pd.DataFrame({
            "Parceiro Descricao": [
                "JUNIO SILVA DE MENDOCA",
                "JOAO DE OLIVEIRA LIMA JUNIOR",
                "LATICINIO SERTANEJO",
                "LATICINIO SERTANEJO",
            ],
            "Data": [
                pd.to_datetime("2026-06-16"),
                pd.to_datetime("2026-06-22"),
                pd.to_datetime("2026-06-16"),
                pd.to_datetime("2026-06-16"),
            ],
            "Valor": [512.50, 199.17, 445.96, 354.04],
            "Evento Descricao": ["Inclusao Vendas Atacado"] * 4,
            "Historico": [
                "pix no valor de 512,50 dia 16/06",
                "pix no valor de 200,00 dia 22/06",   # Caso de desvio - Regra 2
                "16/06 comprovante no valor de 800 reais.",  # Desmembrado p1 - Regra 3
                "16/06 comprovante no valor de 800 reais.",  # Desmembrado p2 - Regra 3
            ],
        })

        self.bank_base = pd.DataFrame({
            "data_banco": [
                pd.to_datetime("2026-06-16"),
                pd.to_datetime("2026-06-22"),
                pd.to_datetime("2026-06-17"),
            ],
            "valor_banco": [512.50, 200.00, 800.00],
            "descricao_banco": ["CREDITO PIX", "CREDITO PIX", "CREDITO PIX"],
        })

    # -----------------------------------------------------------------------
    # Testes da Regra 1: Match Exato
    # -----------------------------------------------------------------------

    def test_regra_1_match_perfeito(self):
        """Testa se o valor de 512.50 e conciliado perfeitamente na mesma data."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_exact()
        self.assertIn("Conciliado_Perfeito", res)
        self.assertFalse(
            res["Conciliado_Perfeito"].empty,
            "A aba Conciliado_Perfeito nao deve estar vazia."
        )
        self.assertTrue(
            any(res["Conciliado_Perfeito"]["Valor"] == 512.50),
            "O registro de 512.50 nao foi encontrado em Conciliado_Perfeito."
        )

    def test_regra_1_nao_concilia_fora_da_janela_temporal(self):
        """Credito bancario com +4 dias nao deve ser conciliado pela Regra 1.
        A janela e 0 a +3 dias, portanto +4 dias deve permanecer como divergencia.
        """
        argos_teste = pd.DataFrame({
            "Parceiro Descricao": ["TESTE"],
            "Data": [pd.to_datetime("2026-06-10")],
            "Valor": [100.00],
            "Historico": [""],
        })
        banco_teste = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-14")],  # +4 dias -> fora da janela
            "valor_banco": [100.00],
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_teste, banco_teste)
        res = engine.match_exact()
        self.assertTrue(
            res["Conciliado_Perfeito"].empty,
            "Credito com +4 dias nao deveria ser conciliado pela Regra 1."
        )

    def test_regra_1_concilia_dentro_da_janela_mais_3_dias(self):
        """Credito bancario com exatamente +3 dias deve ser conciliado pela Regra 1."""
        argos_teste = pd.DataFrame({
            "Parceiro Descricao": ["TESTE"],
            "Data": [pd.to_datetime("2026-06-16")],
            "Valor": [100.00],
            "Historico": [""],
        })
        banco_teste = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-19")],  # exatamente +3 dias
            "valor_banco": [100.00],
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_teste, banco_teste)
        res = engine.match_exact()
        self.assertFalse(
            res["Conciliado_Perfeito"].empty,
            "Credito com +3 dias deveria ser conciliado pela Regra 1."
        )

    # -----------------------------------------------------------------------
    # Testes da Regra 2: Regex Historico
    # -----------------------------------------------------------------------

    def test_regra_2_ajuste_historico(self):
        """Testa se a baixa de 199.17 encontra o credito de 200.00 lendo o historico."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_by_history_regex()
        self.assertIn("Conciliado_Via_Historico", res)
        self.assertFalse(
            res["Conciliado_Via_Historico"].empty,
            "A aba Conciliado_Via_Historico nao deve estar vazia."
        )
        self.assertTrue(
            any(res["Conciliado_Via_Historico"]["valor_real_banco"] == 200.00),
            "O valor real de 200.00 nao foi encontrado em Conciliado_Via_Historico."
        )

    def test_regra_2_nao_concilia_quando_historico_sem_valor_br(self):
        """Historico sem valor monetario BR nao deve gerar match via Regra 2.
        Garante que textos descritivos sem numero formatado nao causam falsos positivos.
        """
        argos_sem_regex = pd.DataFrame({
            "Parceiro Descricao": ["TESTE"],
            "Data": [pd.to_datetime("2026-06-22")],
            "Valor": [199.17],
            "Historico": ["COMPROVANTE ENVIADO SEM VALOR ESPECIFICADO"],
        })
        banco_teste = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-22")],
            "valor_banco": [200.00],
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_sem_regex, banco_teste)
        res = engine.match_by_history_regex()
        self.assertTrue(
            res["Conciliado_Via_Historico"].empty,
            "Nao deveria conciliar quando historico nao tem valor monetario BR."
        )

    def test_regra_2_nao_concilia_fora_da_janela_negativa(self):
        """Credito bancario com -3 dias nao deve ser conciliado (janela e -2 a +3)."""
        argos_teste = pd.DataFrame({
            "Parceiro Descricao": ["TESTE"],
            "Data": [pd.to_datetime("2026-06-22")],
            "Valor": [199.17],
            "Historico": ["pix no valor de 200,00"],
        })
        banco_teste = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-19")],  # -3 dias -> fora da janela
            "valor_banco": [200.00],
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_teste, banco_teste)
        res = engine.match_by_history_regex()
        self.assertTrue(
            res["Conciliado_Via_Historico"].empty,
            "Credito com -3 dias nao deveria ser conciliado pela Regra 2."
        )

    def test_regra_2_documenta_valor_original_e_real(self):
        """Resultado da Regra 2 deve conter as colunas de rastreabilidade."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_by_history_regex()
        df = res["Conciliado_Via_Historico"]
        if not df.empty:
            self.assertIn("valor_original_argos", df.columns)
            self.assertIn("valor_real_banco", df.columns)
            self.assertIn("historico_regex_extraido", df.columns)

    # -----------------------------------------------------------------------
    # Testes da Regra 3: Subset Sum (Desmembrados)
    # -----------------------------------------------------------------------

    def test_regra_3_valores_desmembrados(self):
        """Testa se as duas baixas (445.96 e 354.04) combinam com o credito de 800.00."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_split_sums()
        self.assertIn("Conciliado_Desmembrado", res)
        self.assertEqual(
            len(res["Conciliado_Desmembrado"]), 2,
            "Esperadas 2 linhas do Argos (as partes desmembradas) em Conciliado_Desmembrado."
        )

    def test_regra_3_soma_das_partes_correta(self):
        """A soma das partes desmembradas deve ser igual ao valor bancario."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        res = engine.match_split_sums()
        df = res["Conciliado_Desmembrado"]
        if not df.empty:
            soma = round(df["Valor"].sum(), 2)
            self.assertAlmostEqual(
                soma, 800.00, places=2,
                msg=f"A soma das partes ({soma}) nao equivale a 800.00."
            )

    def test_regra_3_nao_combina_alem_de_3_elementos(self):
        """Garante que combinacoes de 4+ elementos nao sao testadas.
        Isso protege contra explosao combinatoria (controle de performance).
        O limite MAX_COMBINACOES = 3 deve impedir o match quando a solucao
        exige 4 partes.
        """
        argos_4_partes = pd.DataFrame({
            "Parceiro Descricao": ["A", "B", "C", "D"],
            "Data": [pd.to_datetime("2026-06-16")] * 4,
            "Valor": [100.00, 200.00, 300.00, 400.00],
            "Historico": [""] * 4,
        })
        banco_4 = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-16")],
            "valor_banco": [1000.00],  # So encontrado com 4 partes
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_4_partes, banco_4)
        res = engine.match_split_sums()
        self.assertTrue(
            res["Conciliado_Desmembrado"].empty,
            "O motor nao deve testar combinacoes de 4 elementos (limite MAX_COMBINACOES=3)."
        )

    def test_regra_3_nao_combina_fora_da_janela_temporal(self):
        """Baixas fora da janela de +-3 dias do credito bancario nao sao combinadas."""
        argos_fora_janela = pd.DataFrame({
            "Parceiro Descricao": ["A", "B"],
            "Data": [pd.to_datetime("2026-06-01"), pd.to_datetime("2026-06-02")],
            "Valor": [445.96, 354.04],
            "Historico": ["", ""],
        })
        banco_17 = pd.DataFrame({
            "data_banco": [pd.to_datetime("2026-06-17")],  # 15+ dias depois
            "valor_banco": [800.00],
            "descricao_banco": ["PIX"],
        })
        engine = ReconciliationEngine(argos_fora_janela, banco_17)
        res = engine.match_split_sums()
        self.assertTrue(
            res["Conciliado_Desmembrado"].empty,
            "Baixas muito antigas nao devem ser combinadas com credito recente."
        )

    # -----------------------------------------------------------------------
    # Testes do Pipeline Completo
    # -----------------------------------------------------------------------

    def test_pipeline_completo_retorna_quatro_abas(self):
        """Testa se execute_pipeline retorna o dicionario com as 4 abas esperadas."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        resultado = engine.execute_pipeline()
        abas_esperadas = {
            "Conciliado_Perfeito",
            "Conciliado_Via_Historico",
            "Conciliado_Desmembrado",
            "Divergencias_Pendentes",
        }
        self.assertEqual(set(resultado.keys()), abas_esperadas)

    def test_pipeline_nao_duplica_registros(self):
        """Nenhum registro Argos pode ser conciliado mais de uma vez no pipeline.
        O total de registros conciliados nao pode exceder o tamanho do Argos.
        """
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        resultado = engine.execute_pipeline()
        total_conciliados = 0
        for chave in ["Conciliado_Perfeito", "Conciliado_Via_Historico", "Conciliado_Desmembrado"]:
            df = resultado[chave]
            if not df.empty and "Valor" in df.columns:
                total_conciliados += len(df)
        self.assertLessEqual(
            total_conciliados, len(self.argos_base),
            "Pipeline conciliou mais registros que o total do Argos (duplicacao detectada)."
        )

    def test_pipeline_todos_valores_sao_dataframes(self):
        """Todas as 4 chaves do resultado devem conter DataFrames validos."""
        engine = ReconciliationEngine(self.argos_base, self.bank_base)
        resultado = engine.execute_pipeline()
        for chave, df in resultado.items():
            self.assertIsInstance(
                df, pd.DataFrame,
                f"A chave '{chave}' nao contem um DataFrame valido."
            )

    def test_pipeline_sem_dados_nao_levanta_excecao(self):
        """Pipeline com DataFrames vazios nao deve lancar excecao.
        Garante robustez para casos onde nao ha registros a conciliar.
        """
        df_vazio_argos = pd.DataFrame({
            "Parceiro Descricao": pd.Series([], dtype="str"),
            "Data": pd.Series([], dtype="datetime64[ns]"),
            "Valor": pd.Series([], dtype="float64"),
            "Historico": pd.Series([], dtype="str"),
        })
        df_vazio_banco = pd.DataFrame({
            "data_banco": pd.Series([], dtype="datetime64[ns]"),
            "valor_banco": pd.Series([], dtype="float64"),
            "descricao_banco": pd.Series([], dtype="str"),
        })
        try:
            engine = ReconciliationEngine(df_vazio_argos, df_vazio_banco)
            resultado = engine.execute_pipeline()
            for df in resultado.values():
                self.assertIsInstance(df, pd.DataFrame)
        except Exception as exc:
            self.fail(f"Pipeline levantou excecao inesperada com DataFrames vazios: {exc}")


if __name__ == "__main__":
    unittest.main(verbosity=2)
