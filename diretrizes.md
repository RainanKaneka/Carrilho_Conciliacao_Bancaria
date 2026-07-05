# Instruções para o Antigravity (Prompt de Geração)

A ferramenta geradora de código deve seguir estritamente as diretrizes abaixo para a implementação:

1. **Linguagem:** Python 3.9+
2. **Bibliotecas Permitidas:** `pandas`, `openpyxl`, `re`, `itertools`, `unittest`. Não utilize frameworks pesados ou APIs externas.
3. **Robustez:** Inclua blocos `try-except` na leitura de arquivos para capturar erros de codificação ou colunas ausentes, fornecendo mensagens limpas em português ao usuário.
4. **Performance:** Na regra de soma combinatória, restrinja o tamanho do chunk de dados para evitar estouro de memória ou travamentos (máximo de 3 combinações de linhas por subgrupo de datas).
5. **Independência de Interface:** O script deve rodar via CLI simples (linha de comando) ou buscando os arquivos em uma pasta fixa chamada `/dados/`.

Este documento serve como mapa de engenharia completo. Implemente o script contendo todas as classes especificadas e a suíte de testes unitários funcional.