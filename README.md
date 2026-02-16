# 📊 Automação para Saneamento de Planilhas Excel (Python)

Este projeto foi desenvolvido para automatizar o saneamento e a limpeza de bases de dados de inventário, otimizando processos manuais em operações de BPO.

## 🚀 Funcionalidades
* **Tratamento Automatizado**: Realiza a limpeza de exames com base em um arquivo de inventário de referência.
* **Geração de Logs**: Cria um resumo numérico detalhado de todo o processo de saneamento.
* **Relatório de Removidos**: Gera uma aba específica detalhando todos os itens que foram filtrados.

## 🛠️ Tecnologias e Bibliotecas
* **Python**
* **Pandas**: Para manipulação eficiente de grandes volumes de dados.
* **Openpyxl & XlsxWriter**: Motores para criação e formatação de arquivos Excel.

## 📋 Como utilizar
1. Instale as dependências necessárias:
   ```bash
   pip install pandas openpyxl xlsxwriter
[Inventário_AC.xlsx](https://github.com/user-attachments/files/25329247/Inventario_AC.xlsx)
[saneamento_excel.py](https://github.com/user-attachments/files/25329250/saneamento_excel.py)

import pandas as pd
import re
import os

def sanear_planilha(caminho_planilha_base, caminho_inventario_ac='Inventario_AC.xlsx'):
    """
    Executa o saneamento da planilha base de acordo com o inventário de exames.
    """
    try:
        # Carregar o Inventário_AC.xlsx
        inventario_ac = pd.read_excel(caminho_inventario_ac)
        ac_set = set(inventario_ac[inventario_ac['Classificação'] == 'Analises Clinicas']['Exam'].str.strip().str.upper())
    except FileNotFoundError:
        return "Erro: Arquivo Inventario_AC.xlsx não encontrado.", None
    except KeyError:
        return "Erro: Colunas 'Exam' ou 'Classificação' não encontradas em Inventario_AC.xlsx.", None

    try:
        # Carregar a planilha-base
        planilha_base = pd.read_excel(caminho_planilha_base)
    except FileNotFoundError:
        return f"Erro: Arquivo {caminho_planilha_base} não encontrado.", None

    # Preparar DataFrames para as abas de saída
    df_base_saneada = planilha_base.copy()
    df_removidos = pd.DataFrame(columns=['Admissão', 'Original Row', 'Removed Exams', 'Row_Removed'])

    rows_removed_only_ac = 0
    rows_partially_cleaned = 0
    total_rows = len(planilha_base)

    # Identificar a coluna de admissão
    coluna_admissao = None
    if 'Admissão' in planilha_base.columns:
        coluna_admissao = 'Admissão'
    elif 'Admissão: Código da Admissão' in planilha_base.columns:
        coluna_admissao = 'Admissão: Código da Admissão'
    else:
        return "Erro: Coluna 'Admissão' ou 'Admissão: Código da Admissão' não encontrada na planilha base.", None

    # Iterar sobre as linhas da planilha base
    for index, row in planilha_base.iterrows():
        exames_na_admissao_str = str(row.get('Exames na Admissão', ''))
        if not exames_na_admissao_str:
            continue

        exames_originais = [e.strip().upper() for e in exames_na_admissao_str.split(';') if e.strip()]
        exames_a_remover = [e for e in exames_originais if e in ac_set]
        exames_restantes = [e for e in exames_originais if e not in ac_set]

        admissao_codigo = row[coluna_admissao]

        if exames_a_remover:
            if len(exames_a_remover) == len(exames_originais): # Todos os tokens estão em AC_SET
                df_removidos.loc[len(df_removidos)] = {
                    'Admissão': admissao_codigo,
                    'Original Row': index + 2, # Header é linha 1, dados começam na 2
                    'Removed Exams': ';'.join(exames_a_remover),
                    'Row_Removed': True
                }
                df_base_saneada = df_base_saneada.drop(index)
                rows_removed_only_ac += 1
            elif len(exames_a_remover) > 0: # Tokens mistos
                df_removidos.loc[len(df_removidos)] = {
                    'Admissão': admissao_codigo,
                    'Original Row': index + 2,
                    'Removed Exams': ';'.join(exames_a_remover),
                    'Row_Removed': False
                }
                df_base_saneada.loc[index, 'Exames na Admissão'] = ';'.join(exames_restantes)
                rows_partially_cleaned += 1

    rows_remaining = len(df_base_saneada)

    # Criar a aba Log
    df_log = pd.DataFrame({
        'Métrica': ['Total Rows', 'Rows Removed (only AC)', 'Rows Partially Cleaned', 'Rows Remaining'],
        'Valor': [total_rows, rows_removed_only_ac, rows_partially_cleaned, rows_remaining]
    })

    # Salvar o resultado
    stem = re.sub(r'[^a-zA-Z0-9_]', '_', os.path.splitext(os.path.basename(caminho_planilha_base))[0])
    caminho_saida = f"{stem}_clean.xlsx"

    with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
        df_base_saneada.to_excel(writer, sheet_name='Base_Saneada', index=False)
        df_log.to_excel(writer, sheet_name='Log', index=False)
        df_removidos.to_excel(writer, sheet_name='Removidos', index=False)

    return caminho_saida, {
        'Total Rows': total_rows,
        'Rows Removed (only AC)': rows_removed_only_ac,
        'Rows Partially Cleaned': rows_partially_cleaned,
        'Rows Remaining': rows_remaining
    }

if __name__ == '__main__':
    inventario_file = 'Inventário_AC.xlsx'
    base_file = 'Autorizações Visão BKO SP-2026-02-05-08-59-38.xlsx'

    resultado, resumo = sanear_planilha(base_file, inventario_file)

    if resultado:
        print(f"Arquivo saneado salvo em: {resultado}")
        if resumo:
            print("Resumo:")
            for key, value in resumo.items():
                print(f"- {key}: {value}")
    else:
        print(f"Ocorreu um erro: {resumo}")


