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
[saneamento_excel_final.py](https://github.com/user-attachments/files/25329329/saneamento_excel_final.py)

import pandas as pd
import re
import os

def sanear_planilha(caminho_planilha_base, caminho_inventario_ac='Inventario_AC.xlsx'):
    """
    Executa o saneamento da planilha base mantendo a regra original de remoção de AC
    e adicionando as 3 novas regras de remoção completa de linhas.
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

    # Identificar a coluna de admissão (Lógica original)
    coluna_admissao = None
    if 'Admissão' in planilha_base.columns:
        coluna_admissao = 'Admissão'
    elif 'Admissão: Código da Admissão' in planilha_base.columns:
        coluna_admissao = 'Admissão: Código da Admissão'
    else:
        # Fallback para o primeiro índice se não encontrar os nomes
        coluna_admissao = planilha_base.columns[0]

    # --- NOVAS REGRAS DE REMOÇÃO (Fase 1: Filtros de Negócio) ---
    
    # Índices das colunas baseados no pedido do usuário:
    # Coluna D (Convênio) = Índice 3
    # Coluna E (Marca) = Índice 4
    # Coluna G (Tratativa) = Índice 6
    
    col_d = planilha_base.columns[3] if len(planilha_base.columns) > 3 else None
    col_e = planilha_base.columns[4] if len(planilha_base.columns) > 4 else None
    col_g = planilha_base.columns[6] if len(planilha_base.columns) > 6 else None

    # Aplicando as remoções de linhas conforme solicitado
    if col_g:
        # Regra 1: Remover linhas onde Coluna G é "Sim"
        planilha_base = planilha_base[planilha_base[col_g].astype(str).str.strip().str.upper() != 'SIM']

    if col_d and col_e:
        # Regra 2: CARE PLUS deve ficar apenas para marcas específicas
        marcas_care_permitidas = ['EXAME IMAGEM E LABORATÓRIO', 'SALOMÃO ZOPPI', 'DELBONI SALOMÃO ZOPPI']
        is_care_plus = planilha_base[col_d].astype(str).str.strip().str.upper() == 'CARE PLUS'
        is_marca_permitida_care = planilha_base[col_e].astype(str).str.strip().str.upper().isin(marcas_care_permitidas)
        # Remove se for Care Plus e NÃO for marca permitida
        planilha_base = planilha_base[~(is_care_plus & ~is_marca_permitida_care)]

        # Regra 3: Marca Memorial deve ficar apenas para convênios específicos
        marca_memorial = 'IMAGE MEMORIAL LABORATÓRIO E IMAGEM'
        convenios_memorial_permitidos = ['SULAMÉRICA SERVIÇOS DE SAÚDE', 'PETROBRAS AMS']
        is_memorial = planilha_base[col_e].astype(str).str.strip().str.upper() == marca_memorial
        is_convenio_permitido_memorial = planilha_base[col_d].astype(str).str.strip().str.upper().isin(convenios_memorial_permitidos)
        # Remove se for Memorial e NÃO for convênio permitido
        planilha_base = planilha_base[~(is_memorial & ~is_convenio_permitido_memorial)]

    # --- REGRA ORIGINAL (Fase 2: Saneamento de Exames AC) ---

    # Preparar DataFrames para as abas de saída
    df_base_saneada = planilha_base.copy()
    df_removidos = pd.DataFrame(columns=['Admissão', 'Original Row', 'Removed Exams', 'Row_Removed'])

    rows_removed_only_ac = 0
    rows_partially_cleaned = 0
    total_rows_after_filters = len(planilha_base)

    # Iterar sobre as linhas restantes após os novos filtros
    for index, row in planilha_base.iterrows():
        exames_na_admissao_str = str(row.get('Exames na Admissão', ''))
        if not exames_na_admissao_str or exames_na_admissao_str == 'nan' or exames_na_admissao_str.strip() == '':
            continue

        exames_originais = [e.strip().upper() for e in exames_na_admissao_str.split(';') if e.strip()]
        exames_a_remover = [e for e in exames_originais if e in ac_set]
        exames_restantes = [e for e in exames_originais if e not in ac_set]

        admissao_codigo = row[coluna_admissao]

        if exames_a_remover:
            if len(exames_a_remover) == len(exames_originais): # REGRA ORIGINAL: Todos são AC, remove a linha
                df_removidos.loc[len(df_removidos)] = {
                    'Admissão': admissao_codigo,
                    'Original Row': index + 2,
                    'Removed Exams': ';'.join(exames_a_remover),
                    'Row_Removed': True
                }
                df_base_saneada = df_base_saneada.drop(index)
                rows_removed_only_ac += 1
            else: # REGRA ORIGINAL: Sobraram exames, limpa apenas os AC da célula
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
        'Métrica': ['Total Rows (After Business Rules)', 'Rows Removed (Only AC)', 'Rows Partially Cleaned', 'Rows Remaining'],
        'Valor': [total_rows_after_filters, rows_removed_only_ac, rows_partially_cleaned, rows_remaining]
    })

    # Salvar o resultado
    stem = re.sub(r'[^a-zA-Z0-9_]', '_', os.path.splitext(os.path.basename(caminho_planilha_base))[0])
    caminho_saida = f"{stem}_clean.xlsx"

    with pd.ExcelWriter(caminho_saida, engine='xlsxwriter') as writer:
        df_base_saneada.to_excel(writer, sheet_name='Base_Saneada', index=False)
        df_log.to_excel(writer, sheet_name='Log', index=False)
        df_removidos.to_excel(writer, sheet_name='Removidos', index=False)

    return caminho_saida, {
        'Total Rows After Business Rules': total_rows_after_filters,
        'Rows Removed (Only AC)': rows_removed_only_ac,
        'Rows Remaining': rows_remaining
    }

if __name__ == '__main__':
    inventario_file = 'Inventário_AC.xlsx'
    base_file = 'Autorizações Visão IBM - Tell-2026-02-12-16-51-14.xlsx'

    if os.path.exists(base_file) and os.path.exists(inventario_file):
        resultado, resumo = sanear_planilha(base_file, inventario_file)
        if resultado:
            print(f"Arquivo saneado salvo em: {resultado}")
            if resumo:
                print("Resumo:")
                for key, value in resumo.items():
                    print(f"- {key}: {value}")
    else:
        print("Aguardando arquivos de entrada para execução.")
