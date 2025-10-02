import pdfplumber
import pandas as pd
import re
import os
from datetime import datetime
import xlsxwriter


def extrair_dados_pdf_versao_final(caminho_pdf):
    """
    Extrai dados de um PDF usando uma lógica de reconstrução de linhas
    robusta e simplificada, inspirada em métodos comprovadamente eficazes.
    """
    dados_cabecalho = {"report_datetime": None, "progress_percentage": None}

    # --- Parte 1: Extrair dados do cabeçalho com pdfplumber ---
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            page_one = pdf.pages[0]
            text_page_one = page_one.extract_text()
            match_date = re.search(r"Today\n(.*)", text_page_one)
            if match_date:
                dados_cabecalho["report_datetime"] = match_date.group(
                    1).strip()

            tables_p1 = page_one.extract_tables()
            for table in tables_p1:
                if table and table[-1] and 'PROGRESS' in table[-1][0]:
                    try:
                        dados_cabecalho['progress_percentage'] = float(
                            table[-1][-1])
                    except (ValueError, IndexError):
                        pass
                    break
    except Exception as e:
        print(f"Aviso: Não foi possível ler o cabeçalho. Erro: {e}")

    # --- Parte 2: Extrair todas as linhas de tabelas com pdfplumber ---
    todas_as_linhas = []
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            # Itera a partir da página 2 (índice 1)
            for page in pdf.pages[1:]:
                # Usa a extração de tabela padrão, que funciona bem como ponto de partida
                tabela = page.extract_table()
                if tabela:
                    # Adiciona todas as linhas da tabela da página à nossa lista geral
                    todas_as_linhas.extend(tabela)
    except Exception as e:
        print(f"Erro ao extrair tabelas do PDF com pdfplumber: {e}")
        return dados_cabecalho, pd.DataFrame()

    if not todas_as_linhas:
        return dados_cabecalho, pd.DataFrame()

    df_bruto = pd.DataFrame(todas_as_linhas)

    # --- Parte 3: Reconstrução e Limpeza ---

    # Encontra o índice da linha que serve como cabeçalho
    header_index = -1
    for i, row in df_bruto.iterrows():
        # A linha de cabeçalho é a que contém 'SEQ' e 'GROUP'
        if row.astype(str).str.contains('SEQ').any() and row.astype(str).str.contains('GROUP').any():
            header_index = i
            break

    if header_index == -1:
        print("Aviso: Cabeçalho da tabela de tarefas não encontrado.")
        return dados_cabecalho, pd.DataFrame()

    # Define os nomes das colunas e remove as linhas acima do cabeçalho
    colunas = ['PHASE', 'SEQ', 'GROUP', 'DESCRIPTION',
               'STATUS', 'EXTERNAL TASK', 'ORIG']
    df_bruto = df_bruto.drop(
        df_bruto.index[:header_index + 1]).reset_index(drop=True)
    df_bruto.columns = colunas[:len(df_bruto.columns)]

    # Lógica de recomposição de linhas
    dados_recompostos = []
    buffer_linha = None

    for _, row_series in df_bruto.iterrows():
        # Converte a Series para um dicionário para facilitar a manipulação
        row = row_series.to_dict()
        # Limpa o valor de SEQ para verificação
        seq = str(row.get('SEQ', '') or '').replace('\n', ' ').strip()

        # Se SEQ é um número, é o início de uma nova tarefa
        if seq.isdigit():
            # Se já tínhamos uma tarefa no buffer, salva ela na lista
            if buffer_linha:
                dados_recompostos.append(buffer_linha)
            # Inicia o buffer com a nova tarefa
            buffer_linha = row
            # Limpa quebras de linha de todos os campos
            for key in buffer_linha:
                buffer_linha[key] = str(
                    buffer_linha[key] or '').replace('\n', ' ').strip()

        # Se não for uma nova tarefa, mas já temos um buffer, é uma continuação
        elif buffer_linha:
            # Junta o texto de todas as colunas da linha de continuação na descrição
            texto_continuacao = ' '.join(
                str(v or '').replace('\n', ' ').strip()
                for k, v in row.items()
                if v and k not in ['SEQ', 'PHASE']  # Evita juntar lixo
            )
            buffer_linha['DESCRIPTION'] += ' ' + texto_continuacao.strip()

    # Adiciona a última tarefa que ficou no buffer
    if buffer_linha:
        dados_recompostos.append(buffer_linha)

    if not dados_recompostos:
        return dados_cabecalho, pd.DataFrame()

    df_final = pd.DataFrame(dados_recompostos)

    # --- Parte 4: Limpeza final ---
    df_final.drop(columns=['PHASE'], inplace=True, errors='ignore')
    df_final['DESCRIPTION'] = df_final['DESCRIPTION'].str.replace(
        r'\s{2,}', ' ', regex=True).str.strip()
    df_final.loc[df_final['STATUS'].isnull() | (
        df_final['STATUS'] == ''), 'STATUS'] = 'WAIT APPROVAL'

    colunas_ordenadas = ['GROUP', 'SEQ', 'DESCRIPTION',
                         'STATUS', 'EXTERNAL TASK', 'ORIG']
    df_final = df_final.reindex(columns=colunas_ordenadas).fillna('')

    return dados_cabecalho, df_final


# --- Bloco Principal de Execução ---
if __name__ == "__main__":
    # Coloque o nome do seu arquivo PDF aqui
    nome_arquivo_pdf = 'Customer_Report_19000277.pdf'

    try:
        if not os.path.exists(nome_arquivo_pdf):
            print(
                f"AVISO: Arquivo '{nome_arquivo_pdf}' não encontrado no diretório.")
        else:
            dados_cabecalho, df_tarefas = extrair_dados_pdf_versao_final(
                nome_arquivo_pdf)

            if not df_tarefas.empty:
                total_tarefas = len(df_tarefas)
                tarefas_fechadas = len(
                    df_tarefas[df_tarefas['STATUS'] == 'CLOSED'])
                percentual_conclusao = (
                    tarefas_fechadas / total_tarefas) * 100 if total_tarefas > 0 else 0

                nome_arquivo_excel = f'Dashboard_Final_{datetime.now().strftime("%Y-%m-%d")}.xlsx'

                with pd.ExcelWriter(nome_arquivo_excel, engine='xlsxwriter') as writer:
                    df_tarefas.to_excel(
                        writer, sheet_name='Dashboard', index=False, startrow=7)

                    workbook = writer.book
                    worksheet = writer.sheets['Dashboard']

                    formato_titulo = workbook.add_format(
                        {'bold': True, 'font_size': 14, 'align': 'left', 'valign': 'vcenter'})
                    formato_label = workbook.add_format({'bold': True})
                    formato_cabecalho_tabela = workbook.add_format(
                        {'bold': True, 'fg_color': '#4F81BD', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})

                    for col_num, value in enumerate(df_tarefas.columns.values):
                        worksheet.write(7, col_num, value,
                                        formato_cabecalho_tabela)

                    worksheet.write(
                        'A1', 'Dashboard de Acompanhamento de Tarefas', formato_titulo)
                    worksheet.write('A3', 'Progresso Geral:', formato_label)
                    worksheet.write('A4', 'Tarefas Fechadas:', formato_label)
                    worksheet.write('A5', 'Total de Tarefas:', formato_label)
                    worksheet.write('B3', f'{percentual_conclusao:.2f}%')
                    worksheet.write('B4', tarefas_fechadas)
                    worksheet.write('B5', total_tarefas)

                    # --- INÍCIO DA ALTERAÇÃO: Formatação Condicional para a Coluna STATUS ---
                    # Define os formatos de cor para cada status
                    formato_closed = workbook.add_format(
                        {'bg_color': '#C6EFCE'})  # Verde
                    formato_open = workbook.add_format(
                        {'bg_color': '#F2F2F2'})    # Cinza claro
                    formato_wait_approval = workbook.add_format(
                        {'bg_color': '#FFFF00'})  # Amarelo

                    # Tenta encontrar a coluna 'STATUS' para aplicar a formatação
                    try:
                        # Pega o índice da coluna (0-based) pelo nome
                        status_col_index = df_tarefas.columns.get_loc('STATUS')

                        # Define o intervalo de linhas para aplicar o formato.
                        # Os dados começam na linha 9 do Excel, que é o índice 8.
                        start_row = 8
                        end_row = start_row + len(df_tarefas) - 1

                        # Aplica a regra para 'CLOSED'
                        worksheet.conditional_format(start_row, status_col_index, end_row, status_col_index,
                                                     {'type': 'cell',
                                                      'criteria': '==',
                                                      'value': '"CLOSED"',
                                                      'format': formato_closed})

                        # Aplica a regra para 'OPEN'
                        worksheet.conditional_format(start_row, status_col_index, end_row, status_col_index,
                                                     {'type': 'cell',
                                                      'criteria': '==',
                                                      'value': '"OPEN"',
                                                      'format': formato_open})

                        # Aplica a regra para 'WAIT APPROVAL'
                        worksheet.conditional_format(start_row, status_col_index, end_row, status_col_index,
                                                     {'type': 'cell',
                                                      'criteria': '==',
                                                      'value': '"WAIT APPROVAL"',
                                                      'format': formato_wait_approval})
                    except KeyError:
                        print(
                            "Aviso: Coluna 'STATUS' não encontrada. A formatação condicional não foi aplicada.")
                    # --- FIM DA ALTERAÇÃO ---

                    worksheet.autofilter(
                        7, 0, 7 + len(df_tarefas), len(df_tarefas.columns) - 1)
                    worksheet.freeze_panes(8, 0)

                    for idx, col in enumerate(df_tarefas.columns):
                        if col == 'DESCRIPTION':
                            worksheet.set_column(idx, idx, 60)
                        else:
                            # Adicionado um try-except para o caso de colunas vazias
                            try:
                                max_len = max(df_tarefas[col].astype(
                                    str).map(len).max(), len(col))
                                worksheet.set_column(idx, idx, max_len + 2)
                            except (ValueError, KeyError):
                                worksheet.set_column(idx, idx, len(col) + 2)

                print(
                    f"\n✅ Dashboard profissional salvo em: '{nome_arquivo_excel}'")
            else:
                print("\n❌ Nenhum dado de tarefa foi extraído após o processamento.")
    except Exception as e:
        print(f"Ocorreu um erro inesperado durante a execução: {e}")
