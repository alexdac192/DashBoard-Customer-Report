import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
import re
import os
import xlsxwriter
from datetime import datetime


def extrair_dados_com_pre_processamento(pdf_path):
    """
    Abordagem final que replica a estratégia do Colab: pré-processa o PDF com Fitz
    para desenhar guias e depois extrai as tabelas com PDFPlumber.
    """

    # Etapa 1: Extração do Cabeçalho
    dados_cabecalho = {"report_datetime": None, "progress_percentage": None}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            page_one = pdf.pages[0]
            text_page_one = page_one.extract_text()
            match_date = re.search(r"Today\n(.*)", text_page_one)
            if match_date:
                dados_cabecalho["report_datetime"] = match_date.group(
                    1).strip()
    except Exception as e:
        print(f"Aviso: Não foi possível ler o cabeçalho. Erro: {e}")

    # Etapa 2: Pré-processamento com Fitz
    temp_pdf_path = "temp_processed.pdf"
    try:
        with fitz.open(pdf_path) as doc:
            for page in doc:
                for y in range(350, 800, 15):
                    page.draw_line(p1=(20, y), p2=(780, y),
                                   color=(0, 0, 0), width=0.5)
            doc.save(temp_pdf_path)
    except Exception as e:
        print(f"Erro durante o pré-processamento com Fitz: {e}")
        return dados_cabecalho, pd.DataFrame()

    # Etapa 3: Extração de Tabelas do PDF pré-processado
    tabelas_brutas = []
    try:
        with pdfplumber.open(temp_pdf_path) as pdf:
            for page in pdf.pages[1:]:
                tabelas_pagina = page.extract_tables()
                if tabelas_pagina:
                    tabelas_brutas.extend(tabelas_pagina)
    except Exception as e:
        print(f"Erro ao extrair tabelas com PDFPlumber: {e}")
        return dados_cabecalho, pd.DataFrame()
    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)

    if not tabelas_brutas:
        return dados_cabecalho, pd.DataFrame()

    # Etapa 4: Limpeza e Concatenação dos dados extraídos
    dfs = []
    header = ['PHASE', 'SEQ', 'GROUP', 'DESCRIPTION',
              'STATUS', 'EXTERNAL TASK', 'ORIG']
    for tabela in tabelas_brutas:
        if not tabela:
            continue  # Pula tabelas vazias
        df = pd.DataFrame(tabela)

        # ✅ CORREÇÃO APLICADA AQUI para evitar o erro de KeyError
        # Adiciona a verificação len(row) > 2 para ignorar linhas malformadas
        header_row_index = df[df.apply(lambda row: len(row) > 2 and 'SEQ' in str(
            row[1]) and 'GROUP' in str(row[2]), axis=1)].index

        if not header_row_index.empty:
            start_index = header_row_index[0] + 1
            df_data = df.iloc[start_index:].copy()
            df_data = df_data.iloc[:, :len(header)]
            df_data.columns = header
            dfs.append(df_data)

    if not dfs:
        return dados_cabecalho, pd.DataFrame()

    df_bruto = pd.concat(dfs, ignore_index=True)

    # Etapa 5: Pós-processamento para juntar linhas de descrição
    dados_corrigidos = []
    buffer_linha = {}
    for _, row in df_bruto.iterrows():
        linha_atual = {col: str(row[col]).replace(
            '\n', ' ').strip() if pd.notna(row[col]) else '' for col in header}
        if linha_atual['SEQ'].isdigit():
            if buffer_linha:
                dados_corrigidos.append(buffer_linha)
            buffer_linha = linha_atual
        elif buffer_linha:
            texto_continuacao = linha_atual['DESCRIPTION']
            if texto_continuacao:
                buffer_linha['DESCRIPTION'] += " " + texto_continuacao
    if buffer_linha:
        dados_corrigidos.append(buffer_linha)

    df_final = pd.DataFrame(dados_corrigidos)

    # Etapa 6: Limpeza Final
    df_final.loc[df_final['STATUS'] == '', 'STATUS'] = 'WAIT APPROVAL'
    df_final = df_final.drop(columns=['PHASE'], errors='ignore')
    colunas_ordenadas = ['GROUP', 'SEQ', 'DESCRIPTION',
                         'STATUS', 'EXTERNAL TASK', 'ORIG']
    df_final = df_final[[
        col for col in colunas_ordenadas if col in df_final.columns]]

    return dados_cabecalho, df_final


# --- Bloco Principal de Execução ---
if __name__ == "__main__":
    nome_arquivo_pdf = 'Customer_Report_19000277.pdf'
    if not os.path.exists(nome_arquivo_pdf):
        print(f"ERRO: Arquivo '{nome_arquivo_pdf}' não encontrado.")
    else:
        dados_cabecalho, df_tarefas = extrair_dados_com_pre_processamento(
            nome_arquivo_pdf)

        if not df_tarefas.empty:
            # Geração do Dashboard (sem alterações)
            total_tarefas = len(df_tarefas)
            tarefas_fechadas = len(
                df_tarefas[df_tarefas['STATUS'] == 'CLOSED'])
            percentual_conclusao = (
                tarefas_fechadas / total_tarefas) * 100 if total_tarefas > 0 else 0

            nome_arquivo_excel = f'Dashboard_KeyError_Corrigido_{datetime.now().strftime("%Y-%m-%d_%H%M%S")}.xlsx'
            writer = pd.ExcelWriter(nome_arquivo_excel, engine='xlsxwriter')
            df_tarefas.to_excel(
                writer, sheet_name='Dashboard', index=False, startrow=7)

            workbook, worksheet = writer.book, writer.sheets['Dashboard']
            formato_titulo = workbook.add_format(
                {'bold': True, 'font_size': 14, 'align': 'left'})
            formato_label = workbook.add_format({'bold': True})
            formato_cabecalho_tabela = workbook.add_format(
                {'bold': True, 'fg_color': '#4F81BD', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            for col_num, value in enumerate(df_tarefas.columns.values):
                worksheet.write(7, col_num, value, formato_cabecalho_tabela)
            worksheet.write(
                'A1', 'Dashboard de Acompanhamento de Tarefas', formato_titulo)
            worksheet.write('A3', 'Data do Relatório:', formato_label)
            worksheet.write('B3', dados_cabecalho.get(
                "report_datetime", "N/A"))
            worksheet.write('A4', 'Progresso Geral:', formato_label)
            worksheet.write('B4', f'{percentual_conclusao:.2f}%')
            worksheet.write('A5', 'Tarefas Fechadas:', formato_label)
            worksheet.write('B5', tarefas_fechadas)
            worksheet.write('A6', 'Total de Tarefas:', formato_label)
            worksheet.write('B6', total_tarefas)
            worksheet.freeze_panes(8, 0)
            for idx, col in enumerate(df_tarefas.columns):
                if col == 'DESCRIPTION':
                    worksheet.set_column(idx, idx, 70)
                else:
                    worksheet.set_column(idx, idx, max(
                        df_tarefas[col].astype(str).str.len().max(), len(col)) + 2)
            writer.close()
            print(
                f"\n✅ Dashboard profissional salvo com sucesso em: '{nome_arquivo_excel}'")
        else:
            print("\n❌ Nenhum dado de tarefa foi extraído após o processamento.")
