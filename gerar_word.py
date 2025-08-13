# Arquivo: gerar_word.py

import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import io
# unicodedata e re removidos por não serem utilizados

# --- Dicionário de Cores (sem alteração) ---
cores_por_tema = {
    "CONHECIMENTO E INOVAÇÃO": "4400FF",
    "SAÚDE E QUALIDADE DE VIDA": "ED282C",
    "SEGURANÇA E CIDADANIA": "FFB000",
    "DESENVOLVIMENTO SUSTENTÁVEL": "87D200",
    "Gestão, Transparência e Participação": "002060"
}

# --- Funções Auxiliares (sem alterações) ---
def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.replace("#", ""))
    tcPr.append(shd)

def set_paragraph_background(paragraph, color):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color.replace("#", ""))
    pPr.append(shd)

# --- FUNÇÃO PRINCIPAL CORRIGIDA ---
def criar_documentos_por_tema(df: pd.DataFrame):
    df.rename(columns={
        'Órgão': 'Orgao', 'Iniciativa': 'Iniciativa', 'Status Informado': 'Status_Informado',
        'Ação': 'Acao', 'Programa': 'Programa', 'Início Realizado': 'Inicio_Realizado',
        'Término Realizado': 'Termino_Realizado', 'RGS-GGGE': 'RGS_GGGE',
        'Localização Geográfica': 'Localizacao_Geografica', 'Objetivo Estratégico': 'Objetivo_Estrategico'
    }, inplace=True, errors='ignore')

    colunas_necessarias = ['Orgao', 'Iniciativa', 'Status_Informado', 'Acao', 'Programa',
                           'Inicio_Realizado', 'Termino_Realizado', 'RGS_GGGE',
                           'Localizacao_Geografica', 'Objetivo_Estrategico']
    df2 = df[colunas_necessarias].copy()

    df2[['Inicio_Realizado', 'Termino_Realizado']] = df2[['Inicio_Realizado', 'Termino_Realizado']].apply(
        lambda x: pd.to_datetime(x, errors='coerce', dayfirst=True)
    )
    status_filter = ['EM EXECUÇÃO', 'CONCLUÍDO', 'EM LICITAÇÃO', 'LICITAÇÃO CONCLUÍDA', 'OBRA EM LICITAÇÃO']
    df2 = df2[df2['Status_Informado'].isin(status_filter)]
    status_mapping = {'EM LICITAÇÃO': 'EM EXECUÇÃO', 'LICITAÇÃO CONCLUÍDA': 'EM EXECUÇÃO', 'OBRA EM LICITAÇÃO': 'EM EXECUÇÃO'}
    df2['Status_Informado'] = df2['Status_Informado'].replace(status_mapping)

    documentos_gerados = {}
    temas_unicos = df2['Objetivo_Estrategico'].unique()

    for tema in temas_unicos:
        if pd.isna(tema):
            continue

        df_tema = df2[df2['Objetivo_Estrategico'] == tema].sort_values(by='Orgao')
        
        if df_tema.empty: # Adicionado para pular temas sem ações após o filtro
            continue

        doc = Document()
        orgao_anterior = None
        cor = cores_por_tema.get(tema, "D3D3D3")

        for row in df_tema.itertuples(index=False):
            if row.Orgao != orgao_anterior:
                if orgao_anterior is not None:
                    doc.add_page_break()
                p_orgao = doc.add_paragraph()
                p_orgao.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_orgao.paragraph_format.space_after = Pt(0) 
                run = p_orgao.add_run(str(row.Orgao).upper())
                run.font.name = 'Gilroy ExtraBold'
                run.font.size = Pt(12)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 32, 96)
                set_paragraph_background(p_orgao, 'D3D3D3')
                orgao_anterior = row.Orgao

            p_info = doc.add_paragraph()
            p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_info.paragraph_format.space_before = Pt(0)
            p_info.paragraph_format.space_after = Pt(8)

            run_programa = p_info.add_run(str(row.Programa).upper() + '\n')
            run_programa.font.name = 'Gilroy ExtraBold'
            run_programa.font.size = Pt(12)
            run_programa.bold = True
            run_programa.font.color.rgb = RGBColor(255, 255, 255)

            run_acao = p_info.add_run(str(row.Acao).title())
            run_acao.font.name = 'Gilroy Light'
            run_acao.font.size = Pt(12)
            run_acao.font.color.rgb = RGBColor(255, 255, 255)
            
            set_paragraph_background(p_info, cor)

            # --- CORREÇÃO: Caminhos das imagens usando barras normais e nomes sem acentos ---
            status_imagem = 'imagens\concluído.png' if row.Status_Informado == 'CONCLUÍDO' else 'imagens\em_excecucao.png'
            status_texto_label = 'Data de Entrega:' if row.Status_Informado == 'CONCLUÍDO' else 'Data de Início:'
            prazo = row.Termino_Realizado if row.Status_Informado == 'CONCLUÍDO' else row.Inicio_Realizado
            icone_localizacao_path = 'imagens\localização.png'
            icone_calendario_path = 'imagens\calendário.png'

            table = doc.add_table(rows=4, cols=5)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = False

            cell_iniciativa = table.cell(0, 0).merge(table.cell(0, 4))
            p_iniciativa = cell_iniciativa.paragraphs[0]
            p_iniciativa.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p_iniciativa.add_run(str(row.Iniciativa))
            run.font.name = 'Gilroy ExtraBold'
            run.font.size = Pt(10)
            run.bold = True
            cell_iniciativa.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background(cell_iniciativa, 'D3D3D3')

            cell_status = table.cell(1, 0).merge(table.cell(1, 1))
            run_label_status = cell_status.paragraphs[0].add_run()
            run_label_status.add_picture(status_imagem, width=Inches(0.17))
            run_label_status.add_text('  Status:  ')
            run_valor_status = cell_status.paragraphs[0].add_run(str(row.Status_Informado))
            run_valor_status.font.name = 'Neutro'
            run_valor_status.font.size = Pt(10)
            
            cell_data_merged = table.cell(1, 2).merge(table.cell(1, 4))
            run_data_label = cell_data_merged.paragraphs[0].add_run()
            run_data_label.add_picture(icone_calendario_path, width=Inches(0.17))
            run_data_label.add_text(f'  {status_texto_label} ')
            run_data_label.font.name = 'Neutro Thin'
            run_data_label.font.size = Pt(9)
            data_texto = prazo.strftime('%d/%m/%Y') if pd.notnull(prazo) else ''
            run_data_valor = cell_data_merged.paragraphs[0].add_run(f'\t\t {data_texto}')
            run_data_valor.font.name = 'Neutro'
            run_data_valor.font.size = Pt(10)

            cell_loc_label = table.cell(2, 0).merge(table.cell(2, 1))
            run_loc_label = cell_loc_label.paragraphs[0].add_run()
            run_loc_label.add_picture(icone_localizacao_path, width=Inches(0.17))
            run_loc_label.add_text('  Municípios Atendidos: ')
            run_loc_label.font.name = 'Neutro Thin'
            run_loc_label.font.size = Pt(9)

            # --- CORREÇÃO: Texto da localização adicionado apenas uma vez ---
            cell_loc_valor = table.cell(2, 2).merge(table.cell(2, 4))
            localizacao_texto = "" if pd.isnull(row.Localizacao_Geografica) else str(row.Localizacao_Geografica)
            run_loc_valor = cell_loc_valor.paragraphs[0].add_run(localizacao_texto)
            run_loc_valor.font.name = 'Neutro'
            run_loc_valor.font.size = Pt(10)
            
            cell_rgs = table.cell(3, 0).merge(table.cell(3, 4))
            cell_rgs.paragraphs[0].add_run(str(row.RGS_GGGE))
            # Você pode querer definir a fonte aqui também para consistência
            run_rgs = cell_rgs.paragraphs[0].runs[0]
            run_rgs.font.name = 'Neutro'
            run_rgs.font.size = Pt(9)

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)
        
        documentos_gerados[tema] = doc_io

    return documentos_gerados
