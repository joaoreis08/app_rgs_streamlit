import pandas as pd
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
import io

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

# --- Função Principal de Geração (COM AS ALTERAÇÕES) ---
def criar_documento_word(df: pd.DataFrame):
    df.rename(columns={
        'Órgão': 'Orgao', 'Iniciativa': 'Iniciativa', 'Status Informado': 'Status_Informado',
        'Ação': 'Acao', 'Programa': 'Programa', 'Início Realizado': 'Inicio_Realizado',
        'Término Realizado': 'Termino_Realizado', 'RGS-GGGE': 'RGS_GGGE',
        'Localização Geográfica': 'Localizacao_Geografica', 'Objetivo Estratégico': 'Objetivo_Estrategico'
    }, inplace=True, errors='ignore')

    colunas_necessarias = ['Orgao', 'Iniciativa', 'Status_Informado', 'Acao', 'Programa',
                           'Inicio_Realizado', 'Termino_Realizado', 'RGS_GGGE',
                           'Localizacao_Geografica', 'Objetivo_Estrategico']
    df2 = df[colunas_necessarias]

    df2[['Inicio_Realizado', 'Termino_Realizado']] = df2[['Inicio_Realizado', 'Termino_Realizado']].apply(
        lambda x: pd.to_datetime(x, errors='coerce', dayfirst=True)
    )
    ordem_temas = {
        "CONHECIMENTO E INOVAÇÃO": 1, "SAÚDE E QUALIDADE DE VIDA": 2,
        "DESENVOLVIMENTO SUSTENTÁVEL": 3, "SEGURANÇA E CIDADANIA": 4,
        "Gestão, Transparência e Participação": 5
    }
    df2['Ordem_Tema'] = df2['Objetivo_Estrategico'].map(ordem_temas).fillna(99)
    df2 = df2.sort_values(by=['Ordem_Tema', 'Orgao'])

    doc = Document()
    orgao_anterior = None

    for row in df2.itertuples(index=False):
        cor = cores_por_tema.get(row.Objetivo_Estrategico, "D3D3D3")
        
        if row.Orgao != orgao_anterior:
            if orgao_anterior is not None:
                doc.add_page_break()
            p_orgao = doc.add_paragraph()
            p_orgao.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_orgao.paragraph_format.space_before = Pt(0)
            p_orgao.paragraph_format.space_after = Pt(0)
            run = p_orgao.add_run(str(row.Orgao).upper())
            run.font.name = 'Gilroy ExtraBold'
            run.font.size = Pt(12)
            run.bold = True
            run.font.color.rgb = RGBColor(0, 32, 96)
            set_paragraph_background(p_orgao, 'D3D3D3')
            orgao_anterior = row.Orgao

        p_programa = doc.add_paragraph()
        p_programa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_programa.paragraph_format.space_before = Pt(0)
        p_programa.paragraph_format.space_after = Pt(0)
        run = p_programa.add_run(str(row.Programa).upper())
        run.font.name = 'Gilroy ExtraBold'
        run.font.size = Pt(12)
        # --- ALTERAÇÃO SOLICITADA (já estava correto, apenas confirmando) ---
        # Garante que o texto do programa fique em negrito
        run.bold = True 
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_paragraph_background(p_programa, cor)

        p_acao = doc.add_paragraph()
        p_acao.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_acao.paragraph_format.space_before = Pt(0)
        p_acao.paragraph_format.space_after = Pt(0)
        # --- ALTERAÇÃO SOLICITADA ---
        # Altera o texto para Title Case (Primeira letra de cada palavra maiúscula)
        run = p_acao.add_run(str(row.Acao).title())
        run.font.name = 'Gilroy Light'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(255, 255, 255)
        set_paragraph_background(p_acao, cor)
        
        doc.add_paragraph()

        # ... (o resto do código da tabela continua igual) ...
        status_imagem = 'imagens/concluído.png' if row.Status_Informado == 'CONCLUÍDO' else 'imagens/em_excecucao.png'
        status_texto_label = 'Data de Entrega:' if row.Status_Informado == 'CONCLUÍDO' else 'Data de Início:'
        prazo = row.Termino_Realizado if row.Status_Informado == 'CONCLUÍDO' else row.Inicio_Realizado
        icone_localizacao_path = 'imagens/localização.png'
        icone_calendario_path = 'imagens/calendário.png'

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
        run.font.color.rgb = RGBColor(0, 0, 0)
        cell_iniciativa.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_background(cell_iniciativa, 'D3D3D3')

        cell_status = table.cell(1, 0).merge(table.cell(1, 1))
        p_status = cell_status.paragraphs[0]
        run_label_status = p_status.add_run()
        run_label_status.add_picture(status_imagem, width=Inches(0.17))
        run_label_status.add_text('  Status:  ')
        run_label_status.font.name = 'Neutro Thin'
        run_label_status.font.size = Pt(9)
        run_label_status.font.color.rgb = RGBColor(0, 0, 0)
        run_valor_status = p_status.add_run(str(row.Status_Informado))
        run_valor_status.font.name = 'Neutro'
        run_valor_status.font.size = Pt(10)
        
        cell_data_merged = table.cell(1, 2).merge(table.cell(1, 4))
        p_data = cell_data_merged.paragraphs[0]
        run_data_label = p_data.add_run()
        run_data_label.add_picture(icone_calendario_path, width=Inches(0.17))
        run_data_label.add_text(f'  {status_texto_label} ')
        run_data_label.font.name = 'Neutro Thin'
        run_data_label.font.size = Pt(9)
        data_texto = prazo.strftime('%d/%m/%Y') if pd.notnull(prazo) else ''
        run_data_valor = p_data.add_run(f'\t\t {data_texto}')
        run_data_valor.font.name = 'Neutro'
        run_data_valor.font.size = Pt(10)

        cell_loc_label = table.cell(2, 0).merge(table.cell(2, 1))
        p_loc_label = cell_loc_label.paragraphs[0]
        run_loc_label = p_loc_label.add_run()
        run_loc_label.add_picture(icone_localizacao_path, width=Inches(0.17))
        run_loc_label.add_text('  Municípios Atendidos: ')
        run_loc_label.font.name = 'Neutro Thin'
        run_loc_label.font.size = Pt(9)

        cell_loc_valor = table.cell(2, 2).merge(table.cell(2, 4))
        p_loc_valor = cell_loc_valor.paragraphs[0]
        localizacao_texto = "" if pd.isnull(row.Localizacao_Geografica) else str(row.Localizacao_Geografica)
        run_loc_valor = p_loc_valor.add_run(localizacao_texto)
        run_loc_valor.font.name = 'Neutro'
        run_loc_valor.font.size = Pt(10)
        
        cell_rgs = table.cell(3, 0).merge(table.cell(3, 4))
        p_rgs = cell_rgs.paragraphs[0]
        run_rgs = p_rgs.add_run(str(row.RGS_GGGE))
        run_rgs.font.name = 'Neutro'
        run_rgs.font.size = Pt(9)

    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    
    return doc_io