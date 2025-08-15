# Arquivo: app.py

import streamlit as st
import pandas as pd
# ALTERADO: Importa a nova função
from gerar_word import criar_documentos_por_tema
from datetime import datetime
import re
import unicodedata

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Relatórios",
    page_icon="🎨",
    layout="centered"
)

# --- Função para limpar o nome do arquivo ---
def limpar_nome_arquivo(nome):
    """Remove acentos, espaços e caracteres especiais para criar um nome de arquivo seguro."""
    # Normaliza para remover acentos
    s = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn')
    # Remove caracteres especiais, mantém letras, números e underscores
    s = re.sub(r'[^\w\s-]', '', s).strip()
    # Substitui espaços e hífens por underscore
    s = re.sub(r'[-\s]+', '_', s)
    return s.lower()

# --- Interface da Aplicação ---
st.title("📝 Relatório por Objetivo Estratrégico ")
st.write(
    "Faça o upload da sua planilha Excel de iniciativas para gerar "
    "automaticamente os relatórios em formato Word (.docx), separados por tema."
)

# --- Widget de Upload de Arquivo ---
uploaded_file = st.file_uploader(
    "Escolha o arquivo Excel (.xlsx)",
    type=['xlsx']
)

# --- Lógica Principal (ALTERADA) ---
if uploaded_file is not None:
    try:
        with st.spinner("Lendo a planilha e gerando os documentos... Aguarde."):
            df = pd.read_excel(uploaded_file)

            # Chama a nova função que retorna um dicionário de documentos
            documentos_por_tema = criar_documentos_por_tema(df)

        if not documentos_por_tema:
            st.warning("Nenhum tema (Objetivo Estratégico) encontrado na planilha para gerar relatórios.")
        else:
            st.success(f"✅ {len(documentos_por_tema)} documento(s) gerado(s) com sucesso!")
            st.markdown("---")
            
            # Itera sobre o dicionário para criar um botão de download para cada tema
            for tema, doc_io in documentos_por_tema.items():
                
                # Gera um nome de arquivo limpo e dinâmico
                tema_limpo = limpar_nome_arquivo(tema)
                nome_arquivo_final = f"relatorio_{tema_limpo}_{datetime.now().strftime('%Y%m%d')}.docx"
                
                # Cria um container ou coluna para cada botão para melhor espaçamento
                with st.container():
                    st.subheader(f"Objetivo Estratégico: {tema}")
                    st.download_button(
                        label=f"📥 Baixar Relatório: {tema}",
                        data=doc_io,
                        file_name=nome_arquivo_final,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        # Adiciona uma chave única para cada botão, importante para o Streamlit
                        key=f"download_{tema_limpo}"
                    )
                    st.markdown("---")

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        st.warning(
            "Verifique se as colunas da sua planilha Excel correspondem ao "
            "modelo esperado (Órgão, Iniciativa, Objetivo Estratégico, etc.)."
        )