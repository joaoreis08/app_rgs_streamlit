# Arquivo: app.py

import streamlit as st
import pandas as pd
from gerar_word import criar_documento_word # Importa a função do outro arquivo
from datetime import datetime

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Relatórios",
    page_icon="��",
    layout="centered"
)

# --- Interface da Aplicação ---
st.title("📄 Gerador de Relatório de Iniciativas")
st.write(
    "Faça o upload da sua planilha Excel de iniciativas para gerar "
    "automaticamente o relatório em formato Word (.docx)."
)

# --- Widget de Upload de Arquivo ---
uploaded_file = st.file_uploader(
    "Escolha o arquivo Excel (.xlsx)",
    type=['xlsx']
)

# --- Lógica Principal ---
if uploaded_file is not None:
    try:
        # Mostra uma mensagem enquanto o arquivo está sendo processado
        with st.spinner("Lendo a planilha e gerando o documento... Aguarde."):
            # Lê o arquivo Excel para um DataFrame do pandas
            df = pd.read_excel(uploaded_file)

            # Chama a função do outro arquivo para criar o documento
            documento_gerado_em_memoria = criar_documento_word(df)

        st.success("✅ Documento gerado com sucesso!")

        # Gera um nome de arquivo dinâmico com a data e hora
        nome_arquivo_final = f"relatorio_iniciativas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"

        # --- Botão de Download ---
        st.download_button(
            label="📥 Baixar Relatório (.docx)",
            data=documento_gerado_em_memoria,
            file_name=nome_arquivo_final,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
        st.warning(
            "Verifique se as colunas da sua planilha Excel correspondem ao "
            "modelo esperado (Órgão, Iniciativa, Status Informado, etc.)."
        )