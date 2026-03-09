import streamlit as st
import pandas as pd
from io import BytesIO
import tempfile
import os

try:
    import processador 
except ModuleNotFoundError:
    st.error("Erro: O arquivo 'processador.py' não foi encontrado na mesma pasta do app.")
    st.stop()

#Configuração da página
st.set_page_config(layout="wide", page_title="Automação PQ 10", page_icon="👩‍💻")

#Cabeçalho
col1, col2 = st.columns([1, 4])
with col1:
    try:
        st.image("logo.png", width=150) 
    except FileNotFoundError:
        st.warning("Arquivo 'logo.png' não encontrado.")
with col2:
    st.title("PQ 10 (atualizações de normas/leis/decretos)")
    st.caption("Ferramenta para verificação automática de datas de atualização")

st.divider()

with st.container(border=True):
    st.subheader("Carregue o Documento")
    # Texto alterado para Excel
    st.write("Faça o upload do arquivo **Lista Mestra** (.xlsx) para iniciar.")
    
    uploaded_file = st.file_uploader(
        "Selecione o arquivo .xlsx da sua máquina", 
        type=["xlsx"],  # Mudança principal aqui
        label_visibility="collapsed"
    )

if uploaded_file is not None:
    with st.container(border=True):
        st.subheader("Inicie a Verificação")
        st.info(f"Arquivo carregado: **{uploaded_file.name}**")
        
        if st.button("🔍 Iniciar Verificação Agora", type="primary", use_container_width=True):
            # Salva como .xlsx temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                temp_path = tmp_file.name

            with st.spinner('Aguarde, os sites estão sendo acessados...'):
                try:
                    df_resultado, nome_arquivo_excel = processador.executar_verificacao(temp_path)
                    
                    st.success("Verificação concluída com sucesso!")

                    with st.container(border=True):
                        st.subheader("Baixe os Resultados")
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_resultado.to_excel(writer, index=False, sheet_name='Verificacao')
                        excel_bytes = output.getvalue()

                        st.write("Prévia dos resultados:")
                        st.dataframe(df_resultado)
                        st.divider()

                        st.download_button(
                            label="📥 Baixar Planilha de Resultados (.xlsx)",
                            data=excel_bytes,
                            file_name=nome_arquivo_excel,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Ocorreu um erro durante o processamento:")
                    st.exception(e)
                finally:
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

st.divider()
st.write("Desenvolvido para otimizar esse processo.")