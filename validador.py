import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl

st.set_page_config(page_title="Validador de Notas Fiscais", layout="centered")

st.title("📊 Validador de mercadorias/notas fiscais entre bases")

file1 = st.file_uploader("📤 Envie a 1ª planilha (base original)", type=["xls", "xlsx", "csv"])
file2 = st.file_uploader("📤 Envie a 2ª planilha (base de comparação)", type=["xls", "xlsx", "csv"])

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file, dtype=str)
    else:
        return pd.read_excel(file, dtype=str)

if file1 and file2:
    df1 = read_file(file1)
    df2 = read_file(file2)

    st.success("✅ Arquivos carregados com sucesso!")

    col1 = st.selectbox("🔹 Coluna 1ª planilha", df1.columns)
    col2 = st.selectbox("🔸 Coluna 2ª planilha", df2.columns)

    if st.button("✅ Validar e filtrar"):
        notas1 = df1[col1].dropna().astype(str).str.strip()
        notas2 = df2[col2].dropna().astype(str).str.strip()

        set2 = set(notas2)

        # Filtra a planilha 1 apenas com as notas que estão na base 2
        df_filtrado = df1[df1[col1].astype(str).str.strip().isin(set2)]

        st.success(f"✅ {len(df_filtrado)} notas encontradas em comum. Planilha 1 foi filtrada.")

        st.dataframe(df_filtrado)

        # Gerar Excel da planilha 1 filtrada (mantendo dados originais, substituindo ponto por vírgula na coluna "Saldo Pedido", e formatando no Excel)
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                colunas_desejadas = [
                    "Pedido",
                    "Cód.",
                    "Mercadoria",
                    "Cód..1",
                    "Filial",
                    "Cód..2",
                    "Fornecedor",
                    "Quantidade Pedido",
                    "Saldo Pedido",
                    "Nota Fiscal",
                    "Cobertura Atual"
                ]
                df_selecionado = df[colunas_desejadas].copy()  # Cópia das colunas desejadas

                # Substitui ponto por vírgula na coluna "Saldo Pedido" (mantendo como texto)
                df_selecionado["Saldo Pedido"] = df_selecionado["Saldo Pedido"].astype(str).str.replace('.', ',', regex=False)

                # Exporta para Excel
                df_selecionado.to_excel(writer, sheet_name='Notas Encontradas', index=False)

                workbook  = writer.book
                worksheet = writer.sheets['Notas Encontradas']

                # Formatar: centralizar + bordas
                cell_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1  # Bordas em todos os lados
                })

                # Aplicar formatação em todas as células usadas
                num_rows, num_cols = df_selecionado.shape
                worksheet.set_column(0, num_cols - 1, 20, cell_format)  # Ajusta largura e aplica centralização/borda

            return output.getvalue()

        excel_file = to_excel(df_filtrado)

        st.download_button(
            label="⬇️ Baixar Planilha 1 Filtrada",
            data=excel_file,
            file_name="planilha_filtrada_notas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("📂 Envie as duas planilhas acima para começar.")
