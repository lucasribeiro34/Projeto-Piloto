
import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import io
import zipfile

# ===== Funções auxiliares =====
def formatar_cpf(cpf):
    cpf = str(cpf).zfill(11)
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"

def formatar_telefone(telefone):
    telefone = str(telefone)
    return f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"

def gerar_pdf(dados):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    texto = f"""{dados['Nome Completo']}, {dados['Nacionalidade']}, {dados['Estado Civil']}, {dados['Profissão']}, telefone n° {dados['Telefone Formatado']},
com {dados['Tipo do Documento']} n° {dados['Número do Documento']}, inscrito sob o C.P.F. n° {dados['CPF Formatado']},
E-mail: {dados['E-mail']}, residente e domiciliado a {dados['Endereço Completo']}, {dados['Bairro']}, {dados['Cidade']}, {dados['Estado']}."""

    c.setFont("Times-Roman", 13)
    c.drawString(50, height - 100, texto)

    # Assinatura
    c.drawCentredString(width / 2, 150, "__________________________________________________")
    c.drawCentredString(width / 2, 130, dados['Nome Completo'])
    c.drawCentredString(width / 2, 115, dados['CPF Formatado'])

    c.save()
    buffer.seek(0)
    return buffer

# ===== Streamlit App =====
st.title("Gerador de Procurações em PDF")

arquivo = st.file_uploader("Faça upload da planilha Excel (.xlsx)", type=["xlsx"])

if arquivo is not None:
    df = pd.read_excel(arquivo)

    st.success("Arquivo carregado com sucesso!")
    gerar_btn = st.button("Gerar PDFs")

    if gerar_btn:
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for index, dados in df.iterrows():
                dados['CPF Formatado'] = formatar_cpf(dados['CPF'])
                dados['Telefone Formatado'] = formatar_telefone(dados['Telefone para contato'])

                pdf_buffer = gerar_pdf(dados)
                nome_pdf = f"Procuração - {dados['Nome Completo']}.pdf"
                zip_file.writestr(nome_pdf, pdf_buffer.read())

        st.success("PDFs gerados com sucesso!")

        zip_buffer.seek(0)
        st.download_button(
            label="📁 Baixar arquivos ZIP com PDFs",
            data=zip_buffer,
            file_name="Procuracoes.zip",
            mime="application/zip"
        )
