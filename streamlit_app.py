
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

    texto_inicial = f"""{dados['Nome Completo']}, {dados['Nacionalidade']}, {dados['Estado Civil']}, {dados['Profissão']}, telefone n° {dados['Telefone Formatado']},\n
com {dados['Tipo do Documento']} n° {dados['Número do Documento']}, inscrito sob o C.P.F. n° {dados['CPF Formatado']}, E-mail: {dados['E-mail']},\n
residente e domiciliado a {dados['Endereço Completo']}, {dados['Bairro']}, {dados['Cidade']}, {dados['Estado']}.

"""

    poderes = "Confiro poderes para que me represente perante\n[órgãos públicos e privados/situação específica],\npodendo praticar os seguintes atos: [detalhar poderes, como assinar documentos, retirar certidões, etc.].\n
"

    validade = "Esta procuração é válida até [data de validade] ou enquanto durar [condição específica].\n
"

    local_data = "Local e Data: [Cidade], [dia] de [mês] de [ano].\n\nAssinatura [Assinatura do Outorgante] [Nome do Outorgante]\n\n"

    assinatura = f"""__________________________________________________\n{dados['Nome Completo']}\n{dados['CPF Formatado']}"""

    texto_completo = texto_inicial + poderes + validade + local_data + assinatura

    c.setFont("Times-Roman", 12)

    # Inserir texto com quebra de linha
    y = height - 80
    for linha in texto_completo.split("\n"):
        c.drawString(50, y, linha)
        y -= 18  # espaçamento entre linhas

    c.save()
    buffer.seek(0)
    return buffer

# ===== Streamlit App =====
st.title("Gerador de Procurações em PDF - Layout Profissional")

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
            file_name="Procuracoes_Formatadas.zip",
            mime="application/zip"
        )
