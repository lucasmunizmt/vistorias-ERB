import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
from fpdf import FPDF
from PIL import Image
import os

# Configuração da página
st.set_page_config(page_title="Gerador de Relatório ERB Oficial", layout="wide")

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 5px; background-color: #007bff; color: white; }
    </style>
    """, unsafe_allow_html=True)

st.title("🛡️ Sistema de Vistoria Técnica ERB")

# =========================
# INFORMAÇÕES INICIAIS
# =========================
with st.expander("📍 Dados do Processo e Localização", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        processo = st.text_input("Número do processo SEI")
        empresa = st.text_input("Empresa responsável")
        cnpj = st.text_input("CNPJ")
        regiao = st.text_input("Região Administrativa")
    with col2:
        endereco = st.text_input("Endereço Completo")
        licenca = st.text_input("Número da licença (Preencher quando for o caso de Licença de Implantação)")
        certificado = st.text_input("Número do Certificado de Cadastramento (Preencher quando for o caso de Certificado de Cadastramento)")

# =========================
# PERGUNTAS (VISTORIA)
# =========================
st.header("📝 Checklist de Vistoria")
q_col1, q_col2 = st.columns(2)

opcoes = ["Sim", "Não", "Não se aplica"]

with q_col1:
    q1 = st.radio("A área pública foi recuperada?", opcoes)
    q2 = st.radio("A altura da ERB está de acordo com o projeto?", opcoes)
    q3 = st.radio("A locação da ERB está de acordo com o projeto?", opcoes)

with q_col2:
    q4 = st.radio("O cercamento foi executado conforme o projeto?", opcoes)
    q5 = st.radio("Foi possível identificar as caixas de passagem?", opcoes)
    q6 = st.radio("A placa de advertência foi instalada?", opcoes)

# =========================
# AUTOS FISCAIS (MÚLTIPLOS)
# =========================
st.header("⚖️ Autos Fiscais")
if 'lista_autos' not in st.session_state:
    st.session_state.lista_autos = []

with st.container(border=True):
    c1, c2, c3 = st.columns([2, 2, 1])
    with c1:
        tipo_input = st.selectbox("Tipo de Auto", ["Auto de Notificação", "Auto de Infração", "Intimação Demolitória", "Auto de Interdição"])
    with c2:
        num_input = st.text_input("Número do Auto")
    with c3:
        st.write(" ")
        if st.button("➕ Adicionar"):
            if num_input:
                st.session_state.lista_autos.append({"tipo": tipo_input, "numero": num_input})
                st.rerun()

    if st.session_state.lista_autos:
        for i, auto in enumerate(st.session_state.lista_autos):
            cols = st.columns([4, 1])
            cols[0].write(f"✅ {auto['tipo']} nº {auto['numero']}")
            if cols[1].button("🗑️", key=f"del_{i}"):
                st.session_state.lista_autos.pop(i)
                st.rerun()

# =========================
# REGISTRO FOTOGRÁFICO (UPLOAD)
# =========================
st.header("📸 Registro Fotográfico")
fotos_upload = st.file_uploader("Selecione as fotos da vistoria", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

if fotos_upload:
    st.write(f"✅ {len(fotos_upload)} foto(s) carregada(s).")
    cols_fotos = st.columns(4)
    for idx, f in enumerate(fotos_upload):
        cols_fotos[idx % 4].image(f, use_container_width=True, caption=f"Foto {idx+1}")

# =========================
# FUNÇÕES DE LIMPEZA E EXPORTAÇÃO
# =========================

def limpar_caracteres_pdf(texto):
    subs = {
        '—': '-', '–': '-', '“': '"', '”': '"', '‘': "'", '’': "'", 'º': 'o.', 'ª': 'a.',
    }
    for orig, novo in subs.items():
        texto = texto.replace(orig, novo)
    return texto

def processar_imagem_para_relatorio(uploaded_file):
    """Redimensiona para economizar espaço no DOCX/PDF."""
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    img.thumbnail((800, 800))
    tmp_bio = BytesIO()
    img.save(tmp_bio, format="JPEG", quality=80)
    return tmp_bio

def gerar_docx(dados, irregularidades, autos, conclusao, fotos):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    def add_p(text, bold=False, indent=False):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = bold
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(0)
        if indent:
            p.paragraph_format.first_line_indent = Cm(1.25)
        return p

    # Cabeçalho
    add_p(f"PROCESSO SEI: {dados['processo']}", bold=True)
    add_p(f"EMPRESA: {dados['empresa']}", bold=True)
    add_p(f"CNPJ: {dados['cnpj']}", bold=True)
    add_p(f"REGIÃO ADMINISTRATIVA: {dados['regiao']}", bold=True)
    add_p(f"ENDEREÇO: {dados['endereco']}", bold=True)
    if dados['licenca'].strip(): add_p(f"LICENÇA: {dados['licenca']}", bold=True)
    if dados['certificado'].strip(): add_p(f"CERTIFICADO: {dados['certificado']}", bold=True)
    doc.add_paragraph()

    # Itens 1 a 4
    add_p("1. INTRODUÇÃO", bold=True)
    add_p("Trata-se de vistoria de Estação Rádio Base - ERB, nos termos da Lei Complementar n.o 971/2020, regulamentada pelo Decreto n.o 41446/2020. A vistoria consistiu basicamente em verificar a compatibilidade entre a infraestrutura implantada e o projeto aprovado, sem adentrar nos aspectos técnicos construtivos referentes à infraestrutura.", indent=True)

    add_p("2. RELATO DA VISTORIA", bold=True)
    if not irregularidades:
        add_p("Na vistoria, verificou-se a inexistência de irregularidades.", indent=True)
    else:
        add_p("Na vistoria realizada, foram constatadas as seguintes irregularidades:", indent=True)
        for irr in irregularidades: add_p(irr, indent=True)

    add_p("3. AUTOS FISCAIS", bold=True)
    if not autos:
        add_p("Não houve necessidade de emissão de autos fiscais.", indent=True)
    else:
        add_p("Foram emitidos os seguintes autos:", indent=True)
        for a in autos: add_p(f"- {a['tipo']} n.o {a['numero']}", indent=True)

    add_p("4. CONCLUSÃO", bold=True)
    add_p(conclusao, indent=True)

    # 5. REGISTRO FOTOGRÁFICO
    add_p("5. REGISTRO FOTOGRÁFICO", bold=True)
    if not fotos:
        add_p("Não foram anexados registros fotográficos.", indent=True)
    else:
        add_p("Segue abaixo o registro fotográfico colhido durante a vistoria:", indent=True)
        for idx, f in enumerate(fotos):
            img_stream = processar_imagem_para_relatorio(f)
            doc.add_picture(img_stream, width=Cm(14))
            parag_foto = doc.add_paragraph()
            parag_foto.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_foto = parag_foto.add_run(f"Foto {idx+1}")
            run_foto.italic = True

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def gerar_pdf(dados, irregularidades, autos, conclusao, fotos):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    
    def write_pdf_p(txt, bold=False, indent=False):
        txt = limpar_caracteres_pdf(txt)
        pdf.set_font("Arial", 'B' if bold else '', 11)
        prefix = "      " if indent else ""
        pdf.multi_cell(0, 8, txt=(prefix + txt).encode('latin-1', 'replace').decode('latin-1'), align='J')

    # Identificação
    campos = [("PROCESSO SEI", dados['processo']), ("EMPRESA", dados['empresa']), ("CNPJ", dados['cnpj']), ("REGIÃO ADMINISTRATIVA", dados['regiao']), ("ENDEREÇO", dados['endereco'])]
    if dados['licenca'].strip(): campos.append(("LICENÇA", dados['licenca']))
    if dados['certificado'].strip(): campos.append(("CERTIFICADO", dados['certificado']))
    for k, v in campos: write_pdf_p(f"{k}: {v}", bold=True)
    
    pdf.ln(5); write_pdf_p("1. INTRODUÇÃO", bold=True)
    write_pdf_p("Trata-se de vistoria de Estação Rádio Base - ERB, nos termos da Lei Complementar n.o 971/2020, regulamentada pelo Decreto n.o 41446/2020. A vistoria consistiu basicamente em verificar a compatibilidade entre a infraestrutura implantada e o projeto aprovado, sem adentrar nos aspectos técnicos construtivos referentes à infraestrutura.", indent=True)
    
    pdf.ln(5); write_pdf_p("2. RELATO DA VISTORIA", bold=True)
    if not irregularidades: write_pdf_p("Na vistoria, verificou-se a inexistência de irregularidades.", indent=True)
    else:
        write_pdf_p("Na vistoria realizada, foram constatadas as seguintes irregularidades:", indent=True)
        for irr in irregularidades: write_pdf_p(irr, indent=True)

    pdf.ln(5); write_pdf_p("3. AUTOS FISCAIS", bold=True)
    if not autos: write_pdf_p("Não houve necessidade de emissão de autos fiscais.", indent=True)
    else:
        write_pdf_p("Foram emitidos os seguintes autos:", indent=True)
        for a in autos: write_pdf_p(f"- {a['tipo']} n.o {a['numero']}", indent=True)

    pdf.ln(5); write_pdf_p("4. CONCLUSÃO", bold=True)
    write_pdf_p(conclusao, indent=True)

    pdf.ln(5); write_pdf_p("5. REGISTRO FOTOGRÁFICO", bold=True)
    if not fotos:
        write_pdf_p("Não foram anexados registros fotográficos.", indent=True)
    else:
        pdf.ln(2)
        for idx, f in enumerate(fotos):
            img_stream = processar_imagem_para_relatorio(f)
            temp_name = f"temp_{idx}.jpg"
            with open(temp_name, "wb") as t: t.write(img_stream.getvalue())
            # Adiciona nova página se a imagem não couber (simplificado)
            if pdf.get_y() > 200: pdf.add_page()
            pdf.image(temp_name, x=25, w=160)
            pdf.set_font("Arial", 'I', 9)
            pdf.cell(0, 8, f"Foto {idx+1}", ln=True, align='C')
            pdf.ln(5)
            os.remove(temp_name)

    return pdf.output(dest='S').encode('latin-1')

# =========================
# LÓGICA DE GERAÇÃO
# =========================
st.divider()
if st.button("🚀 GERAR RELATÓRIO FINAL"):
    irr_list = []
    if q1 == "Não": irr_list.append("- Falta recuperar a área pública degradada em decorrência da obra")
    if q2 == "Não": irr_list.append("- A altura da Estação Rádio Base está em desacordo com o projeto aprovado")
    if q3 == "Não": irr_list.append("- A locação da Estação Rádio Base está em desacordo com o projeto aprovado")
    if q4 == "Não": irr_list.append("- Não foi realizado o cercamento conforme o projeto aprovado")
    if q5 == "Não": irr_list.append("- Não foi possível identificar as caixas de passagem do projeto aprovado")
    if q6 == "Não": irr_list.append("- Falta a placa de advertência conforme o projeto aprovado")

    concl_text = ("Constatou-se que o equipamento encontra-se implantado em conformidade com o projeto aprovado, "
                  "não tendo, até o presente momento, a necessidade de adoção de medidas fiscais adicionais.") if not irr_list else \
                 ("Tendo em vista as irregularidades listadas no item \"Relato da Vistoria\", foram realizadas as ações "
                  "fiscais indicadas no item \"Autos Fiscais\", visando a regularização da Estação Rádio Base.")

    cabecalho_visual = f"PROCESSO SEI: {processo}\nEMPRESA: {empresa}\nCNPJ: {cnpj}\nREGIÃO ADMINISTRATIVA: {regiao}\nENDEREÇO: {endereco}"
    if licenca.strip(): cabecalho_visual += f"\nLICENÇA: {licenca}"
    if certificado.strip(): cabecalho_visual += f"\nCERTIFICADO: {certificado}"

    relato_visual = "Na vistoria, verificou-se a inexistência de irregularidades." if not irr_list else "Na vistoria realizada, foram constatadas as seguintes irregularidades:\n" + "\n".join(irr_list)
    autos_visual = "Não houve necessidade de emissão de autos fiscais." if not st.session_state.lista_autos else "Foram emitidos os seguintes autos:\n" + "\n".join([f"- {a['tipo']} n.o {a['numero']}" for a in st.session_state.lista_autos])
    
    fotos_visual = f"{len(fotos_upload)} foto(s) anexada(s)." if fotos_upload else "Nenhum registro fotográfico anexado."

    # PRÉVIA TEXTUAL
    relatorio_completo_visual = f"""{cabecalho_visual}

1. INTRODUÇÃO
Trata-se de vistoria de Estação Rádio Base - ERB, nos termos da Lei Complementar n.o 971/2020, regulamentada pelo Decreto n.o 41446/2020. A vistoria consistiu basicamente em verificar a compatibilidade entre a infraestrutura implantada e o projeto aprovado, sem adentrar nos aspectos técnicos construtivos referentes à infraestrutura.

2. RELATO DA VISTORIA
{relato_visual}

3. AUTOS FISCAIS
{autos_visual}

4. CONCLUSÃO
{concl_text}

5. REGISTRO FOTOGRÁFICO
{fotos_visual}
"""
    
    st.subheader("📋 Prévia do Relatório Completo")
    st.text_area("Texto pronto para cópia:", relatorio_completo_visual, height=450)

    dados_app = {"processo": processo, "empresa": empresa, "cnpj": cnpj, "regiao": regiao, "endereco": endereco, "licenca": licenca, "certificado": certificado}
    
    st.write("---")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        st.download_button("⬇️ Baixar DOCX (Word)", data=gerar_docx(dados_app, irr_list, st.session_state.lista_autos, concl_text, fotos_upload), file_name=f"Relatorio_{processo}.docx")
    with col_d2:
        st.download_button("⬇️ Baixar PDF (Oficial)", data=gerar_pdf(dados_app, irr_list, st.session_state.lista_autos, concl_text, fotos_upload), file_name=f"Relatorio_{processo}.pdf")