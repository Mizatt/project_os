"""
RPM Eletrodiesel — Gerador de Ordem de Serviço
Versão Web (Streamlit)

Instalação:
    pip install streamlit python-docx

Execução local:
    streamlit run project_os.py

Deploy:
    1. Suba este arquivo + LOGO_Horizontal_cor.png para o GitHub
    2. Acesse share.streamlit.io e conecte o repositório
"""

import base64
import io
import os
import re
from datetime import datetime

import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ─────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────
LOGO_FILENAME = "LOGO_Horizontal_cor.png"
EMPRESA_ENDE  = "RODOVIA RAPOSO TAVARES KM 104,750 - 245 - IPANEMA DO MEIO - SOROCABA - SP"
COR_ESCURA    = "17233F"
COR_MEDIA     = "4A90E2"
COR_LINHA_PAR = "EBEBEB"


# ─────────────────────────────────────────────
# Helpers docx
# ─────────────────────────────────────────────

def _bg(cell, cor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def _bordas(cell, cor="AAAAAA", esp="4",
            top=True, bottom=True, left=True, right=True):
    tcPr    = cell._tc.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for lado, ativo in (("top", top), ("bottom", bottom),
                        ("left", left), ("right", right)):
        el = OxmlElement("w:" + lado)
        if ativo:
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    esp)
            el.set(qn("w:color"), cor)
        else:
            el.set(qn("w:val"), "nil")
        borders.append(el)
    tcPr.append(borders)


def _texto_celula(cell, texto, bold=False, size=10,
                  align=WD_ALIGN_PARAGRAPH.LEFT, cor_hex=None):
    p   = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(texto)
    run.bold      = bold
    run.font.size = Pt(size)
    if cor_hex:
        run.font.color.rgb = RGBColor(*bytes.fromhex(cor_hex))
    return run


def _par(doc, texto="", bold=False, size=11,
         align=WD_ALIGN_PARAGRAPH.LEFT, depois=0):
    p   = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = Pt(depois)
    run = p.add_run(texto)
    run.bold      = bold
    run.font.size = Pt(size)
    return p


# ─────────────────────────────────────────────
# Geração do .docx em memória (sem salvar em disco)
# ─────────────────────────────────────────────

def gerar_os_bytes(dados: dict) -> bytes:
    """Gera o .docx e retorna os bytes — ideal para download via web."""
    logo_path = LOGO_FILENAME if os.path.isfile(LOGO_FILENAME) else None

    doc = Document()
    for sec in doc.sections:
        sec.top_margin    = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin   = Cm(2.0)
        sec.right_margin  = Cm(2.0)

    # ── Cabeçalho ─────────────────────────────
    tbl_hdr = doc.add_table(rows=2, cols=2)
    tbl_hdr.style   = "Table Grid"
    tbl_hdr.autofit = False
    tbl_hdr.columns[0].width = Cm(8)
    tbl_hdr.columns[1].width = Cm(9)

    c_logo = tbl_hdr.cell(0, 0)
    _bg(c_logo, "FFFFFF")
    _bordas(c_logo, cor="CCCCCC")
    c_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_logo = c_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if logo_path:
        p_logo.add_run().add_picture(logo_path, width=Cm(7))
    else:
        r = p_logo.add_run("RPM ELETRODIESEL")
        r.bold = True
        r.font.size = Pt(14)
        r.font.color.rgb = RGBColor(0x17, 0x23, 0x3F)

    c_tit = tbl_hdr.cell(0, 1)
    _bg(c_tit, "FFFFFF")
    _bordas(c_tit, cor="CCCCCC")
    c_tit.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p_tit = c_tit.paragraphs[0]
    p_tit.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_tit.add_run("ORDEM DE SERVICO")
    r1.bold = True
    r1.font.size = Pt(15)
    r1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    c_end = tbl_hdr.cell(1, 0)
    c_end.merge(tbl_hdr.cell(1, 1))
    _bg(c_end, COR_MEDIA)
    _bordas(c_end, cor="FFFFFF")
    _texto_celula(c_end, EMPRESA_ENDE, size=9,
                  align=WD_ALIGN_PARAGRAPH.CENTER, cor_hex="FFFFFF")

    doc.add_paragraph()

    # ── OS / Data ─────────────────────────────
    tbl_os = doc.add_table(rows=1, cols=2)
    tbl_os.style   = "Table Grid"
    tbl_os.autofit = False
    tbl_os.columns[0].width = Cm(8.5)
    tbl_os.columns[1].width = Cm(8.5)
    for cel, txt in ((tbl_os.cell(0, 0), "O.S.: " + dados["os_numero"]),
                     (tbl_os.cell(0, 1), "DATA: " + dados["data_os"])):
        _bg(cel, COR_ESCURA)
        _bordas(cel, cor="FFFFFF")
        aln = WD_ALIGN_PARAGRAPH.RIGHT if "DATA" in txt else WD_ALIGN_PARAGRAPH.LEFT
        _texto_celula(cel, txt, bold=True, size=11, align=aln, cor_hex="FFFFFF")

    doc.add_paragraph()

    # ── Cliente / Veículo ─────────────────────
    campos_cli = [
        ("CLIENTE",  dados["cliente"]),
        ("CIDADE",   dados["cidade"]),
        ("MOTOR",    dados["motor"]),
        ("VEICULO",  dados["veiculo"]),
        ("PLACA",    dados["placa"]),
    ]
    tbl_cli = doc.add_table(rows=len(campos_cli), cols=2)
    tbl_cli.style   = "Table Grid"
    tbl_cli.autofit = False
    tbl_cli.columns[0].width = Cm(3.5)
    tbl_cli.columns[1].width = Cm(13.5)
    for i, (label, valor) in enumerate(campos_cli):
        cl = tbl_cli.cell(i, 0)
        cv = tbl_cli.cell(i, 1)
        _bg(cl, COR_MEDIA)
        _bordas(cl, cor="CCCCCC")
        _bordas(cv, cor="CCCCCC")
        _texto_celula(cl, label, bold=True, size=9, cor_hex="FFFFFF")
        _texto_celula(cv, valor, size=10)

    doc.add_paragraph()

    # ── Serviços ──────────────────────────────
    p = _par(doc, "SERVICOS REALIZADOS", bold=True, size=10, depois=2)
    p.runs[0].font.color.rgb = RGBColor(*bytes.fromhex(COR_ESCURA))
    tbl_svc = doc.add_table(rows=len(dados["servicos"]), cols=1)
    tbl_svc.style   = "Table Grid"
    tbl_svc.autofit = False
    tbl_svc.columns[0].width = Cm(17)
    for i, svc in enumerate(dados["servicos"]):
        c = tbl_svc.cell(i, 0)
        _bg(c, COR_LINHA_PAR if i % 2 == 0 else "FFFFFF")
        _bordas(c, cor="CCCCCC")
        _texto_celula(c, "{:02d}. {}".format(i + 1, svc), size=10)

    doc.add_paragraph()

    # ── Peças ─────────────────────────────────
    p = _par(doc, "PECAS UTILIZADAS", bold=True, size=10, depois=2)
    p.runs[0].font.color.rgb = RGBColor(*bytes.fromhex(COR_ESCURA))
    tbl_pcs = doc.add_table(rows=1 + len(dados["pecas"]), cols=3)
    tbl_pcs.style   = "Table Grid"
    tbl_pcs.autofit = False
    tbl_pcs.columns[0].width = Cm(3)
    tbl_pcs.columns[1].width = Cm(5)
    tbl_pcs.columns[2].width = Cm(9)
    for j, titulo in enumerate(("QTDE.", "CODIGO", "DESCRICAO")):
        ch = tbl_pcs.cell(0, j)
        _bg(ch, COR_MEDIA)
        _bordas(ch, cor="FFFFFF")
        _texto_celula(ch, titulo, bold=True, size=9,
                      align=WD_ALIGN_PARAGRAPH.CENTER, cor_hex="FFFFFF")
    for i, peca in enumerate(dados["pecas"]):
        bg = COR_LINHA_PAR if i % 2 == 0 else "FFFFFF"
        for j, (val, aln) in enumerate([
            (str(peca["qtd"]),   WD_ALIGN_PARAGRAPH.CENTER),
            (peca["codigo"],     WD_ALIGN_PARAGRAPH.LEFT),
            (peca["descricao"],  WD_ALIGN_PARAGRAPH.LEFT),
        ]):
            c = tbl_pcs.cell(i + 1, j)
            _bg(c, bg)
            _bordas(c, cor="CCCCCC")
            _texto_celula(c, val, size=10, align=aln)

    doc.add_paragraph()

    # ── Observações ───────────────────────────
    p = _par(doc, "OBSERVACOES", bold=True, size=10, depois=2)
    p.runs[0].font.color.rgb = RGBColor(*bytes.fromhex(COR_ESCURA))
    tbl_obs = doc.add_table(rows=1, cols=1)
    tbl_obs.style   = "Table Grid"
    tbl_obs.autofit = False
    tbl_obs.columns[0].width = Cm(17)
    c_obs = tbl_obs.cell(0, 0)
    _bg(c_obs, "FFFFFF")
    _bordas(c_obs, cor="AAAAAA")
    _texto_celula(c_obs, dados["obs"] if dados["obs"] else "-", size=10)

    doc.add_paragraph()

    # ── Colaborador ───────────────────────────
    p = _par(doc, "COLABORADOR QUE EXECUTOU O SERVICO", bold=True, size=10, depois=2)
    p.runs[0].font.color.rgb = RGBColor(*bytes.fromhex(COR_ESCURA))
    tbl_col = doc.add_table(rows=1, cols=1)
    tbl_col.style   = "Table Grid"
    tbl_col.autofit = False
    tbl_col.columns[0].width = Cm(17)
    c_col = tbl_col.cell(0, 0)
    _bg(c_col, "FFFFFF")
    _bordas(c_col, cor="AAAAAA")
    _texto_celula(c_col, dados["colaborador"], size=10)

    doc.add_paragraph()

    # ── Assinaturas ───────────────────────────
    tbl_ass = doc.add_table(rows=2, cols=2)
    tbl_ass.style   = "Table Grid"
    tbl_ass.autofit = False
    tbl_ass.columns[0].width = Cm(8.5)
    tbl_ass.columns[1].width = Cm(8.5)
    for col, label in enumerate(("Assinatura do Tecnico", "Assinatura do Cliente")):
        c_lbl = tbl_ass.cell(0, col)
        c_sig = tbl_ass.cell(1, col)
        _bg(c_lbl, COR_MEDIA)
        _bordas(c_lbl, cor="FFFFFF")
        _texto_celula(c_lbl, label, bold=True, size=9,
                      align=WD_ALIGN_PARAGRAPH.CENTER, cor_hex="FFFFFF")
        _bordas(c_sig, cor="AAAAAA")
        _texto_celula(c_sig, "", size=16)

    # ── Retorna bytes (sem salvar em disco) ───
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Validações
# ─────────────────────────────────────────────

def _validar_data(val: str) -> str | None:
    """Retorna mensagem de erro ou None se válida."""
    if not re.match(r"^\d{2}/\d{2}/\d{4}$", val):
        return "Use o formato DD/MM/AAAA (ex.: 25/02/2026)."
    try:
        d, m, a = val.split("/")
        datetime(int(a), int(m), int(d))
    except ValueError:
        return "Data inexistente. Verifique dia, mês e ano."
    return None


# ─────────────────────────────────────────────
# Interface Streamlit
# ─────────────────────────────────────────────

st.set_page_config(
    page_title="RPM Eletrodiesel — Gerador de O.S.",
    page_icon="🔧",
    layout="centered",
)

# Cabeçalho visual com logo
def _logo_base64():
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_FILENAME)
    if os.path.isfile(caminho):
        with open(caminho, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return None

_logo_b64 = _logo_base64()
_logo_tag = (
    f'<img src="data:image/png;base64,{_logo_b64}" style="height:72px;object-fit:contain;">' 
    if _logo_b64 else
    '<span style="color:#8AABCF;font-size:13px">RPM ELETRODIESEL</span>'
)

st.markdown(f"""
<div style="background:#17233F;padding:16px 24px;border-radius:8px;
            margin-bottom:24px;display:flex;align-items:center;
            justify-content:space-between;gap:16px;">
  <div>
    <span style="color:white;font-size:22px;font-weight:700">RPM ELETRODIESEL</span><br>
    <span style="color:#8AABCF;font-size:13px">Gerador de Ordem de Serviço</span>
  </div>
  <div style="background:white;border-radius:6px;padding:8px 16px;">
    {_logo_tag}
  </div>
</div>
""", unsafe_allow_html=True)

# ── Seção: OS ─────────────────────────────────
st.subheader("📋 Dados da Ordem de Serviço")
col1, col2 = st.columns(2)
with col1:
    os_numero = st.text_input("Número da O.S. *", placeholder="ex.: 6897")
with col2:
    data_os = st.text_input(
        "Data *",
        placeholder="DD/MM/AAAA",
        value=datetime.today().strftime("%d/%m/%Y"),
    )

# ── Seção: Cliente ────────────────────────────
st.subheader("👤 Dados do Cliente / Veículo")
col1, col2 = st.columns(2)
with col1:
    cliente = st.text_input("Cliente *", placeholder="Nome do cliente")
with col2:
    cidade  = st.text_input("Cidade *",  placeholder="ex.: Sorocaba - SP")

col3, col4 = st.columns(2)
with col3:
    motor   = st.text_input("Motor *",   placeholder="ex.: Cummins ISF 04 Cil.")
with col4:
    veiculo = st.text_input("Veículo *", placeholder="ex.: VW 9-160")

placa = st.text_input("Placa", placeholder="ex.: ABC1D23 (máx. 8 caracteres)", max_chars=8)

# ── Seção: Serviços ───────────────────────────
st.subheader("🔧 Serviços Realizados")

if "num_servicos" not in st.session_state:
    st.session_state.num_servicos = 1

servicos_vals = []
for i in range(st.session_state.num_servicos):
    obrig = " *" if i == 0 else ""
    servicos_vals.append(
        st.text_input(f"Serviço {i+1:02d}{obrig}", key=f"svc_{i}",
                      placeholder="Descreva o serviço")
    )

col_add, col_rem = st.columns([1, 1])
with col_add:
    if st.button("+ Adicionar serviço"):
        st.session_state.num_servicos += 1
        st.rerun()
with col_rem:
    if st.session_state.num_servicos > 1:
        if st.button("− Remover último"):
            st.session_state.num_servicos -= 1
            st.rerun()

# ── Seção: Peças ──────────────────────────────
st.subheader("🔩 Peças Utilizadas")

if "num_pecas" not in st.session_state:
    st.session_state.num_pecas = 1

pecas_vals = []
for i in range(st.session_state.num_pecas):
    obrig = " *" if i == 0 else ""
    c1, c2, c3 = st.columns([3, 1, 4])
    with c1:
        cod = st.text_input(f"Código{obrig}", key=f"peca_cod_{i}",
                            placeholder="ex.: CDP-125014")
    with c2:
        qtd = st.text_input(f"Qtde.{obrig}", key=f"peca_qtd_{i}",
                            placeholder="1")
    with c3:
        desc = st.text_input(f"Descrição{obrig}", key=f"peca_desc_{i}",
                             placeholder="ex.: Sensor de pressão")
    pecas_vals.append((cod, qtd, desc))

col_add2, col_rem2 = st.columns([1, 1])
with col_add2:
    if st.button("+ Adicionar peça"):
        st.session_state.num_pecas += 1
        st.rerun()
with col_rem2:
    if st.session_state.num_pecas > 1:
        if st.button("− Remover última"):
            st.session_state.num_pecas -= 1
            st.rerun()

# ── Seção: Observações / Colaborador ──────────
st.subheader("📝 Observações e Execução")
obs         = st.text_area("Observações", placeholder="Informações adicionais sobre o serviço...")
colaborador = st.text_input("Colaborador que executou o serviço *",
                             placeholder="Nome do técnico responsável")

st.markdown("---")
st.caption("\\* Campos obrigatórios")

# ── Botão Gerar ───────────────────────────────
if st.button("⚙️  GERAR ORDEM DE SERVIÇO", type="primary", use_container_width=True):
    erros = []

    # Validações
    if not os_numero.strip():
        erros.append("Número da O.S.")

    erro_data = _validar_data(data_os.strip())
    if erro_data:
        erros.append("Data — " + erro_data)

    if not cliente.strip():  erros.append("Cliente")
    if not cidade.strip():   erros.append("Cidade")
    if not motor.strip():    erros.append("Motor")
    if not veiculo.strip():  erros.append("Veículo")

    if placa.strip() and len(placa.strip()) > 8:
        erros.append("Placa (máximo 8 caracteres)")

    # Serviços
    servicos = []
    for i, s in enumerate(servicos_vals):
        sv = s.strip()
        if i == 0 and not sv:
            erros.append("Serviço 01 (obrigatório)")
        elif sv:
            servicos.append(sv)

    # Peças
    pecas = []
    for i, (cod, qtd, desc) in enumerate(pecas_vals):
        c, q, d = cod.strip(), qtd.strip(), desc.strip()
        if i == 0 and not c:
            erros.append("Código da 1ª peça (obrigatório)")
            continue
        if not c:
            continue
        if not q or not q.isdigit() or int(q) < 1:
            erros.append(f"Quantidade da peça {i+1} (número inteiro > 0)")
            continue
        if not d:
            erros.append(f"Descrição da peça {i+1} (obrigatória)")
            continue
        pecas.append({"codigo": c, "qtd": int(q), "descricao": d})

    if not colaborador.strip():
        erros.append("Colaborador que executou o serviço")

    if erros:
        st.error("**Corrija os seguintes campos antes de gerar:**\n\n" +
                 "\n".join(f"- {e}" for e in erros))
    else:
        dados = {
            "os_numero":   os_numero.strip(),
            "data_os":     data_os.strip(),
            "cliente":     cliente.strip(),
            "cidade":      cidade.strip(),
            "motor":       motor.strip(),
            "veiculo":     veiculo.strip(),
            "placa":       placa.strip().upper(),
            "servicos":    servicos,
            "pecas":       pecas,
            "obs":         obs.strip(),
            "colaborador": colaborador.strip(),
        }

        try:
            docx_bytes = gerar_os_bytes(dados)

            cliente_safe = re.sub(r'[\\/:*?"<>|]', "", dados["cliente"]).strip()
            nome_arquivo = "OS_{}_{}_{}.docx".format(
                dados["os_numero"],
                cliente_safe,
                dados["data_os"].replace("/", ""),
            )

            st.success("✅ Ordem de Serviço gerada com sucesso!")
            st.download_button(
                label="⬇️  Baixar " + nome_arquivo,
                data=docx_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        except Exception as e:
            st.error("Erro ao gerar o documento: " + str(e))