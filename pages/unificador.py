import os
import sys
import io
import streamlit as st
from pypdf import PdfWriter, PdfReader
from datetime import datetime

ROOT_DIR = os.path.dirname(os.path.dirname(__file__))
sys.path.insert(0, os.path.join(ROOT_DIR, "config"))
sys.path.insert(0, os.path.join(ROOT_DIR, "services"))

st.set_page_config(
    page_title="Unificador de PDFs — FastPlac",
    page_icon=os.path.join(ROOT_DIR, "images", "favicon-16x16.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Sora', sans-serif;
}

[data-testid="stSidebar"] {
    background-color: #242480 !important;
    min-width: 280px !important;
    max-width: 320px !important;
}
[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.2) !important;
}

[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {
    background-color: rgba(255,255,255,0.12) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    border-radius: 6px !important;
    font-family: 'Sora', sans-serif !important;
    font-size: 0.85rem !important;
    min-height: 38px !important;
    width: 100% !important;
    padding: 0.4rem 1rem !important;
    transition: background 0.2s;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-primary"]:hover {
    background-color: rgba(255,255,255,0.22) !important;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
    background-color: transparent !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    color: #FFFFFF !important;
    border-radius: 4px !important;
    font-size: 0.72rem !important;
    min-height: 26px !important;
    max-height: 26px !important;
    padding: 0 8px !important;
    margin-top: 2px !important;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
    background-color: rgba(255, 80, 80, 0.35) !important;
}

[data-testid="stMain"] [data-testid="stBaseButton-primary"],
.stDownloadButton > button {
    background-color: #242480 !important;
    color: #FFFFFF !important;
    border: none !important;
    border-radius: 6px !important;
    font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
    padding: 0.5rem 1.2rem !important;
    transition: background 0.2s, transform 0.1s;
}
[data-testid="stMain"] [data-testid="stBaseButton-primary"]:hover,
.stDownloadButton > button:hover {
    background-color: #1a1a60 !important;
    transform: translateY(-1px);
}

h1, h2, h3 {
    font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important;
    color: #1a1a1a !important;
}
.section-title {
    font-family: 'Sora', sans-serif;
    font-size: 1.3rem;
    font-weight: 600;
    color: #1a1a1a;
    border-left: 4px solid #242480;
    padding-left: 12px;
    margin-bottom: 1rem;
    margin-top: 0.5rem;
}
.sidebar-logo {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 4px 0 12px 0;
}
.sidebar-logo span {
    font-size: 1.4rem;
    font-weight: 700;
    letter-spacing: 1px;
    color: #FFFFFF !important;
}
.badge-count {
    display: inline-block;
    background: rgba(255,255,255,0.2);
    color: #fff;
    border-radius: 20px;
    padding: 1px 10px;
    font-size: 0.8rem;
    font-weight: 600;
    margin-left: 6px;
}
.pdf-card {
    background: #f8f9ff;
    border: 1px solid #e0e4f5;
    border-radius: 8px;
    padding: 10px 14px;
    margin-bottom: 6px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.pdf-card-nome {
    font-size: 0.88rem;
    font-weight: 600;
    color: #242480;
}
.pdf-card-info {
    font-size: 0.78rem;
    color: #888;
}

.stProgress > div > div > div {
    background-color: #242480 !important;
}
hr { border-color: #e8eaf0 !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
if "pdfs_upload"   not in st.session_state: st.session_state.pdfs_upload   = []
if "pdf_unificado" not in st.session_state: st.session_state.pdf_unificado = None

# ── Sidebar ──
import base64
with st.sidebar:
    logo_path = os.path.join(ROOT_DIR, "images", "logo_aguia1.png")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
        logo_html = f'<img src="data:image/png;base64,{logo_b64}" width="50" />'
    else:
        logo_html = ""

    st.markdown(f"""
    <div class="sidebar-logo">
        {logo_html}
        <span>FastPlac</span>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    total = len(st.session_state.pdfs_upload)
    st.markdown(
        f'<div style="font-weight:600; font-size:0.9rem; margin-bottom:10px;">'
        f'PDFs carregados <span class="badge-count">{total}</span></div>',
        unsafe_allow_html=True
    )

    if st.session_state.pdfs_upload:
        for i, pdf in enumerate(st.session_state.pdfs_upload):
            col_info, col_up, col_down, col_del = st.columns([6, 1, 1, 1])
            with col_info:
                st.markdown(
                    f'<div style="font-size:0.8rem; font-weight:600; margin-top:4px;">'
                    f'{i+1}. {pdf["nome"]}</div>'
                    f'<div style="font-size:0.72rem; opacity:0.7;">{pdf["paginas"]} pág.</div>',
                    unsafe_allow_html=True
                )
            with col_up:
                if i > 0 and st.button("↑", key=f"up_{i}", help="Mover para cima"):
                    st.session_state.pdfs_upload[i], st.session_state.pdfs_upload[i-1] = \
                        st.session_state.pdfs_upload[i-1], st.session_state.pdfs_upload[i]
                    st.session_state.pdf_unificado = None
                    st.rerun()
            with col_down:
                if i < total - 1 and st.button("↓", key=f"down_{i}", help="Mover para baixo"):
                    st.session_state.pdfs_upload[i], st.session_state.pdfs_upload[i+1] = \
                        st.session_state.pdfs_upload[i+1], st.session_state.pdfs_upload[i]
                    st.session_state.pdf_unificado = None
                    st.rerun()
            with col_del:
                if st.button("✕", key=f"del_{i}", help="Remover"):
                    st.session_state.pdfs_upload.pop(i)
                    st.session_state.pdf_unificado = None
                    st.rerun()

        st.divider()
        if st.button("Limpar todos", use_container_width=True, type="primary"):
            st.session_state.pdfs_upload   = []
            st.session_state.pdf_unificado = None
            st.rerun()
    else:
        st.markdown(
            '<div style="font-size:0.82rem; opacity:0.7; padding: 6px 0;">'
            'Nenhum PDF carregado ainda.</div>',
            unsafe_allow_html=True
        )

# ── Conteúdo principal ──
st.markdown('<div class="section-title">Unificador de PDFs</div>', unsafe_allow_html=True)
st.caption("Faça upload dos relatórios, organize a ordem e baixe o PDF unificado.")

# Upload
arquivos = st.file_uploader(
    "Selecione os PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if arquivos:
    nomes_existentes = {p["nome"] for p in st.session_state.pdfs_upload}
    adicionados = 0
    for arq in arquivos:
        if arq.name not in nomes_existentes:
            conteudo = arq.read()
            reader   = PdfReader(io.BytesIO(conteudo))
            st.session_state.pdfs_upload.append({
                "nome":     arq.name,
                "bytes":    conteudo,
                "paginas":  len(reader.pages),
            })
            adicionados += 1

    if adicionados:
        st.session_state.pdf_unificado = None
        st.rerun()

st.divider()

# ── Unificar ──
st.markdown('<div class="section-title">Unificar</div>', unsafe_allow_html=True)

if not st.session_state.pdfs_upload:
    st.info("Faça upload de pelo menos dois PDFs para unificar.")

else:
    total_pags = sum(p["paginas"] for p in st.session_state.pdfs_upload)
    st.caption(
        f"{len(st.session_state.pdfs_upload)} arquivo(s)  ·  "
        f"{total_pags} página(s) no total  ·  "
        f"Ordem definida na barra lateral"
    )

    nome_arquivo = st.text_input(
        "Nome do arquivo unificado",
        value=f"Relatórios Unificados — {datetime.now().strftime('%d-%m-%Y')}",
    )

    col_btn, _ = st.columns([1, 7])
    with col_btn:
        unificar = st.button("Unificar PDFs", type="primary", use_container_width=True)

    if unificar:
        with st.spinner("Unificando..."):
            writer = PdfWriter()
            for pdf in st.session_state.pdfs_upload:
                reader = PdfReader(io.BytesIO(pdf["bytes"]))
                for page in reader.pages:
                    writer.add_page(page)
            out = io.BytesIO()
            writer.write(out)
            st.session_state.pdf_unificado = out.getvalue()
        st.success("PDFs unificados com sucesso!")

    if st.session_state.pdf_unificado:
        _, col_dl, _ = st.columns([4, 2, 4])
        with col_dl:
            st.download_button(
                label="Baixar PDF Unificado",
                data=st.session_state.pdf_unificado,
                file_name=f"{nome_arquivo}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
            )

# ── Rodapé ──
ano = datetime.now().year
st.markdown(f"""
<div style="
    border-top: 1px solid #e0e4f5;
    padding: 10px 24px;
    font-family: 'Sora', sans-serif;
    font-size: 0.78rem;
    color: #888;
    text-align: center;
    margin-top: 2rem;
">
    {ano} © FastPlac by <a href="https://aguiasistemas.com.br/" target="_blank"
    style="color:#242480; font-weight:700; text-decoration:none;">Águia Sistemas</a>
    &nbsp;·&nbsp; Version 1.0.0
</div>
""", unsafe_allow_html=True)