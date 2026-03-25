import os
import sys
import json
import base64
import streamlit as st
from datetime import datetime

# ── Paths das pastas ──
ROOT_DIR    = os.path.dirname(__file__)
SECRETS_DIR = os.path.join(ROOT_DIR, "secrets")

sys.path.insert(0, os.path.join(ROOT_DIR, "config"))
sys.path.insert(0, os.path.join(ROOT_DIR, "services"))

from field_config import TEMPLATE_IDS, FOLDER_ID, CAMPOS_COMUNS, CAMPOS_ESPECIFICOS
from google_api import gerar_pdf_consolidado
from report import gerar_relatorio

st.set_page_config(
    page_title="FastPlac",
    page_icon=os.path.join(ROOT_DIR, "images", "favicon-16x16.png"),
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;500;600;700&display=swap');

/* ── Forçar modo light ── */
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
    color-scheme: light !important;
    background-color: #ffffff !important;
    color: #1a1a1a !important;
}
[data-theme="dark"] { color-scheme: light !important; }

html, body, [class*="css"] {
    font-family: 'Sora', sans-serif;
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background-color: #242480 !important;
}
[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.2) !important;
}
[data-testid="stSidebar"] [data-testid="stAlert"] {
    background-color: rgba(255,255,255,0.1) !important;
    border: none !important;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
    background-color: transparent !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    color: #FFFFFF !important;
    border-radius: 4px !important;
    font-size: 0.65rem !important;
    min-height: 24px !important;
    max-height: 24px !important;
    min-width: 24px !important;
    max-width: 24px !important;
    padding: 0 !important;
    margin-top: 4px !important;
    line-height: 1 !important;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
    background-color: rgba(255, 80, 80, 0.35) !important;
    border-color: rgba(255, 100, 100, 0.5) !important;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-primary"] {
    background-color: rgba(255,255,255,0.12) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(255,255,255,0.25) !important;
    border-radius: 6px !important;
    font-family: 'Sora', sans-serif !important;
    font-size: 0.85rem !important;
    min-height: 38px !important;
    max-height: none !important;
    width: 100% !important;
    padding: 0.4rem 1rem !important;
    transition: background 0.2s;
}
[data-testid="stSidebar"] [data-testid="stBaseButton-primary"]:hover {
    background-color: rgba(255,255,255,0.22) !important;
}

/* ── Botões principais ── */
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

/* ── Tipografia ── */
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
    margin-top: 1.5rem;
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
.drive-link {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    background: #242480;
    border-radius: 6px;
    padding: 10px 14px;
    font-size: 0.9rem;
    font-weight: 600;
    color: #FFFFFF !important;
    text-decoration: none !important;
    width: 100%;
    box-sizing: border-box;
}
.drive-link:hover { background: #1a1a60; }

/* ── Inputs ── */
.stTextInput input, .stNumberInput input {
    border-radius: 6px !important;
    font-family: 'Sora', sans-serif !important;
    border-color: #d0d5e8 !important;
    background-color: #ffffff !important;
    color: #1a1a1a !important;
}
.stTextInput input:focus, .stNumberInput input:focus {
    border-color: #242480 !important;
    box-shadow: 0 0 0 2px rgba(36,36,128,0.1) !important;
}
.stSelectbox > div > div {
    border-radius: 6px !important;
    font-family: 'Sora', sans-serif !important;
    background-color: #ffffff !important;
    color: #1a1a1a !important;
}

/* ── Responsividade ── */
[data-testid="stMain"] .block-container {
    padding: 1.5rem 1rem !important;
    max-width: 100% !important;
}

/* Botões sempre largura total no mobile */
@media (max-width: 768px) {
    [data-testid="stMain"] .block-container {
        padding: 1rem 0.5rem !important;
    }
    .section-title {
        font-size: 1.1rem;
    }
    /* Faz colunas stackarem verticalmente */
    [data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
    }
    [data-testid="stHorizontalBlock"] > [data-testid="stColumn"] {
        min-width: 100% !important;
        width: 100% !important;
        flex: 1 1 100% !important;
    }
    /* Botões ocupam largura total */
    [data-testid="stMain"] [data-testid="stBaseButton-primary"],
    .stDownloadButton > button {
        width: 100% !important;
    }
    .drive-link {
        width: 100% !important;
    }
}

/* ── Tira a barra colorida do Streamlit ── */            
div[data-testid="stDecoration"] { display: none !important; }
header[data-testid="stHeader"]  { background: none !important; }

/* ── Misc ── */
.stProgress > div > div > div {
    background-color: #242480 !important;
}
hr { border-color: #e8eaf0 !important; }
</style>
""", unsafe_allow_html=True)

# ── Session state ──
if "placas"     not in st.session_state: st.session_state.placas     = []
if "pdf_pronto" not in st.session_state: st.session_state.pdf_pronto = None
if "pdf_nome"   not in st.session_state: st.session_state.pdf_nome   = ""
if "link_drive" not in st.session_state: st.session_state.link_drive = ""
if "relatorio"  not in st.session_state: st.session_state.relatorio  = None

if "credentials" not in st.session_state:
    try:
        if "google" in st.secrets:
            cred = st.secrets["google"]["credentials"]
            if isinstance(cred, str):
                st.session_state.credentials = json.loads(cred)
            else:
                st.session_state.credentials = dict(cred)
        else:
            raise KeyError
    except Exception:
        cred_path = os.path.join(SECRETS_DIR, "credentials.json")
        if os.path.exists(cred_path):
            with open(cred_path) as f:
                st.session_state.credentials = json.load(f)
        else:
            st.session_state.credentials = None

# ── Sidebar ──
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

    total_placas = len(st.session_state.placas)
    st.markdown(
        f'<div style="font-weight:600; font-size:0.9rem; margin-bottom:10px;">'
        f'Placas adicionadas <span class="badge-count">{total_placas}</span></div>',
        unsafe_allow_html=True
    )

    if st.session_state.placas:
        for i, placa in enumerate(st.session_state.placas):
            cliente = placa["dados"].get("Cliente", "—")
            pedido  = placa["dados"].get("N° do Pedido", "—")
            qtd     = placa["dados"].get("Quantidade de Placas", 1)

            col_info, col_del = st.columns([10, 1])
            with col_info:
                st.markdown(
                    f'<div style="font-size:0.82rem; font-weight:600; margin-top:4px;">'
                    f'{i+1}. {placa["tipo"]}</div>'
                    f'<div style="font-size:0.75rem; opacity:0.75;">'
                    f'{cliente} · Ped. {pedido} · Qtd. {qtd}</div>',
                    unsafe_allow_html=True
                )
            with col_del:
                if st.button("✕", key=f"del_{i}", help="Remover"):
                    st.session_state.placas.pop(i)
                    st.session_state.pdf_pronto = None
                    st.rerun()

        st.divider()
        if st.button("Limpar todas", use_container_width=True, type="primary"):
            st.session_state.placas     = []
            st.session_state.pdf_pronto = None
            st.session_state.relatorio  = None
            st.rerun()
    else:
        st.markdown(
            '<div style="font-size:0.82rem; opacity:0.7; padding: 6px 0;">'
            'Nenhuma placa adicionada ainda.</div>',
            unsafe_allow_html=True
        )

# ── Adicionar Placa ──
st.markdown('<div class="section-title">Adicionar Placa</div>', unsafe_allow_html=True)

tipo_selecionado = st.selectbox("Tipo de Placa", options=list(TEMPLATE_IDS.keys()), index=0)

if tipo_selecionado:
    campos = CAMPOS_COMUNS + CAMPOS_ESPECIFICOS.get(tipo_selecionado, [])
    st.caption(f"Campos para: {tipo_selecionado}")

    with st.container(border=True):
        col_esq, col_dir = st.columns(2)
        form_values = {}

        for idx, campo in enumerate(campos):
            col = col_esq if idx % 2 == 0 else col_dir
            key_widget = f"form_{tipo_selecionado}_{campo['key']}"
            with col:
                if campo["type"] == "number":
                    form_values[campo["key"]] = st.number_input(
                        campo["label"], min_value=1, value=1, key=key_widget,
                    )
                else:
                    form_values[campo["key"]] = st.text_input(
                        campo["label"], key=key_widget,
                    )

        col_btn, _ = st.columns([2, 6])
        with col_btn:
            adicionar = st.button("Adicionar à lista", type="primary", use_container_width=True)

    if adicionar:
        if not form_values.get("Cliente", "").strip():
            st.error("O campo Cliente é obrigatório.")
        elif not form_values.get("N° do Pedido", "").strip():
            st.error("O campo N° do Pedido é obrigatório.")
        else:
            form_values["Cliente"] = form_values["Cliente"].strip().upper()
            st.session_state.placas.append({"tipo": tipo_selecionado, "dados": form_values.copy()})
            st.session_state.pdf_pronto = None
            st.session_state.relatorio  = None
            st.success(f"{tipo_selecionado} adicionada à lista.")
            st.rerun()

# ── Gerar PDF ──
st.divider()
st.markdown('<div class="section-title">Gerar PDF Consolidado</div>', unsafe_allow_html=True)

if not st.session_state.placas:
    st.info("Adicione pelo menos uma placa para liberar a geração do PDF.")

elif not st.session_state.credentials:
    st.error("Credenciais não encontradas. Verifique a pasta secrets/ ou os Streamlit Secrets.")

else:
    cliente_principal = st.session_state.placas[0]["dados"].get("Cliente", "Cliente")
    pedido_principal  = st.session_state.placas[0]["dados"].get("N° do Pedido", "")
    nome_padrao       = f"Placas - {cliente_principal} ({pedido_principal})"
    nome_arquivo = st.text_input("Nome do arquivo PDF", value=nome_padrao)

    col_gerar, _ = st.columns([2, 6])
    with col_gerar:
        gerar = st.button("Gerar PDF", type="primary", use_container_width=True)

    if gerar:
        barra  = st.progress(0, text="Iniciando...")
        status = st.empty()

        def atualizar_progresso(pct: float, msg: str):
            barra.progress(pct, text=msg)
            status.markdown(f"_{msg}_")

        try:
            pdf_bytes, link_drive = gerar_pdf_consolidado(
                placas=st.session_state.placas,
                folder_id=FOLDER_ID,
                template_ids=TEMPLATE_IDS,
                nome_arquivo=nome_arquivo,
                progress_callback=atualizar_progresso,
            )
            st.session_state.pdf_pronto = pdf_bytes
            st.session_state.pdf_nome   = nome_arquivo
            st.session_state.link_drive = link_drive

            cliente_rel = st.session_state.placas[0]["dados"].get("Cliente", "")
            pedido_rel  = st.session_state.placas[0]["dados"].get("N° do Pedido", "")
            st.session_state.relatorio = gerar_relatorio(
                placas=st.session_state.placas,
                nome_cliente=cliente_rel,
                nome_pedido=pedido_rel,
            )

            barra.progress(1.0, text="Concluído!")
            status.success("PDF gerado e salvo na pasta de concluídos.")

        except Exception as e:
            barra.empty()
            status.empty()
            st.error(f"Erro ao gerar o PDF:\n\n```\n{e}\n```")

    if st.session_state.pdf_pronto:
        _, col_dl, col_rel, col_drive, _ = st.columns([1, 3, 3, 3, 1])

        with col_dl:
            st.download_button(
                label="Baixar PDF",
                data=st.session_state.pdf_pronto,
                file_name=f"{st.session_state.pdf_nome}.pdf",
                mime="application/pdf",
                use_container_width=True,
                type="primary",
            )

        with col_rel:
            if st.session_state.relatorio:
                st.download_button(
                    label="Baixar Relatório",
                    data=st.session_state.relatorio,
                    file_name=f"Relatório - {st.session_state.pdf_nome}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                )

        with col_drive:
            if st.session_state.link_drive:
                st.markdown(
                    '<a class="drive-link" href="' + st.session_state.link_drive + '" target="_blank">'
                    '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#FFFFFF" stroke-width="2">'
                    '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>'
                    '</svg>&nbsp;Abrir no Drive</a>',
                    unsafe_allow_html=True
                )

# ── Rodapé ──
ano = datetime.now().year
st.markdown(f"""
<div style="
    bottom: 0; left: 0; right: 0;
    border-top: 1px solid #e0e4f5;
    padding: 10px 24px;
    font-family: 'Sora', sans-serif;
    font-size: 0.78rem;
    color: #888;
    text-align: center;
">
    {ano} © FastPlac by <a href="https://aguiasistemas.com.br/" target="_blank" style="color:#242480; font-weight:700; text-decoration:none;">Águia Sistemas</a> &nbsp;·&nbsp; Version 1.0.0
</div>
""", unsafe_allow_html=True)