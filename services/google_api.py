import io
import os
import re
import json
import time
import copy
import html as _html_mod
import zipfile
import tempfile
import requests as http_requests
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from pypdf import PdfWriter, PdfReader
from lxml import etree

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

_BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
_SECRETS_DIR = os.path.join(_BASE_DIR, "secrets")
TOKEN_FILE   = os.path.join(_SECRETS_DIR, "token.json")
CREDS_FILE   = os.path.join(_SECRETS_DIR, "client_secret.json")


# ─────────────────────────────────────────────────────────────
#  Autenticação
# ─────────────────────────────────────────────────────────────

def get_services():
    creds = None

    try:
        import streamlit as st
        if "google" in st.secrets and "token" in st.secrets["google"]:
            token_info = json.loads(st.secrets["google"]["token"])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    except Exception:
        pass

    if creds is None and os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
        try:
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
        except Exception:
            pass

    if not creds or not creds.valid:
        try:
            import streamlit as st
            if "google" in st.secrets and "client_secret" in st.secrets["google"]:
                client_info = json.loads(st.secrets["google"]["client_secret"])
                with tempfile.NamedTemporaryFile(mode="w", suffix=".json", delete=False) as tmp:
                    json.dump(client_info, tmp)
                    tmp_path = tmp.name
                flow = InstalledAppFlow.from_client_secrets_file(tmp_path, SCOPES)
                os.unlink(tmp_path)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
        except Exception:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)

        creds = flow.run_local_server(port=0)
        os.makedirs(_SECRETS_DIR, exist_ok=True)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    drive  = build("drive",  "v3", credentials=creds)
    slides = build("slides", "v1", credentials=creds)
    return drive, slides, creds


# ─────────────────────────────────────────────────────────────
#  Retry para Drive API
# ─────────────────────────────────────────────────────────────

def _execute_with_retry(request, max_retries: int = 7):
    for attempt in range(max_retries):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status in (429, 500, 503) and attempt < max_retries - 1:
                wait = min(2 ** (attempt + 2), 60)
                print(f"[RETRY {e.resp.status}] Aguardando {wait}s (tentativa {attempt + 1}/{max_retries})...")
                time.sleep(wait)
            else:
                raise


# ─────────────────────────────────────────────────────────────
#  Download de template como PPTX (com cache por tipo)
# ─────────────────────────────────────────────────────────────

def _download_template_pptx(creds, template_id: str) -> bytes:
    """Baixa um Google Slides como PPTX."""
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url  = f"https://docs.google.com/presentation/d/{template_id}/export/pptx"
    resp = http_requests.get(
        url,
        headers={"Authorization": f"Bearer {creds.token}"},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.content


# ─────────────────────────────────────────────────────────────
#  Preenchimento LOCAL de placeholders via XML direto no ZIP
# ─────────────────────────────────────────────────────────────

def _xml_escape(text: str) -> str:
    """Escapa caracteres especiais XML no valor substituído."""
    return _html_mod.escape(str(text), quote=False)


def _substituir_em_paragrafo_xml(para_xml: str, replacements: dict) -> str:
    """
    Recebe o XML de um parágrafo <a:p>, concatena o texto de todos os <a:t>,
    aplica os replacements (case-insensitive) e reconstrói colocando o texto
    resultante no primeiro <a:t>, zerando os demais.

    Trabalha diretamente no XML — sem python-pptx — então imagens,
    QR codes e qualquer outro elemento não-texto ficam intactos.
    """
    texts = re.findall(r'<a:t(?:[^>]*)>(.*?)</a:t>', para_xml, re.DOTALL)
    if not texts:
        return para_xml

    full_text = ''.join(texts)
    new_text  = full_text

    for key, value in replacements.items():
        pattern  = re.escape(f'{{{{{key}}}}}')
        new_text = re.sub(pattern, _xml_escape(value), new_text, flags=re.IGNORECASE)

    if new_text == full_text:
        return para_xml  # Sem alteração — retorna intacto

    # Reconstrói: primeiro <a:t> recebe o novo texto, demais ficam vazios
    first_done = False

    def replace_t(m):
        nonlocal first_done
        attrs = m.group(1)  # atributos do <a:t>, ex: xml:space="preserve"
        if not first_done:
            first_done = True
            return f'<a:t{attrs}>{new_text}</a:t>'
        return '<a:t></a:t>'

    return re.sub(r'<a:t([^>]*)>(.*?)</a:t>', replace_t, para_xml, flags=re.DOTALL)


def _fill_pptx_placeholders(pptx_bytes: bytes, data: dict) -> bytes:
    """
    Substitui {{chave}} diretamente no XML do PPTX via manipulação de ZIP.

    Não passa pelo modelo de objetos do python-pptx, portanto:
    - Preserva 100% das imagens e QR codes originais
    - Lida com placeholders fragmentados entre múltiplos runs
    - Não reserializa o XML (zero risco de perda de elementos)
    - Matching case-insensitive (replica comportamento da Slides API)
    """
    replacements = {
        k: (str(v).strip() if v not in (None, "") else "")
        for k, v in data.items()
    }

    src_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(src_buf, 'r') as src_zip:
        with zipfile.ZipFile(out_buf, 'w', zipfile.ZIP_DEFLATED) as out_zip:
            for name in src_zip.namelist():
                raw = src_zip.read(name)

                # Processa apenas XMLs de slides — não toca em imagens, layouts, etc.
                if re.match(r'ppt/slides/slide\d+\.xml$', name):
                    xml = raw.decode('utf-8')
                    xml = re.sub(
                        r'<a:p\b[^>]*>.*?</a:p>',
                        lambda m: _substituir_em_paragrafo_xml(m.group(0), replacements),
                        xml,
                        flags=re.DOTALL,
                    )
                    raw = xml.encode('utf-8')

                out_zip.writestr(name, raw)

    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Duplicação LOCAL do primeiro slide (sem API)
# ─────────────────────────────────────────────────────────────

def _duplicate_first_slide_pptx(pptx_bytes: bytes, extra_copies: int) -> bytes:
    """
    Duplica o primeiro slide N vezes dentro de um PPTX, operando
    diretamente no ZIP/XML sem chamar nenhuma API.
    """
    if extra_copies <= 0:
        return pptx_bytes

    NS_PPTX   = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_REL    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_CT     = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    CT_SLIDE  = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    src_buf = io.BytesIO(pptx_bytes)
    out_buf = io.BytesIO()

    with zipfile.ZipFile(src_buf, "r") as src_zip:
        names = src_zip.namelist()

        slide_names = sorted(
            [n for n in names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)],
            key=lambda x: int(re.search(r"[0-9]+", x).group()),
        )
        if not slide_names:
            return pptx_bytes

        first_slide     = slide_names[0]
        first_slide_num = int(re.search(r"[0-9]+", first_slide).group())
        max_slide_num   = max(int(re.search(r"[0-9]+", n).group()) for n in slide_names)

        pres_root = etree.fromstring(src_zip.read("ppt/presentation.xml"))
        rels_root = etree.fromstring(src_zip.read("ppt/_rels/presentation.xml.rels"))
        ct_root   = etree.fromstring(src_zip.read("[Content_Types].xml"))

        sldIdLst = pres_root.find(f"{{{NS_PPTX}}}sldIdLst")
        max_id   = max(
            (int(e.get("id", 256)) for e in sldIdLst.findall(f"{{{NS_PPTX}}}sldId")),
            default=256,
        )
        max_rel_id = max(
            (int(re.sub(r"\D", "", e.get("Id", "0")) or 0) for e in rels_root),
            default=100,
        )

        extra_files = {}
        for i in range(extra_copies):
            new_num      = max_slide_num + i + 1
            new_path     = f"ppt/slides/slide{new_num}.xml"
            new_rel_path = f"ppt/slides/_rels/slide{new_num}.xml.rels"
            rel_id       = f"rId{max_rel_id + i + 1}"

            extra_files[new_path] = src_zip.read(first_slide)

            first_rel = f"ppt/slides/_rels/slide{first_slide_num}.xml.rels"
            if first_rel in names:
                extra_files[new_rel_path] = src_zip.read(first_rel)

            max_id += 1
            el = etree.SubElement(sldIdLst, f"{{{NS_PPTX}}}sldId")
            el.set("id", str(max_id))
            el.set(f"{{{NS_REL}}}id", rel_id)

            rel_el = etree.SubElement(rels_root, "Relationship")
            rel_el.set("Id", rel_id)
            rel_el.set("Type", REL_SLIDE)
            rel_el.set("Target", f"slides/slide{new_num}.xml")

            ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
            ov.set("PartName", f"/{new_path}")
            ov.set("ContentType", CT_SLIDE)

        with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
            for name in names:
                if name == "ppt/presentation.xml":
                    out_zip.writestr(name, etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                elif name == "ppt/_rels/presentation.xml.rels":
                    out_zip.writestr(name, etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                elif name == "[Content_Types].xml":
                    out_zip.writestr(name, etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True))
                else:
                    out_zip.writestr(name, src_zip.read(name))

            for name, data in extra_files.items():
                out_zip.writestr(name, data)

    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Mesclagem LOCAL de múltiplos PPTX
# ─────────────────────────────────────────────────────────────

def _merge_pptx(pptx_bytes_list: list[bytes]) -> bytes:
    """
    Mescla múltiplos PPTX trabalhando direto no ZIP/XML.
    Preserva slides, mídias e relacionamentos de cada arquivo.
    """
    NS_PPTX   = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_REL    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    NS_CT     = "http://schemas.openxmlformats.org/package/2006/content-types"
    REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    CT_SLIDE  = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

    zips       = [zipfile.ZipFile(io.BytesIO(b), "r") for b in pptx_bytes_list]
    base_zip   = zips[0]
    base_names = set(base_zip.namelist())
    out_buf    = io.BytesIO()

    slide_count = len([n for n in base_names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)])
    media_names = {n for n in base_names if n.startswith("ppt/media/")}
    extra_files = {}
    new_pairs   = []

    for src_zip in zips[1:]:
        src_names  = set(src_zip.namelist())
        src_slides = sorted(
            [n for n in src_names if re.match(r"ppt/slides/slide[0-9]+\.xml$", n)],
            key=lambda x: int(re.search(r"[0-9]+", x).group()),
        )
        for name in src_names:
            if name.startswith("ppt/media/") and name not in media_names:
                extra_files[name] = src_zip.read(name)
                media_names.add(name)
        for slide_path in src_slides:
            slide_count += 1
            new_path = f"ppt/slides/slide{slide_count}.xml"
            extra_files[new_path] = src_zip.read(slide_path)
            rel_src = slide_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            rel_dst = new_path.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
            if rel_src in src_names:
                extra_files[rel_dst] = src_zip.read(rel_src)
            new_pairs.append((new_path, f"rId{100 + slide_count}"))

    pres_root = etree.fromstring(base_zip.read("ppt/presentation.xml"))
    sldIdLst  = pres_root.find(f"{{{NS_PPTX}}}sldIdLst")
    max_id    = max(
        (int(e.get("id", 256)) for e in sldIdLst.findall(f"{{{NS_PPTX}}}sldId")),
        default=256,
    )
    for _, rel_id in new_pairs:
        max_id += 1
        el = etree.SubElement(sldIdLst, f"{{{NS_PPTX}}}sldId")
        el.set("id", str(max_id))
        el.set(f"{{{NS_REL}}}id", rel_id)

    rels_root = etree.fromstring(base_zip.read("ppt/_rels/presentation.xml.rels"))
    for slide_path, rel_id in new_pairs:
        el = etree.SubElement(rels_root, "Relationship")
        el.set("Id", rel_id)
        el.set("Type", REL_SLIDE)
        el.set("Target", slide_path.replace("ppt/", ""))

    ct_root = etree.fromstring(base_zip.read("[Content_Types].xml"))
    for slide_path, _ in new_pairs:
        ov = etree.SubElement(ct_root, f"{{{NS_CT}}}Override")
        ov.set("PartName", f"/{slide_path}")
        ov.set("ContentType", CT_SLIDE)

    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as out_zip:
        for name in base_zip.namelist():
            if name == "ppt/presentation.xml":
                out_zip.writestr(name, etree.tostring(pres_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "ppt/_rels/presentation.xml.rels":
                out_zip.writestr(name, etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            elif name == "[Content_Types].xml":
                out_zip.writestr(name, etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True))
            else:
                out_zip.writestr(name, base_zip.read(name))
        for name, data in extra_files.items():
            out_zip.writestr(name, data)

    for z in zips:
        z.close()
    out_buf.seek(0)
    return out_buf.read()


# ─────────────────────────────────────────────────────────────
#  Drive helpers
# ─────────────────────────────────────────────────────────────

def rename_file(drive, file_id: str, new_name: str) -> str:
    result = _execute_with_retry(
        drive.files().update(
            fileId=file_id,
            body={"name": new_name},
            fields="id, webViewLink",
        )
    )
    return result.get("webViewLink", "")


def delete_file(drive, file_id: str):
    try:
        drive.files().delete(fileId=file_id).execute()
    except HttpError as e:
        print(f"[WARN] Não foi possível deletar {file_id}: {e}")


def upload_pdf(drive, pdf_bytes: bytes, nome: str, folder_id: str) -> str:
    metadata = {"name": f"{nome}.pdf", "parents": [folder_id], "mimeType": "application/pdf"}
    media    = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
    arquivo  = _execute_with_retry(
        drive.files().create(body=metadata, media_body=media, fields="id, webViewLink")
    )
    return arquivo.get("webViewLink", "")


def export_as_pdf(creds, presentation_id: str) -> bytes:
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url  = f"https://docs.google.com/presentation/d/{presentation_id}/export/pdf"
    resp = http_requests.get(
        url,
        headers={"Authorization": f"Bearer {creds.token}"},
        timeout=120,
    )
    resp.raise_for_status()
    return resp.content


def merge_pdfs(pdf_bytes_list: list[bytes]) -> bytes:
    writer = PdfWriter()
    for pdf_bytes in pdf_bytes_list:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────
#  Função principal
# ─────────────────────────────────────────────────────────────

def gerar_pdf_consolidado(
    placas: list[dict],
    folder_id: str,
    template_ids: dict,
    nome_arquivo: str = "Placas",
    progress_callback=None,
) -> tuple[bytes, str, list[dict], str]:
    """
    1. Baixa cada template UMA vez por tipo (cache em memória)
    2. Preenche placeholders diretamente no XML do ZIP (preserva imagens/QR codes)
    3. Duplica slides localmente se Qtd > 1
    4. Mescla todos os PPTX localmente
    5. Faz 1 upload do PPTX mesclado → exporta como PDF
    6. Mantém o PPTX mesclado como Slides consolidado no Drive
    7. Faz 1 upload do PDF final
    """
    drive, _, creds = get_services()
    total = len(placas)

    # ── 1. Download dos templates (por tipo, com cache) ──
    if progress_callback:
        progress_callback(0.02, "Baixando templates...")

    tipos_unicos: list[str]       = list({p["tipo"] for p in placas})
    template_cache: dict[str, bytes] = {}

    for i, tipo in enumerate(tipos_unicos):
        if progress_callback:
            progress_callback(
                0.02 + (i / len(tipos_unicos)) * 0.15,
                f"Baixando template {i + 1}/{len(tipos_unicos)}: {tipo}",
            )
        template_cache[tipo] = _download_template_pptx(creds, template_ids[tipo])
        time.sleep(0.3)

    # ── 2 + 3. Preenche e duplica cada placa localmente ──
    filled_pptx_list: list[bytes] = []
    slides_info:      list[dict]  = []

    for idx, placa in enumerate(placas):
        tipo  = placa["tipo"]
        dados = placa["dados"]
        qtd   = max(1, int(dados.get("Quantidade de Placas") or 1))

        if progress_callback:
            pct = 0.17 + (idx / total) * 0.55
            progress_callback(pct, f"Preenchendo placa {idx + 1}/{total}: {tipo}")

        filled = _fill_pptx_placeholders(template_cache[tipo], dados)

        if qtd > 1:
            filled = _duplicate_first_slide_pptx(filled, qtd - 1)

        filled_pptx_list.append(filled)

        cliente = dados.get("Cliente", "")
        slides_info.append({"tipo": tipo, "cliente": cliente, "link": ""})

    # ── 4. Mescla todos os PPTX localmente ──
    if progress_callback:
        progress_callback(0.73, f"Mesclando {total} apresentações...")

    merged_pptx = _merge_pptx(filled_pptx_list)

    # ── 5. Faz upload do PPTX mesclado como Google Slides ──
    if progress_callback:
        progress_callback(0.80, "Enviando para o Google Drive...")

    nome_slides = f"Slides - {nome_arquivo}"
    metadata = {
        "name":     nome_slides,
        "parents":  [folder_id],
        "mimeType": "application/vnd.google-apps.presentation",
    }
    media = MediaIoBaseUpload(
        io.BytesIO(merged_pptx),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        resumable=True,
    )
    result = _execute_with_retry(
        drive.files().create(body=metadata, media_body=media, fields="id, webViewLink")
    )
    slides_id   = result["id"]
    link_slides = result.get("webViewLink", "")

    # ── 6. Exporta como PDF a partir do Slides consolidado ──
    if progress_callback:
        progress_callback(0.90, "Exportando PDF...")

    time.sleep(3)
    pdf_bytes = export_as_pdf(creds, slides_id)

    # ── 7. Faz upload do PDF ──
    if progress_callback:
        progress_callback(0.96, "Salvando PDF na pasta de concluídos...")

    link_pdf = upload_pdf(drive, pdf_bytes, nome_arquivo, folder_id)

    if progress_callback:
        progress_callback(1.0, "Concluído!")

    return pdf_bytes, link_pdf, slides_info, link_slides