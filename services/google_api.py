"""
services/google_api.py  –  Integração com Google Drive + Slides API (OAuth2)
"""
import io
import os
import json
import time
import tempfile
import requests as http_requests
import google.auth.transport.requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from pypdf import PdfWriter, PdfReader
from pptx import Presentation as PptxPresentation
import copy

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

_BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
_SECRETS_DIR = os.path.join(_BASE_DIR, "secrets")
TOKEN_FILE   = os.path.join(_SECRETS_DIR, "token.json")
CREDS_FILE   = os.path.join(_SECRETS_DIR, "client_secret.json")


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


def _execute_with_retry(request, max_retries: int = 5):
    for attempt in range(max_retries):
        try:
            return request.execute()
        except HttpError as e:
            if e.resp.status == 429 and attempt < max_retries - 1:
                wait = 2 ** (attempt + 1)
                print(f"[RATE LIMIT] Aguardando {wait}s...")
                time.sleep(wait)
            else:
                raise


def copy_template(drive, template_id: str, name: str, folder_id: str) -> str:
    body = {"name": name, "parents": [folder_id]}
    result = _execute_with_retry(drive.files().copy(fileId=template_id, body=body))
    return result["id"]


def rename_file(drive, file_id: str, new_name: str) -> str:
    """Renomeia o arquivo e retorna o webViewLink."""
    result = _execute_with_retry(
        drive.files().update(
            fileId=file_id,
            body={"name": new_name},
            fields="id, webViewLink",
        )
    )
    return result.get("webViewLink", "")


def _export_as_pptx(creds, presentation_id: str) -> bytes:
    """Exporta um Google Slides como PPTX e retorna os bytes."""
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url  = f"https://docs.google.com/presentation/d/{presentation_id}/export/pptx"
    resp = http_requests.get(url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=60)
    resp.raise_for_status()
    return resp.content


def _merge_pptx(pptx_bytes_list: list[bytes]) -> bytes:
    """
    Mescla múltiplos PPTX trabalhando direto no ZIP/XML.
    Preserva slides, mídias e relacionamentos de cada arquivo.
    """
    import zipfile as _zf, re as _re
    from lxml import etree as _et

    NS_PPTX   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    NS_REL    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    NS_CT     = 'http://schemas.openxmlformats.org/package/2006/content-types'
    REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide'
    CT_SLIDE  = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'

    zips       = [_zf.ZipFile(io.BytesIO(b), 'r') for b in pptx_bytes_list]
    base_zip   = zips[0]
    base_names = set(base_zip.namelist())
    out_buf    = io.BytesIO()

    slide_count = len([n for n in base_names if _re.match('ppt/slides/slide[0-9]+[.]xml$', n)])
    media_names = {n for n in base_names if n.startswith('ppt/media/')}
    extra_files = {}
    new_pairs   = []  # (zip_path, rel_id)

    for src_zip in zips[1:]:
        src_names  = set(src_zip.namelist())
        src_slides = sorted(
            [n for n in src_names if _re.match('ppt/slides/slide[0-9]+[.]xml$', n)],
            key=lambda x: int(_re.search('[0-9]+', x).group())
        )
        for name in src_names:
            if name.startswith('ppt/media/') and name not in media_names:
                extra_files[name] = src_zip.read(name)
                media_names.add(name)
        for slide_path in src_slides:
            slide_count += 1
            new_path = f'ppt/slides/slide{slide_count}.xml'
            extra_files[new_path] = src_zip.read(slide_path)
            rel_src = slide_path.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
            rel_dst = new_path.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
            if rel_src in src_names:
                extra_files[rel_dst] = src_zip.read(rel_src)
            new_pairs.append((new_path, f'rId{100 + slide_count}'))

    # presentation.xml
    pres_root = _et.fromstring(base_zip.read('ppt/presentation.xml'))
    sldIdLst  = pres_root.find(f'{{{NS_PPTX}}}sldIdLst')
    max_id    = max((int(e.get('id', 256)) for e in sldIdLst.findall(f'{{{NS_PPTX}}}sldId')), default=256)
    for _, rel_id in new_pairs:
        max_id += 1
        el = _et.SubElement(sldIdLst, f'{{{NS_PPTX}}}sldId')
        el.set('id', str(max_id))
        el.set(f'{{{NS_REL}}}id', rel_id)

    # presentation.xml.rels
    rels_root = _et.fromstring(base_zip.read('ppt/_rels/presentation.xml.rels'))
    for slide_path, rel_id in new_pairs:
        el = _et.SubElement(rels_root, 'Relationship')
        el.set('Id', rel_id)
        el.set('Type', REL_SLIDE)
        el.set('Target', slide_path.replace('ppt/', ''))

    # [Content_Types].xml
    ct_root = _et.fromstring(base_zip.read('[Content_Types].xml'))
    for slide_path, _ in new_pairs:
        ov = _et.SubElement(ct_root, f'{{{NS_CT}}}Override')
        ov.set('PartName', f'/{slide_path}')
        ov.set('ContentType', CT_SLIDE)

    with _zf.ZipFile(out_buf, 'w', _zf.ZIP_DEFLATED) as out_zip:
        for name in base_zip.namelist():
            if name == 'ppt/presentation.xml':
                out_zip.writestr(name, _et.tostring(pres_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            elif name == 'ppt/_rels/presentation.xml.rels':
                out_zip.writestr(name, _et.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            elif name == '[Content_Types].xml':
                out_zip.writestr(name, _et.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True))
            else:
                out_zip.writestr(name, base_zip.read(name))
        for name, data in extra_files.items():
            out_zip.writestr(name, data)

    for z in zips:
        z.close()
    out_buf.seek(0)
    return out_buf.read()


def create_consolidated_slides(
    drive, creds, file_ids: list[str], nome: str, folder_id: str
) -> str:
    """
    Exporta cada Slides como PPTX, mescla e faz upload convertendo
    para Google Slides editável. Retorna o webViewLink.
    """
    pptx_list = []
    for file_id in file_ids:
        pptx_list.append(_export_as_pptx(creds, file_id))
        time.sleep(0.5)

    merged_pptx = _merge_pptx(pptx_list)

    metadata = {
        "name": nome,
        "parents": [folder_id],
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
    return result.get("webViewLink", "")


def delete_file(drive, file_id: str):
    try:
        drive.files().delete(fileId=file_id).execute()
    except HttpError as e:
        print(f"[WARN] Não foi possível deletar {file_id}: {e}")


def upload_pdf(drive, pdf_bytes: bytes, nome: str, folder_id: str) -> str:
    metadata = {"name": f"{nome}.pdf", "parents": [folder_id], "mimeType": "application/pdf"}
    media = MediaIoBaseUpload(io.BytesIO(pdf_bytes), mimetype="application/pdf")
    arquivo = _execute_with_retry(
        drive.files().create(body=metadata, media_body=media, fields="id, webViewLink")
    )
    return arquivo.get("webViewLink", "")


def duplicate_first_slide(slides_svc, presentation_id: str, extra_copies: int):
    pres = _execute_with_retry(slides_svc.presentations().get(presentationId=presentation_id))
    first_id = pres["slides"][0]["objectId"]
    for _ in range(extra_copies):
        time.sleep(1)
        _execute_with_retry(
            slides_svc.presentations().batchUpdate(
                presentationId=presentation_id,
                body={"requests": [{"duplicateObject": {"objectId": first_id}}]},
            )
        )


def fill_placeholders(slides_svc, presentation_id: str, data: dict):
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": f"{{{{{key}}}}}"},
                "replaceText": str(value) if value not in (None, "") else "",
            }
        }
        for key, value in data.items()
    ]
    if requests:
        _execute_with_retry(
            slides_svc.presentations().batchUpdate(
                presentationId=presentation_id,
                body={"requests": requests},
            )
        )


def export_as_pdf(creds, presentation_id: str) -> bytes:
    if creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
    url = f"https://docs.google.com/presentation/d/{presentation_id}/export/pdf"
    resp = http_requests.get(url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=60)
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


def gerar_pdf_consolidado(
    placas: list[dict],
    folder_id: str,
    template_ids: dict,
    nome_arquivo: str = "Placas",
    progress_callback=None,
) -> tuple[bytes, str, list[dict]]:
    """
    Retorna:
        merged      — bytes do PDF consolidado
        link_pdf    — link do PDF no Drive
        slides_info — lista de dicts com {tipo, cliente, link} para cada Slides individual
        link_slides — link do Google Slides consolidado no Drive
    """
    drive, slides_svc, creds = get_services()
    pdf_parts:            list[bytes] = []
    slides_info:          list[dict]  = []
    file_ids_preenchidos: list[str]   = []
    total = len(placas)

    for idx, placa in enumerate(placas):
        tipo  = placa["tipo"]
        dados = placa["dados"]
        qtd   = max(1, int(dados.get("Quantidade de Placas") or 1))

        if progress_callback:
            progress_callback(idx / total, f"Processando placa {idx + 1}/{total}: {tipo}")

        # Nome final do Slides: "Slides - <tipo> - <cliente> - <pedido>"
        cliente    = dados.get("Cliente", "")
        pedido     = dados.get("N° do Pedido", "")
        nome_slide = f"Slides - {cliente} - {pedido}"
        nome_temp  = f"TEMP_{nome_slide}_{int(time.time() * 1000)}"

        file_id = copy_template(drive, template_ids[tipo], nome_temp, folder_id)

        if qtd > 1:
            duplicate_first_slide(slides_svc, file_id, qtd - 1)

        fill_placeholders(slides_svc, file_id, dados)
        pdf_bytes = export_as_pdf(creds, file_id)
        pdf_parts.append(pdf_bytes)

        # Renomeia e mantém o Slides individual no Drive
        link_slide = rename_file(drive, file_id, nome_slide)
        slides_info.append({"tipo": tipo, "cliente": cliente, "link": link_slide})
        file_ids_preenchidos.append(file_id)

        if idx < total - 1:
            time.sleep(1.5)

    if progress_callback: progress_callback(0.88, "Juntando os PDFs...")
    merged = merge_pdfs(pdf_parts)

    if progress_callback: progress_callback(0.92, "Criando Slides consolidado...")
    nome_slides_consolidado = f"Slides - {nome_arquivo}"
    link_slides = create_consolidated_slides(
        drive, creds, file_ids_preenchidos, nome_slides_consolidado, folder_id
    )

    # Apaga os Slides individuais após consolidar
    if progress_callback: progress_callback(0.94, "Removendo arquivos temporários...")
    for file_id in file_ids_preenchidos:
        delete_file(drive, file_id)

    if progress_callback: progress_callback(0.96, "Salvando PDF na pasta de concluídos...")
    link_pdf = upload_pdf(drive, merged, nome_arquivo, folder_id)

    if progress_callback: progress_callback(1.0, "Concluído!")
    return merged, link_pdf, slides_info, link_slides