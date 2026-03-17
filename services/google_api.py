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

SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
]

# Paths para secrets/ (raiz do projeto)
_BASE_DIR    = os.path.dirname(os.path.dirname(__file__))
_SECRETS_DIR = os.path.join(_BASE_DIR, "secrets")
TOKEN_FILE   = os.path.join(_SECRETS_DIR, "token.json")
CREDS_FILE   = os.path.join(_SECRETS_DIR, "client_secret.json")


def get_services():
    creds = None

    # 1. Produção: token via Streamlit Secrets
    try:
        import streamlit as st
        if "google" in st.secrets and "token" in st.secrets["google"]:
            token_info = json.loads(st.secrets["google"]["token"])
            creds = Credentials.from_authorized_user_info(token_info, SCOPES)
    except Exception:
        pass

    # 2. Desenvolvimento: token em secrets/token.json
    if creds is None and os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # 3. Renova token se expirado
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(google.auth.transport.requests.Request())
        try:
            with open(TOKEN_FILE, "w") as f:
                f.write(creds.to_json())
        except Exception:
            pass

    # 4. Sem token: abre browser para login (só funciona local)
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
) -> tuple[bytes, str]:
    drive, slides_svc, creds = get_services()
    pdf_parts: list[bytes] = []
    total = len(placas)

    for idx, placa in enumerate(placas):
        tipo  = placa["tipo"]
        dados = placa["dados"]
        qtd   = max(1, int(dados.get("Quantidade de Placas") or 1))

        if progress_callback:
            progress_callback(idx / total, f"Processando placa {idx + 1}/{total}: {tipo}")

        nome_temp = f"TEMP_{tipo}_{dados.get('Cliente', '')}_{int(time.time() * 1000)}"
        file_id   = copy_template(drive, template_ids[tipo], nome_temp, folder_id)

        if qtd > 1:
            duplicate_first_slide(slides_svc, file_id, qtd - 1)

        fill_placeholders(slides_svc, file_id, dados)
        pdf_bytes = export_as_pdf(creds, file_id)
        pdf_parts.append(pdf_bytes)
        delete_file(drive, file_id)

        if idx < total - 1:
            time.sleep(1.5)

    if progress_callback: progress_callback(0.90, "Juntando os PDFs...")
    merged = merge_pdfs(pdf_parts)

    if progress_callback: progress_callback(0.95, "Salvando na pasta de concluídos...")
    link_drive = upload_pdf(drive, merged, nome_arquivo, folder_id)

    if progress_callback: progress_callback(1.0, "Concluído!")
    return merged, link_drive