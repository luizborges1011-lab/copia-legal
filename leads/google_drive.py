"""Google Drive integration — creates client folders using a Service Account.

Required env var:
    GOOGLE_SERVICE_ACCOUNT_JSON  — full JSON content of the service account key file

The service account must have the folder CLIENTES_ATIVOS_FOLDER_ID shared with it
(Editor or Content Manager permission).
"""

import json
import os

CLIENTES_ATIVOS_FOLDER_ID = "18zvvTtSIJqATHZ5kW8czZ6VOTx_S8I6S"
_SCOPES = ["https://www.googleapis.com/auth/drive"]


def _get_service():
    creds_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not creds_json:
        raise RuntimeError("Variável de ambiente GOOGLE_SERVICE_ACCOUNT_JSON não configurada.")
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    info = json.loads(creds_json)
    creds = service_account.Credentials.from_service_account_info(info, scopes=_SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def create_lead_folder(lead_name: str) -> str:
    """Creates a folder named `lead_name` inside CLIENTES ATIVOS.
    Returns the public folder URL.
    """
    service = _get_service()
    metadata = {
        "name": lead_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [CLIENTES_ATIVOS_FOLDER_ID],
    }
    folder = service.files().create(body=metadata, fields="id").execute()
    folder_id = folder["id"]
    return f"https://drive.google.com/drive/folders/{folder_id}"
