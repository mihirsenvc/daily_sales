"""Google Drive listing and chunked download for large .xlsx files."""

from __future__ import annotations

import io
import os
from typing import Callable, Optional

from dates import parse_date_from_name

ProgressFn = Optional[Callable[[int, str], None]]

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLS_MIME = "application/vnd.ms-excel"
SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
FOLDER_MIME = "application/vnd.google-apps.folder"

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


class DriveError(RuntimeError):
    pass


def credentials_path(root: str) -> str:
    return os.path.join(root, "credentials.json")


def token_path(root: str) -> str:
    return os.path.join(root, "token.json")


def has_oauth_files(root: str) -> bool:
    return os.path.isfile(credentials_path(root))


def get_service(root: str):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise DriveError(
            "Google Drive libraries are missing. Run start.bat (or pip install -r requirements.txt)."
        ) from exc

    creds_file = credentials_path(root)
    if not os.path.isfile(creds_file):
        raise DriveError(
            "Add credentials.json (Google Cloud OAuth Desktop client) next to app.py to read Drive folders."
        )

    creds = None
    token_file = token_path(root)
    if os.path.isfile(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w", encoding="utf-8") as fh:
            fh.write(creds.to_json())
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def list_folder_files(root: str, folder_id: str, recursive: bool = True) -> list:
    service = get_service(root)
    folder_id = (folder_id or "").strip()
    if not folder_id:
        raise DriveError("Enter a Google Drive folder ID.")
    collected = []
    _walk(service, folder_id, collected, depth=0, recursive=recursive)
    collected.sort(key=lambda f: (f["dateKey"], f.get("modifiedTime") or ""), reverse=True)
    return collected


def _walk(service, folder_id: str, acc: list, depth: int, recursive: bool) -> None:
    if depth > 2 or len(acc) >= 400:
        return
    query = (
        f"'{folder_id}' in parents and trashed = false and "
        f"(mimeType = '{XLSX_MIME}' or mimeType = '{XLS_MIME}' or "
        f"mimeType = '{SHEETS_MIME}' or mimeType = '{FOLDER_MIME}')"
    )
    page_token = None
    while True:
        resp = (
            service.files()
            .list(
                q=query,
                spaces="drive",
                fields="nextPageToken, files(id, name, mimeType, size, modifiedTime, owners)",
                pageToken=page_token,
                pageSize=100,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            )
            .execute()
        )
        for item in resp.get("files", []):
            mime = item.get("mimeType")
            if mime == FOLDER_MIME:
                if recursive:
                    _walk(service, item["id"], acc, depth + 1, recursive)
                continue
            date_key = parse_date_from_name(item.get("name") or "")
            if not date_key:
                continue
            acc.append(
                {
                    "id": item["id"],
                    "name": item.get("name") or "",
                    "mime": mime,
                    "size": int(item.get("size") or 0),
                    "modifiedTime": item.get("modifiedTime") or "",
                    "dateKey": date_key,
                    "source": "drive",
                }
            )
        page_token = resp.get("nextPageToken")
        if not page_token:
            break


def download_file(root: str, file_meta: dict, dest_path: str, progress: ProgressFn = None) -> str:
    from googleapiclient.http import MediaIoBaseDownload

    os.makedirs(os.path.dirname(dest_path) or ".", exist_ok=True)
    service = get_service(root)
    file_id = file_meta["id"]
    mime = file_meta.get("mime") or ""
    if mime == SHEETS_MIME:
        request = service.files().export_media(fileId=file_id, mimeType=XLSX_MIME)
    else:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

    with io.FileIO(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request, chunksize=8 * 1024 * 1024)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            if status and progress:
                pct = 4 + int(status.progress() * 18)
                progress(pct, f"Downloading {file_meta.get('name', '')}… {int(status.progress() * 100)}%")
    return dest_path
