import os
import json
import logging
import ssl
import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util import Retry
import msal
from pathlib import Path
from urllib.parse import quote

import config

logger = logging.getLogger("crm-automation")

def create_retry_session(
    retries: int = 5,
    backoff_factor: float = 1.0,
    status_forcelist: tuple[int, ...] = (429, 500, 502, 503, 504),
) -> requests.Session:
    """Create a requests Session equipped with urllib3 Retry adapter for transient connection/SSL errors."""
    session = requests.Session()
    retry_strategy = Retry(
        total=retries,
        connect=retries,
        read=retries,
        status=retries,
        backoff_factor=backoff_factor,
        status_forcelist=list(status_forcelist),
        allowed_methods=frozenset(["GET", "POST", "PUT", "DELETE", "HEAD", "OPTIONS"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry_strategy)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session

class AuthProvider:
    """Acquires access tokens from Azure AD for Microsoft Graph API."""
    def __init__(self) -> None:
        self.tenant_id = config.AZURE_TENANT_ID
        self.client_id = config.AZURE_CLIENT_ID
        self.client_secret = config.AZURE_CLIENT_SECRET
        self._msal_app = None

    @property
    def msal_app(self) -> msal.ConfidentialClientApplication:
        if self._msal_app is None:
            authority = f"https://login.microsoftonline.com/{self.tenant_id}"
            self._msal_app = msal.ConfidentialClientApplication(
                client_id=self.client_id,
                client_credential=self.client_secret,
                authority=authority,
            )
        return self._msal_app

    def get_access_token(self) -> str:
        # Default scope for MS Graph daemon apps
        result = self.msal_app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            error_desc = result.get("error_description", result.get("error", "Unknown error"))
            raise RuntimeError(f"MSAL Token acquisition failed: {error_desc}")
        return result["access_token"]

    def get_headers(self) -> dict[str, str]:
        token = self.get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

class SharePointClient:
    """Downloads and uploads files from/to SharePoint via MS Graph API."""
    def __init__(self, auth: AuthProvider, default_drive_id: str = "") -> None:
        self.auth = auth
        self.default_drive_id = default_drive_id or config.SHAREPOINT_SOURCE_DRIVE_ID
        self.session = create_retry_session()

    def _url(self, path: str = "", drive_id: str = "") -> str:
        active_drive = drive_id or self.default_drive_id
        return f"{config.GRAPH_BASE}/drives/{active_drive}/{path}".rstrip("/")

    def download_file(
        self,
        remote_file_path: str,
        local_path: Path,
        drive_id: str = "",
        max_retries: int = 3,
        backoff_factor: float = 2.0,
    ) -> Path:
        """Download file by its path relative to the drive root with automatic retry."""
        logger.info("Downloading file from SharePoint: %s ...", remote_file_path)
        escaped_path = "/".join(quote(p) for p in remote_file_path.split("/"))
        url = self._url(f"root:/{escaped_path}", drive_id=drive_id)
        
        last_error = None
        for attempt in range(1, max_retries + 1):
            temp_file = local_path.with_suffix(f"{local_path.suffix}.tmp")
            try:
                # 1. Get metadata to get fresh download URL
                response = self.session.get(url, headers=self.auth.get_headers(), timeout=30)
                response.raise_for_status()
                
                dl_url = response.json().get("@microsoft.graph.downloadUrl")
                if not dl_url:
                    raise ValueError(f"No download URL found for remote path: {remote_file_path}")
                    
                local_path.parent.mkdir(parents=True, exist_ok=True)
                
                # 2. Stream the download into a temporary file first
                dl_response = self.session.get(dl_url, stream=True, timeout=(15, 120))
                dl_response.raise_for_status()
                
                with open(temp_file, "wb") as f:
                    for chunk in dl_response.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)
                
                # Atomic replace on completion
                if temp_file.exists():
                    temp_file.replace(local_path)
                    
                logger.info("[OK] Download complete: %s (%d bytes)", local_path.name, local_path.stat().st_size)
                return local_path
            except (requests.exceptions.RequestException, ssl.SSLError, ConnectionError, OSError) as e:
                last_error = e
                if temp_file.exists():
                    try:
                        temp_file.unlink()
                    except Exception:
                        pass
                if attempt < max_retries:
                    wait_time = backoff_factor ** attempt
                    logger.warning(
                        "[SharePointClient] Download failed (attempt %d/%d): %s. Retrying in %ds...",
                        attempt, max_retries, e, wait_time
                    )
                    time.sleep(wait_time)
                else:
                    logger.error("[SharePointClient] Download failed after %d attempts for %s: %s", max_retries, remote_file_path, e)
                    raise
        if last_error:
            raise last_error

    def check_file_exists(self, remote_file_path: str, drive_id: str = "") -> bool:
        """Check if file exists on SharePoint."""
        escaped_path = "/".join(quote(p) for p in remote_file_path.split("/"))
        url = self._url(f"root:/{escaped_path}", drive_id=drive_id)
        try:
            response = self.session.get(url, headers=self.auth.get_headers(), timeout=30)
            return response.status_code == 200
        except Exception as e:
            logger.warning("check_file_exists encountered error for %s: %s", remote_file_path, e)
            return False

    def upload_file(self, local_path: Path, remote_file_path: str, drive_id: str = "") -> dict:
        """Upload file to SharePoint. Uses chunked upload session for files > 4MB."""
        file_size = local_path.stat().st_size
        logger.info("Uploading file to SharePoint: %s -> %s (%d bytes) ...", local_path.name, remote_file_path, file_size)
        if not local_path.exists():
            raise FileNotFoundError(f"Local file not found to upload: {local_path}")
            
        escaped_path = "/".join(quote(p) for p in remote_file_path.split("/"))
        
        # Simple PUT for files <= 4MB
        if file_size <= 4 * 1024 * 1024:
            url = self._url(f"root:/{escaped_path}:/content", drive_id=drive_id)
            headers = {k: v for k, v in self.auth.get_headers().items() if k != "Content-Type"}
            with open(local_path, "rb") as f:
                response = self.session.put(url, headers=headers, data=f, timeout=60)
            response.raise_for_status()
            logger.info("[OK] Simple upload complete: %s", remote_file_path)
            return response.json()
            
        # Chunked upload session for files > 4MB
        logger.info("File size > 4MB. Creating chunked upload session...")
        create_session_url = self._url(f"root:/{escaped_path}:/createUploadSession", drive_id=drive_id)
        
        # createUploadSession tra 409/423 khi item dang bi lock/checkout
        # -> mot lan loi la ca run mat buoc upload (su co 2026-10-19).
        for attempt in range(1, 4):
            session_response = self.session.post(
                create_session_url, headers=self.auth.get_headers(), json={}, timeout=30
            )
            if session_response.status_code not in (409, 423):
                break
            if attempt < 3:
                logger.warning(
                    "[SharePointClient] createUploadSession %d, thu lai %d/3 sau %ds",
                    session_response.status_code, attempt, 15 * attempt,
                )
                time.sleep(15 * attempt)
        session_response.raise_for_status()
        
        upload_url = session_response.json().get("uploadUrl")
        if not upload_url:
            raise ValueError("Failed to obtain uploadUrl from createUploadSession.")
            
        # Chunk size must be a multiple of 320 KB (327680 bytes)
        # We use 320 KB * 10 = 3,276,800 bytes (~3.1 MB) chunks
        chunk_size = 3276800
        
        with open(local_path, "rb") as f:
            start = 0
            while start < file_size:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                chunk_len = len(chunk)
                end = start + chunk_len - 1
                
                headers = {
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                    "Content-Length": str(chunk_len)
                }
                
                # Chunk upload with retry for network resilience
                chunk_uploaded = False
                last_chunk_err = None
                for c_attempt in range(1, 4):
                    try:
                        response = self.session.put(upload_url, headers=headers, data=chunk, timeout=60)
                        response.raise_for_status()
                        chunk_uploaded = True
                        break
                    except (requests.exceptions.RequestException, ssl.SSLError) as err:
                        last_chunk_err = err
                        if c_attempt < 3:
                            time.sleep(2 * c_attempt)
                        else:
                            raise last_chunk_err
                
                start += chunk_len
                logger.info("  Uploaded chunk %d-%d / %d bytes", start - chunk_len, end, file_size)
                
        logger.info("[OK] Chunked upload complete: %s", remote_file_path)
        return response.json()
