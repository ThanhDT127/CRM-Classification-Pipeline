import os
import json
import logging
import requests
import msal
from pathlib import Path
from urllib.parse import quote

import config

logger = logging.getLogger("crm-automation")

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
    def __init__(self, auth: AuthProvider) -> None:
        self.auth = auth
        self.drive_id = config.SHAREPOINT_DRIVE_ID
        self.session = requests.Session()

    def _url(self, path: str = "") -> str:
        return f"{config.GRAPH_BASE}/drives/{self.drive_id}/{path}".rstrip("/")

    def download_file(self, remote_file_path: str, local_path: Path) -> Path:
        """Download file by its path relative to the drive root."""
        logger.info("Downloading file from SharePoint: %s ...", remote_file_path)
        # Url encode path segment
        escaped_path = "/".join(quote(p) for p in remote_file_path.split("/"))
        url = self._url(f"root:/{escaped_path}")
        
        # Get metadata to get download URL
        response = self.session.get(url, headers=self.auth.get_headers())
        response.raise_for_status()
        
        dl_url = response.json().get("@microsoft.graph.downloadUrl")
        if not dl_url:
            raise ValueError(f"No download URL found for remote path: {remote_file_path}")
            
        local_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Stream the download
        dl_response = self.session.get(dl_url, stream=True)
        dl_response.raise_for_status()
        
        with open(local_path, "wb") as f:
            for chunk in dl_response.iter_content(chunk_size=8192):
                f.write(chunk)
                
        logger.info("[OK] Download complete: %s (%d bytes)", local_path.name, local_path.stat().st_size)
        return local_path

    def upload_file(self, local_path: Path, remote_file_path: str) -> dict:
        """Upload file to SharePoint overwriting the file at remote_file_path."""
        logger.info("Uploading file to SharePoint: %s -> %s ...", local_path.name, remote_file_path)
        if not local_path.exists():
            raise FileNotFoundError(f"Local file not found to upload: {local_path}")
            
        escaped_path = "/".join(quote(p) for p in remote_file_path.split("/"))
        url = self._url(f"root:/{escaped_path}:/content")
        
        # Binary put request: headers must exclude Content-Type JSON
        headers = {k: v for k, v in self.auth.get_headers().items() if k != "Content-Type"}
        
        with open(local_path, "rb") as f:
            response = self.session.put(url, headers=headers, data=f)
            
        response.raise_for_status()
        logger.info("[OK] Upload complete: %s", remote_file_path)
        return response.json()
