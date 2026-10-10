"""SharePoint list-item attachments for an unattended task (no browser, no person).

SharePoint REST rejects an app-only token backed by a client SECRET
("Unsupported app only token"); it accepts one backed by a CERTIFICATE. So this
module gets the SharePoint token with the app's certificate (same Azure app,
same tenant; the public .cer is uploaded once under the app registration's
Certificates & secrets) and downloads an item's attachment into a folder the
caller owns and deletes.

    from sp_attachments import sp_cert_session, list_attachments, download_attachment
    sp = sp_cert_session()                       # cert from %LOCALAPPDATA%\\uhc_reviewer\\sp_app_cert.pem
    names = list_attachments(sp, item_id)        # [{'name': ..., 'url': ...}]
    path = download_attachment(sp, item_id, names[0]['name'], dest_dir)

Nothing here prints file names or content. The cert path can be overridden with
env UHC_SP_CERT_PEM. NOT wired into any job; a one-record test comes first.
"""
from __future__ import annotations

import hashlib
import importlib.util
import os
import sys
from pathlib import Path

import msal
import requests
from cryptography import x509
from cryptography.hazmat.primitives import serialization

_HERE = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("uhc_pipeline", _HERE / "uhc_pipeline.py")
P = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(P)

CERT_PEM = Path(os.environ.get("UHC_SP_CERT_PEM")
                or Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "uhc_reviewer" / "sp_app_cert.pem")


def _cert_credential(pem_path: Path = CERT_PEM) -> dict:
    """MSAL client_credential dict from a PEM holding the private key + cert."""
    pem = pem_path.read_bytes()
    cert = x509.load_pem_x509_certificate(pem)
    thumb = hashlib.sha1(cert.public_bytes(serialization.Encoding.DER)).hexdigest()
    return {"private_key": pem.decode(), "thumbprint": thumb}


def sp_cert_token(scope: str = f"https://{P.SP_HOSTNAME}/.default") -> str:
    creds = P.load_azure_creds()
    app = msal.ConfidentialClientApplication(
        client_id=creds["client_id"],
        authority=f"https://login.microsoftonline.com/{creds['tenant_id']}",
        client_credential=_cert_credential())
    result = app.acquire_token_for_client(scopes=[scope])
    if "access_token" not in result:
        raise RuntimeError(f"certificate token failed: {result.get('error')}: {result.get('error_description', '')[:200]}")
    return result["access_token"]


def sp_cert_session() -> requests.Session:
    return P.sp_session(sp_cert_token())


def list_attachments(sp: requests.Session, item_id, list_id: str = P.SP_LIST_ID) -> list[dict]:
    r = sp.get(f"{P.SP_SITE_URL}/_api/web/lists(guid'{list_id}')/items({int(item_id)})/AttachmentFiles"
               "?$select=FileName,ServerRelativeUrl")
    r.raise_for_status()
    out = []
    for a in (r.json().get("d") or {}).get("results") or []:
        out.append({"name": a["FileName"], "url": a["ServerRelativeUrl"]})
    return out


def download_attachment(sp: requests.Session, item_id, file_name: str, dest_dir: str,
                        list_id: str = P.SP_LIST_ID, as_name: str | None = None) -> str:
    """Stream one attachment to <dest_dir>/<as_name or file_name>. Returns the path.
    The caller owns the folder and deletes the file when done."""
    url = (f"{P.SP_SITE_URL}/_api/web/lists(guid'{list_id}')/items({int(item_id)})"
           f"/AttachmentFiles('{file_name.replace(chr(39), chr(39) * 2)}')/$value")
    dest = Path(dest_dir); dest.mkdir(parents=True, exist_ok=True)
    out = dest / (as_name or file_name)
    with sp.get(url, stream=True, headers={"Accept": "application/octet-stream"}) as r:
        r.raise_for_status()
        with open(out, "wb") as fh:
            for chunk in r.iter_content(1 << 16):
                fh.write(chunk)
    return str(out)


if __name__ == "__main__":
    # Connectivity check only: token + a metadata call. Prints status codes, never names.
    tok = sp_cert_token()
    sp = P.sp_session(tok)
    r = sp.get(f"{P.SP_SITE_URL}/_api/web/lists(guid'{P.SP_LIST_ID}')?$select=Title")
    print("certificate token: OK | SP REST list read:", r.status_code)
    sys.exit(0 if r.ok else 1)
