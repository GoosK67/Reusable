import msal
import requests
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from pathlib import Path

SITE_URL = "https://<tenant>.sharepoint.com/sites/ProductManagement"
LIBRARY = "Product Management Library"

CLIENT_ID = "04f0c124-f2bc-4f59-9d9b-89fddf3f8f32"  # Microsoft’s official public client (no app reg needed)
SCOPES = ["https://<tenant>.sharepoint.com/.default"]   # Replace with your tenant

def acquire_token():
    app = msal.PublicClientApplication(
        client_id=CLIENT_ID,
        authority=f"https://login.microsoftonline.com/common"
    )

    # Try silent login first
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result:
            return result["access_token"]

    # Force interactive login (browser popup)
    result = app.acquire_token_interactive(SCOPES)
    return result["access_token"]

def build_context():
    token = acquire_token()
    ctx = ClientContext(SITE_URL).with_access_token(token)
    return ctx

def download_recursive(ctx, folder, local_root):
    folder = folder.expand(["Files", "Folders"]).get().execute_query()

    local_path = local_root / folder.serverRelativeUrl.replace("/", "_")
    local_path.mkdir(parents=True, exist_ok=True)

    for f in folder.files:
        if f.name.lower().endswith(".docx"):
            print("Downloading:", f.name)
            with open(local_path / f.name, "wb") as fh:
                f.download(fh).execute_query()

    for sub in folder.folders:
        download_recursive(ctx, sub, local_root)

def fetch_all_sd_files():
    ctx = build_context()
    root = ctx.web.get_folder_by_server_relative_url(LIBRARY)
    download_recursive(ctx, root, Path("input"))

if __name__ == "__main__":
    fetch_all_sd_files()
