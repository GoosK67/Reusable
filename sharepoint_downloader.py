from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
from office365.sharepoint.folders.folder import Folder
from pathlib import Path

SITE_URL = "https://<tenant>.sharepoint.com/sites/ProductManagement"
LIBRARY = "Product Management Library"

def connect():
    auth = AuthenticationContext(SITE_URL)
    auth.acquire_token_interactive()   # MFA login popup
    return ClientContext(SITE_URL, auth)

def download_recursive(ctx, folder: Folder, local_root: Path):
    folder = folder.expand(["Folders", "Files"]).get().execute_query()

    # Create local folder
    local_path = local_root / folder.serverRelativeUrl.replace("/", "_")
    local_path.mkdir(parents=True, exist_ok=True)

    # Download .docx files
    for f in folder.files:  # type: File
        if f.name.lower().endswith(".docx"):
            out = local_path / f.name
            print("Downloading:", f.name)
            with open(out, "wb") as fh:
                f.download(fh).execute_query()

    # Recurse
    for sf in folder.folders:
        download_recursive(ctx, sf, local_root)

def fetch_all_sd_files():
    ctx = connect()

    root_folder = ctx.web.get_folder_by_server_relative_url(LIBRARY)
    download_recursive(ctx, root_folder, Path("input"))

if __name__ == "__main__":
    fetch_all_sd_files()
