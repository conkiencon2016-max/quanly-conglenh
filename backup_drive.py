import os
import json
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

DB_FILE = "conglenh.db"

creds_json = os.environ["GOOGLE_CREDENTIALS_JSON"]
creds_dict = json.loads(creds_json)

creds = service_account.Credentials.from_service_account_info(
    creds_dict,
    scopes=["https://www.googleapis.com/auth/drive.file"]
)

service = build("drive", "v3", credentials=creds)

today = datetime.now().strftime("%Y%m%d_%H%M")

file_metadata = {
"name": f"conglenh_backup_{today}.db"
}

media = MediaFileUpload(DB_FILE, mimetype="application/octet-stream")

file = service.files().create(
body=file_metadata,
media_body=media,
fields="id"
).execute()

print("Backup Google Drive thành công:", file.get("id"))
