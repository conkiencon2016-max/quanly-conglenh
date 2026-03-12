import os, json
from google.oauth2 import service_account

creds_json = os.environ["GOOGLE_CREDENTIALS_JSON"]
creds_dict = json.loads(creds_json)

creds = service_account.Credentials.from_service_account_info(
    creds_dict,
    scopes=['https://www.googleapis.com/auth/drive.file']
)