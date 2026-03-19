import os

BACKUP_DIR = "backups"

def upload_all():

    if not os.path.exists(BACKUP_DIR):
        print("Không có thư mục backup")
        return

    files = os.listdir(BACKUP_DIR)

    if not files:
        print("Không có file để upload")
        return

    for f in files:
        path = os.path.join(BACKUP_DIR, f)

        if os.path.isfile(path):
            try:
                upload_file(path)
            except Exception as e:
                print("Upload lỗi:", e)

if __name__ == "__main__":
    upload_all()
