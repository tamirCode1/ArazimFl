import json
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date
from email.mime.base import MIMEBase
from email import encoders
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import io
import tempfile

creds_json = os.getenv("GOOGLE_CREDENTIALS")
CREDENTS_DICT = json.loads(creds_json)  # הנתיב לקובץ שלך
SCOPES = ['https://www.googleapis.com/auth/drive']
CREDENTIALS = service_account.Credentials.from_service_account_info(CREDENTS_DICT, scopes=SCOPES)
DRIVE_SERVICE = build('drive', 'v3', credentials=CREDENTIALS)

def get_file_from_drive(file_id):
    # צור נתיב זמני בלי לפתוח את הקובץ מראש
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)  # לסגור את ה-file descriptor מיידית

    try:
        # בקשת הורדת הקובץ
        request = DRIVE_SERVICE.files().get_media(fileId=file_id)
        with io.FileIO(temp_path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    print(f"Downloading {int(status.progress() * 100)}%")

        print(f"קובץ הורד בהצלחה ל: {temp_path}")

        # טען את הקובץ ל-DataFrame
        df = pd.read_excel(temp_path)
        print("DataFrame נטען בהצלחה!")

        return df

    except Exception as e:
        print(f"שגיאה בהורדה או טעינה: {e}")
        return None

    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
                print("הקובץ הזמני נמחק.")
            except PermissionError:
                print("הקובץ עדיין נעול, נדלג על מחיקה.")


def reload_dataframe_to_drive(df, file_id, filename):
    # צור נתיב זמני בלי לפתוח את הקובץ מראש
    fd, temp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)

    try:
        df.to_excel(temp_path, index=False)

        media = MediaFileUpload(
            temp_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        updated_file = DRIVE_SERVICE.files().update(
            fileId=file_id,
            media_body=media,
            body={"name": filename}
        ).execute()

        print(f"הקובץ עודכן בהצלחה: {updated_file.get('id')}")
        return updated_file

    except Exception as e:
        print(f"שגיאה בהעלאה: {e}")
        return None

    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except PermissionError:
                print("הקובץ עדיין נעול, נדלג על מחיקה.")

ALLOWED_EXTENSIONS = "/tmp", {"xlsx"}

def open_config():
    with open("static/config.json", 'r', encoding='utf-8') as file:
        data = json.load(file)
        main_id = data["id"]
        waiting_path = data["waiting_path"]
        manager = data["manager_pass"]
        user = data["user_pass"]
        sender_mail = data["sender_mail"]
        sender_pass = data["sender_pass"]
        fields = data["fields"]
        secret_key = data["secret_key"]
        return main_id, waiting_path, manager, user, sender_mail, sender_pass, fields, secret_key


def waiting_list_cards(path):
    try:
        waiting_df = pd.read_excel(path, engine='openpyxl')

        tickets = []
        for index, row in waiting_df.iterrows():
            tickets.append(dict(row))
        print(tickets)
        print("File loaded successfully!")
    except FileNotFoundError:
        print(f"File '{path}' not found.")
        tickets = ["שגיאה - לא הצלחנו למצוא את הקובץ"]
    except Exception as e:
        print(f"An error occurred: {e}")
        tickets = ["שגיאה."]

    print(tickets)
    return tickets


def handle_card_action(card, approved, waiting_path, arz):
    try:
        waiting_df = pd.read_excel(waiting_path, engine='openpyxl')
        key_fields = ['ספק', 'מייל', 'טלפון ראשי', 'נייד', 'שם פרטי', 'שם משפחה', 'טלפון 2',  'פקס', "איך מופיע בתמצית", "איך מופיע באתר הדואר", "איך מופיע בהודעת צד ג'", "חשבון בנק", "מס סניף", "מס בנק"]

        def row_matches(row):
            return all(
                (pd.isna(card.get(f)) and pd.isna(row[f])) or
                (str(card.get(f)) == str(row[f]))
                for f in key_fields
            )

        matching_rows = waiting_df[waiting_df.apply(row_matches, axis=1)]

        if matching_rows.empty:
            print("לא נמצאה שורה תואמת למחיקה/העברה.")
            return

        if approved == "t":
            arz.add_row(card)

        waiting_df.drop(matching_rows.index).to_excel(waiting_path, index=False)
        print("השורה הוסרה מקובץ ההמתנה")

    except Exception as e:
        print(f"שגיאה: {e}")


def add_to_waiting_list(card, waiting_path):
    df = pd.read_excel(waiting_path, engine='openpyxl')
    df = pd.concat([df, pd.DataFrame([card])], ignore_index=True)
    df.to_excel(waiting_path, index=False, engine='openpyxl')


def to_str(val):
    return "" if val is None else str(val)

def send_email(sender_email, sender_password, recipient_email, subject, body_html):
    # בונים את ההודעה
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email

    # חלק HTML עם כתיבה מימין לשמאל
    html_part = f"""
    <html dir="rtl" lang="he">
    <body style="font-family: Arial; direction: rtl; text-align: right;">
    {body_html}
    </body>
    </html>
    """
    # מוסיפים את החלק ל־MIME
    msg.attach(MIMEText(html_part, "html", "utf-8"))

    attachment_path = "static/arazim2024.xlsx"
    if not os.path.exists(attachment_path):
        raise FileNotFoundError(f"Attachment not found: {attachment_path}")

    with open(attachment_path, "rb") as f:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        part.set_payload(f.read())

    encoders.encode_base64(part)
    filename = os.path.basename(attachment_path)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    part.add_header("Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    msg.attach(part)

    # מציגים את ההודעה לפני שליחה
    print("----- PREVIEW EMAIL -----")
    print("To:", recipient_email)
    print("Subject:", subject)
    print("Body (HTML):")
    print(html_part)
    print("-------------------------")


    # שליחה דרך Gmail SMTP (למשל)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())

    print("Email sent successfully!")


def send_mail_with_xlsx(
    to_address: str,
    subject: str,
    html_body: str,
    attachment_path: str = "static/arazim2024.xlsx",
    smtp_host: str = "smtp.office365.com",  # ל-Office 365: smtp.office365.com
    smtp_port: int = 587,                   # STARTTLS
    smtp_user: str = "you@example.com",
    smtp_password: str = "YOUR_SMTP_PASSWORD",
    from_address: str | None = None,
):
    """
    שולח מייל עם קובץ XLSX מצורף באמצעות SMTP (STARTTLS).
    html_body — טקסט HTML. הוסף dir="rtl" לכיוון עברי.
    """

    if from_address is None:
        from_address = smtp_user

    # בניית המעטפה
    msg = MIMEMultipart()
    msg["From"] = from_address
    msg["To"] = to_address
    msg["Subject"] = subject

    # גוף HTML (RTL)
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # צירוף הקובץ
    if not os.path.exists(attachment_path):
        raise FileNotFoundError(f"Attachment not found: {attachment_path}")

    with open(attachment_path, "rb") as f:
        part = MIMEBase(
            "application",
            "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        part.set_payload(f.read())

    encoders.encode_base64(part)
    filename = os.path.basename(attachment_path)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    part.add_header("Content-Type",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    msg.attach(part)

    # שליחה דרך SMTP (STARTTLS)
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtp_user, smtp_password)
        server.sendmail(from_address, [to_address], msg.as_string())



def backup():
    try:
        backup_date = date.today()
        send_email(sender_email="tamir.f.e.11@gmail.com",sender_password= "ynyc xgpe akeg pkqf",recipient_email = "davidcoh@eca.gov.il", subject = f"גיבוי - ארזים - {backup_date}", body_html = "מערכת פעילה.")

    except Exception as e:
        print("Error in backup:", e)


