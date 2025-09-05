from flask import Flask
from flask_wtf.csrf import CSRFProtect
app = Flask(__name__)
from routs import *
from Tools import open_config, send_email, backup, os, get_file_from_drive
from datetime import timedelta, datetime
import schedule
import time
import threading


app.config["SECRET_KEY"] = open_config()[-1]
csrf = CSRFProtect(app)
app.permanent_session_lifetime = timedelta(minutes=30)

schedule.every(1).days.at("19:00").do(backup)  # לדוגמה בשעה 10:00
def run_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)

def main():
  scheduler_thread = threading.Thread(target=run_scheduler)
  scheduler_thread.daemon = True  # כדי שה-thread ייסגר יחד עם התוכנית הראשית
  scheduler_thread.start()
  port = int(os.environ.get("PORT", 5000))
  app.run(host="0.0.0.0", port=port)


if __name__ == '__main__':
    main()


