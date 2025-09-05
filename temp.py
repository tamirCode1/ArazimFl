from Tools import encrypt, json

def open_config():
    with open("static/config.json", 'r', encoding='utf-8') as file:
        data = json.load(file)
        main_path = data["path"]
        waiting_path = data["waiting_path"]
        manager = data["manager_pass"]
        user = data["user_pass"]
        mail_sub = data["mail_subject"]
        mail_body = data["mail_body"]
        secret_key = data["secret_key"]
        return main_path, waiting_path, manager, user, mail_sub, mail_body, secret_key

