from flask import render_template, request, url_for, send_file, session, flash, jsonify, get_flashed_messages
from werkzeug.utils import redirect
from app import app
from Tools import *
from Classes import ARAZIM, AddVendorForm, LoginForm, SearchForm, PrepareEmailForm, EmailForm, DummyForm, UploadXLSXForm
import urllib.parse
import ast

FILE_ID, WAITING_P, MANAGER_P, USER_P, SENDER_MAIL, SENDER_PASS, FIELDS = open_config()[:-1]
FOLDER_PATH = "static/"
ARZ = ARAZIM(FILE_ID)
UPLOAD_FOLDER = "/tmp"

@app.route("/get_flash")
def get_flash():
    # מקבל את כל ההודעות שנשמרו ב-session
    messages = get_flashed_messages(with_categories=True)
    result = []
    for category, message in messages:
        result.append({
            "title": "הצלחה" if category=="success" else "שגיאה",
            "text": message,
            "icon": category
        })
    return jsonify(messages=result)


@app.route("/", methods=["GET", "POST"])
def login():
    form = LoginForm()
    if form.validate_on_submit():
        password = form.password.data
        if password == MANAGER_P:
            session.permanent = True
            session['role'] = 'manager'
            session["last_search"] = ""
            session["last_type"] = 'ספק'
            return redirect(url_for("home"))
        elif password == USER_P:
            session.permanent = True
            session['role'] = 'user'
            return redirect(url_for("user"))
    return render_template("login.html", form=form)


@app.route("/user", methods=["GET", "POST"])
def user():
    if 'role' not in session or session['role'] != 'user':
        return redirect(url_for('login'))

    form = AddVendorForm()

    if form.validate_on_submit():
        # שים לב – form.data זה מילון מסודר


        new_data = {FIELDS.get(k, k): v for k, v in form.data.items()}
        add_to_waiting_list(new_data, WAITING_P)
        return redirect(url_for('user'))
    return render_template("user_add.html", form=form)



@app.route("/home", methods=["GET", "POST"])
def home():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = SearchForm()

    if form.validate_on_submit():
        session["last_type"] = form.filename.data
        session["last_search"] = form.query.data
        return redirect(url_for("search"))

    return render_template("home.html", form=form)

@app.route('/search', methods=['GET', 'POST'])
def search():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = SearchForm()

    if form.validate_on_submit():
        print(form.data)
        session["last_type"] = form.filename.data
        session["last_search"] = form.query.data
        return redirect(url_for('search'))

    # במקרה של GET, או אם לא הוגש הטופס, נטען את הערכים מה-session
    last_type = session.get("last_type", "ספק")   # ברירת מחדל אם לא קיים
    last_search = session.get("last_search", "")

    # הרץ חיפוש לפי הערכים
    matching_rows = ARZ.search_fields(last_type, last_search)

    return render_template(
        'search.html',
        form=form,
        type=last_type,
        placeholder=last_search,
        items=matching_rows
    )

@app.route("/clear_search", methods=["GET"])
def clear_search():
    session["last_search"] = ""
    return redirect(url_for("search"))



@app.route('/select_item', methods=['POST'])
def select_item():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    cards = ARZ.open_ticket(request.form['colum'], request.form['item'])
    fields_order = list(FIELDS.values())[:9]
    print("CARD::: \n")
    print(cards)
    print("***********\n\n")
    return render_template('tickets.html', cards=cards, field_order = fields_order)


@app.route('/select_items', methods=['POST'])
def select_items():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    print( request.form.getlist('selected_items'))
    selected_items = request.form.getlist('selected_items')
    cards = ARZ.open_ticket(request.form['colum'], selected_items)
    fields_order = list(FIELDS.values())[:9]
    print(fields_order)
    return render_template('tickets.html', cards=cards, field_order = fields_order)



@app.route('/prepare_email', methods=['GET', 'POST'])
def prepare_email():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    email = request.args.get('email')
    name = request.args.get('name')

    form = PrepareEmailForm()
    chambers = {"נתניה": "247783", "תל אביב": "247759", "חיפה": "250338", "ירושלים": "250697", "טבריה": "250736", "צפת": "250825", "קרית שמונה": "250883", "עכו": "251009", "עפולה": "250728", "נצרת": "250744", "בית שאן": "250906", "חדרה": "248860", "קריות חיפה": "250980", "באר שבע": "248894", "קרית גת": "250752", "כפר סבא": "248797", "ראשון לציון": "248878", "פתח תקווה": "250168", "רחובות": "248886", "אשקלון": "250304", "אילת": "251041", "דימונה": "250998", "אשדוד": "250671", "הרצליה": "250388", "רמלה": "250370"}
    session["chambers"] = chambers
    form.chamber.choices = [(c, c) for c in chambers.keys()]

    if request.method == 'GET':
        form.name.data = name
        form.account.data = list(chambers.values())[0]

    if request.method == 'POST' and form.validate_on_submit():
        chamb = form.chamber.data
        the_date = form.date.data
        amount = form.amount.data
        name = form.name.data
        account_num = chambers[chamb]


        subject = f" הודעה מלשכת הוצלפ {chamb} בדבר קבלת כספים ללא אסמכתא"
        body_text = f""" {name} שלום רב 
            ביום: {the_date} התקבל סך של: {amount} ₪, לחשבון לשכת הוצלפ {chamb} מס' חשבון: {account_num}.
            הכספים התקבלו ללא אסמכתא מתאימה ובניגוד להנחיות רשות האכיפה והגבייה."""

        # מעבר למסך תצוגה מקדימה
        session["email"] = email
        session["subject"] = subject
        session["body"] = body_text
        return redirect(url_for("send_email_route"))

    return render_template('prepare_email.html', email=email, form=form)

@app.route("/send_email")
def send_email_route():
    try:
        # בניית מחרוזת mailto
        to = session.get("email")
        subject = session.get("subject")
        body =session.get("body")

        body = body.strip().replace('\r\n', '\n').replace('\r', '\n')

        # קידוד מתאים ל-URL
        subject_encoded = urllib.parse.quote(subject)
        body_encoded = urllib.parse.quote(body)

        mailto_url = f"mailto:{to}?subject={subject_encoded}&body={body_encoded}"

        flash("המייל נפתח בהצלחה! אם זה לא קרה, ודא ש-Outlook הוא ברירת המחדל לשליחת דוא\"ל", "success")
        return redirect(mailto_url)

    except Exception as e:
        flash(f"שגיאה בפתיחת המייל: {str(e)}", "error")
        return redirect(url_for("home"))


@app.route('/delete_card', methods=['GET', 'POST'])
def delete_card():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    card = request.args.get("card_dt")
    card = ast.literal_eval(card)
    reverse_fields = {v: k for k, v in FIELDS.items()}  # הפוך את המיפוי
    card = {reverse_fields.get(k, k): v for k, v in card.items()}
    ARZ.delete_row(card)
    return redirect(url_for("search"))

@app.route('/edit_card', methods=['GET', 'POST'])
def edit_card():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = AddVendorForm()

    card = request.args.get("card_dt")
    card = ast.literal_eval(card)

    if request.method == "GET":
        reverse_fields = {v: k for k, v in FIELDS.items()}  # הפוך את המיפוי
        card_form = {reverse_fields.get(k, k): v for k, v in card.items()}
        form = AddVendorForm(data=card_form)

    if form.validate_on_submit():
        ARZ.delete_row(card)
        new_data = {FIELDS.get(k, k): v for k, v in form.data.items()}
        ARZ.add_row(new_data)


        fields_order = list(FIELDS.values())[:9]
        return render_template('tickets.html', cards=[new_data], field_order=fields_order)

    return render_template("add.html", form=form, purpose="ערוך כרטיס (בטבלה)")





@app.route('/add', methods=['GET', 'POST'])
def add():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = AddVendorForm()

    if form.validate_on_submit():
        print("im here- stupid")
        # שים לב – form.data זה מילון מסודר
        new_data = {FIELDS.get(k, k): v for k, v in form.data.items()}
        ARZ.add_row(new_data)

        return redirect(url_for('add'))
    return render_template("add.html", form=form, purpose="הוסף")


@app.route('/waiting_list', methods=['GET','POST'])
def waiting_list():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = DummyForm()
    if request.method == 'POST':
        data = request.form.to_dict()
        approved = data.pop('approved')
        handle_card_action(data, approved, WAITING_P, ARZ)
        print("Card:", data)


    cards = waiting_list_cards(WAITING_P)
    fields_order = list(FIELDS.values())[:8]
    return render_template("waitings.html", cards = cards, form=form, field_order = fields_order)

@app.route('/download')
def download_file():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    try:
        df = ARZ.get_df()
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        df.to_excel(temp_file.name, index=False)

        # מחזירים למשתמש כקובץ להורדה
        return send_file(
            temp_file.name,
            as_attachment=True,
            download_name="arazim2024.xlsx",  # שם שיוצע להורדה
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        flash(f"שגיאה בהורדת הקובץ: {e}", "error")
        return redirect(url_for("home"))


os.makedirs(FOLDER_PATH, exist_ok=True)
@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    if 'role' not in session or session['role'] != 'manager':
        return redirect(url_for('login'))

    form = UploadXLSXForm()
    if request.method == "POST":
        if form.validate_on_submit():
            file = form.file.data  # זה ה־FileStorage של Flask

            try:
                df = pd.read_excel(file)
                ARZ.set_df(df)
                flash("הקובץ הועלה ונטען ל-DataFrame בהצלחה!", "success")
                print(df.head())

                return redirect(url_for("home"))

            except Exception as e:
                flash(f"שגיאה בקריאת הקובץ: {e}", "error")
        else:
            flash("הקובץ לא עבר אישור- שים לב שרק קבצי XLSX מאושרים!", "error")
    return render_template("files.html", form=form)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


