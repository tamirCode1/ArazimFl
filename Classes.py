import re
import pandas as pd
from flask_wtf import FlaskForm
from wtforms import StringField, IntegerField,PasswordField, SubmitField, HiddenField, SelectField, TextAreaField, DateField
from wtforms.validators import DataRequired, Length, Email, Optional
from flask_wtf.file import FileField, FileAllowed, FileRequired
import numpy as np
from Tools import get_file_from_drive, reload_dataframe_to_drive

class ARAZIM:
    def __init__(self, file_id):
        self._file_id = file_id
        self._file_df = None
        try:
            self._file_df = get_file_from_drive(file_id)
            print("File loaded successfully!")
        except FileNotFoundError:
            print(f"File not found.")
        except Exception as e:
            print(f"An error occurred: {e}")

    def set_df(self, df):
        self._file_df = df
        self.reload()

    def get_df(self):
        return self._file_df

    def print_rows(self):
        print(self._file_df)  # display first few rows

    def search_fields(self, colum_name, search_text):
        escaped_text = re.escape(search_text)
        filtered_df = self._file_df[
            self._file_df[colum_name].fillna('').astype(str).str.contains(escaped_text, na=False)
        ]

        column_values = filtered_df[colum_name].tolist()
        return column_values

    def open_ticket(self, column_name, ticket_name):
        print(ticket_name)
        tickets = []
        if isinstance(ticket_name, list):
            filtered_df_hold = []
            ticket_name = list(set(ticket_name))
            for item in ticket_name:
                filtered_df = self._file_df[
                    self._file_df[column_name].fillna('').astype(str) == str(item)
                    ]
                filtered_df_hold.append(filtered_df)
            if filtered_df_hold:
                filtered_df = pd.concat(filtered_df_hold, ignore_index=True)
            else:
                filtered_df = pd.DataFrame(columns=self._file_df.columns)
        else:
            filtered_df = self._file_df[
                self._file_df[column_name].fillna('').astype(str) == str(ticket_name)
                ]
        print(filtered_df, "TYPE: ", type(filtered_df))
        # המרה של כל הערכים למחרוזת
        for index, row in filtered_df.iterrows():
            row_dict = {col: "---" if pd.isna(val) else str(val) for col, val in row.items()}
            tickets.append(row_dict)
        print(tickets)
        return tickets

    def add_row(self, new_row):
        df = pd.concat([self._file_df, pd.DataFrame([new_row])], ignore_index=True)
        self._file_df = df
        self.reload()

    def delete_row(self, row_dict):
        df = self._file_df.copy()
        # בודקים כל עמודה
        mask_list =[]
        for col, val in row_dict.items():
            if col in df.columns:
                mask = values_match(df[col], val)
                mask_list.append(mask)

        final_mask = np.logical_and.reduce(mask_list)
        matching_rows = df[final_mask]
        print("שורות שנמצאו למחיקה:")
        print(matching_rows)

        if matching_rows.empty:
            print("DIDN'T FIND ANYTHING")
            return False

        # מוחק את השורות שנמצאו
        df = df.drop(matching_rows.index)

        # עדכון פנימי
        self._file_df = df.reset_index(drop=True)
        self.reload()

        # כתיבה חזרה ל-Excel
        print(f"✅ נמחקו {len(matching_rows)} שורות")
        return True


    def reload(self):
        reload_dataframe_to_drive(self._file_df, self._file_id, "arazim2025.xlsx")


class AddVendorForm(FlaskForm):
    vendor = StringField('ספק', validators=[DataRequired(), Length(max=100)])
    email = StringField('מייל', validators=[DataRequired(), Email(), Length(max=100)])
    phone_a = StringField('טלפון ראשי', validators=[Optional(), Length(max=20)])
    phone_b = StringField('טלפון משני', validators=[Optional(), Length(max=20)])
    cell_phone = StringField('נייד', validators=[Optional(), Length(max=20)])
    fax = StringField('פקס', validators=[Optional(), Length(max=20)])
    fname = StringField('שם פרטי', validators=[Optional(), Length(max=50)])
    lname = StringField('שם משפחה', validators=[Optional(), Length(max=50)])
    role = StringField('תפקיד', validators=[Optional(), Length(max=50)])
    summary = StringField('בתמצית', validators=[Optional(), Length(max=100)])
    on_post = StringField('איך מופיע באתר הדואר', validators=[Optional(), Length(max=100)])
    on_msg = StringField('איך מופיע בהודעת צד ג', validators=[Optional(), Length(max=100)])
    bank_account = IntegerField('מספר חשבון', validators=[Optional()])
    bank_branch = IntegerField('מספר סניף', validators=[Optional()])
    bank_number = IntegerField('מספר בנק', validators=[Optional()])


class LoginForm(FlaskForm):
    password = PasswordField('Password', validators=[DataRequired(), Length(min=6, max=50)])
    submit = SubmitField('כניסה')


class SearchForm(FlaskForm):
    query = StringField('Query', validators=[Optional(), Length(max=100)])
    filename = HiddenField('Filename', validators=[DataRequired()])
    submit = SubmitField('חפש')

class PrepareEmailForm(FlaskForm):
    chamber = SelectField('בחר לשכה:', validators=[DataRequired()], choices=["בחר לשכה"])
    date = DateField('תאריך דד ליין', validators=[DataRequired()], render_kw={"placeholder": "dd/mm/yyyy"})
    amount = StringField('סכום לגביה', render_kw={"placeholder": "0"})
    name = StringField('')
    account = StringField('')
    submit = SubmitField('הצג מייל')

class EmailForm(FlaskForm):
    subject = StringField('נושא', validators=[DataRequired()])
    body_html = TextAreaField('גוף ההודעה', validators=[DataRequired()])
    submit = SubmitField('שלח מייל')

class DummyForm(FlaskForm):
    pass

class UploadXLSXForm(FlaskForm):
    file = FileField("בחר קובץ XLSX", validators=[
        FileRequired(),
        FileAllowed(['xlsx'], 'רק קבצי XLSX מותרים!')
    ])
    submit = SubmitField("העלה")


def values_match(col_series, val):
    empty_values = [None, "", "---"]
    val = normalize_value(val)
    if isinstance(col_series, (pd.Series, pd.DataFrame)):
        col_series = col_series.apply(normalize_value)
        if val is None:  # מקרה מיוחד להשוואה ל־None
            return col_series.isna()
        return col_series == val
    else:
        col_series = normalize_value(col_series)
        if val in empty_values:
            val = None
        return val == col_series

def normalize_value(val):
    if pd.isna(val) or val in ["", "---", None]:
        return None
    return str(val)
