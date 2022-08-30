import datetime
import pathlib
import re

from babel.dates import format_datetime
from pathlib import Path
from docxtpl import DocxTemplate
from num2words import num2words
import pandas as pd
from tkinter import *
from tkinter import filedialog
import os

base_dir = Path(__file__).resolve().parents[1]
desktop = pathlib.Path.home() / 'Desktop'

win = Tk()

win.withdraw()

def open_file_xls():
    file = filedialog.askopenfile(mode='r', initialdir=desktop, title="Вибір екселя", filetypes=[('Виберіть ексель', '*.xlsx')])
    if file:
        filepath = os.path.abspath(file.name)
        return filepath
def open_doc():
    file = filedialog.askopenfile(mode='r', title="Вибір прикладу документа", filetypes=[('Виберіть що генерувати', '*.docx')])
    if file:
        filepath_doc = os.path.abspath(file.name)
        return filepath_doc


pth_xls = open_file_xls()
pth_doc = open_doc()
win.destroy()
win.mainloop()


# Nachalo auto zarach

word_template = pth_doc
excel_path = pth_xls

#utput_dir = base_dir / "Вихід документів"

#output_dir.mkdir(exist_ok=True)
df = pd.read_excel(excel_path, dtype={"Реєстраційни номер":str}, sheet_name = "Лист1")
df["kod1"] = pd.Index(df["Назва групи"])
df["nomer"] = pd.Index(df["Реєстраційни номер"])
df["name1"] = pd.Index(df["Прізвище"])
df["name2"] = pd.Index(df["Ім'я"])
df["name3"] = pd.Index(df["По батькові"])
df["adresa"] = pd.Index(df["Адреса"])
df["mob_number"] = pd.Index(df["Телефон"])
df["form_b"] = pd.Index(df["Бютжет чи контракт"])
df["gr_num"] = pd.Index(df["Номер групи"])
df["stupen"] = pd.Index(df["Освітній ступінь"])
df["spc"] = pd.Index(df["Спеціальність"])
df["seria_pass"] = pd.Index(df["Номер паспорту"])
df["vydan"] = pd.Index(df["Ким виданий"])
df["nakaz"] = pd.Index(df["Наказ"])
df["ser_sv"] = pd.Index(df["Серія свідоцтва"])
df["num_sv"] = pd.Index(df["Номер свідоцтва"])
df["kym_vydany"] = pd.Index(df["Хто видав свідоцтво"])
df["zno_num"] = pd.Index(df["Номер зно"])
df["zno_rik"] = pd.Index(df["Рік зно"])
df["forma_nav"] = pd.Index(df["Форма навчання"])
df["baly_ukr"] = pd.Index(df["ЗНО.Укр.мов і літ."])
def f(row):
       return num2words(row["ЗНО.Укр.мов і літ."])
df["baly_ukr_slova"] = df["ЗНО.Укр.мов і літ."].mask(df["ЗНО.Укр.мов і літ."].isna(), df["ЗНО Історія України"]).apply(num2words, lang='uk')
df["baly_mat"] = pd.Index(df["ЗНО Матем."])
def f(row):
    return num2words(row["ЗНО Матем."])
df["baly_mat_slova"] = df["ЗНО Матем."].apply(num2words, lang='uk')
df["baly_istor"] = pd.Index(df["ЗНО Історія України"])
def f(row):
    return num2words(row["ЗНО Історія України"])
df["baly_istor_slova"] = df["ЗНО Історія України"].apply(num2words, lang='uk')
#df["reg_cof"] = pd.Index(df["Регіональний коефіцієнт"])
#df["baly_reg_cof_slova"] = num2words(df["reg_cof"],  lang='uk')
#df["galuz_cof"] = pd.Index(df["Галузевий коефіцієнт"])
#df["baly_galuz_cof_slova"] = num2words(df["galuz_cof"],  lang='uk')
#df["sils_cof"] = pd.Index(df["Сільський коефіцієнт"])
#df["baly_sils_cof_slova"] = num2words(df["sils_cof"],  lang='uk')
#df["dod_bal"] = pd.Index(df["Додаткові бали за успішне закінчення підготовчих курсів"])
#df["dod_bal_slova"] = num2words(df["dod_bal"],  lang='uk')


df["data_sv"] = pd.to_datetime(df["Дата видачі свідотства"], errors='coerce').dt.strftime('%d.%m.%Y')
df["zayava_vid"] = pd.to_datetime(df["Дата подачі заяви"], errors='coerce').dt.strftime('%d.%m.%Y')
df["data"] = pd.to_datetime(df["Дата видачі"], errors='coerce').dt.strftime('%d.%m.%Y')
df["data_vstup"] = pd.to_datetime(df["Дата вступу"], errors='coerce').dt.strftime('%d.%m.%Y')
df["data_nakaz"] = pd.to_datetime(df["Дата наказу"], errors='coerce').dt.strftime('%d.%m.%Y')
df["d"] = datetime.datetime.today().strftime("%d")
df["m"] = format_datetime(datetime.datetime.today(), "MMMM", locale='uk_UA')
df["Y"] = datetime.datetime.today().strftime("%Y")


for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template)
    doc.render(record)
    output_dir = base_dir / f"{record['spc']}"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"{record['name1']}-повідомлення.docx"
    doc.save(output_path)
    #os.startfile(f"{output_path}", "print")

    print(output_path)