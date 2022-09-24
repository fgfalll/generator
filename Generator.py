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
example = Path(__file__).resolve().parents[0]
exam_dir = example / "Приклади"
desktop = pathlib.Path.home() / 'Desktop'

win = Tk()

win.withdraw()
win.lift()
win.attributes("-topmost", True)

def open_file_xls():
    file = filedialog.askopenfile(mode='r', initialdir=desktop, title="Вибір екселя"
                                  , filetypes=[('Виберіть ексель', '*.xlsx')])
    if file:
        filepath = os.path.abspath(file.name)
        return filepath

pth_xls = open_file_xls()
win.destroy()
win.mainloop()
excel_path = pth_xls

#Vybor example
print("Який документ генерувати (введіть цифру)")
print("1. Анкета для банку")
print("2. Аркуш випробувань")
print("3. Витяг з наказу")
print("4. Опис справи")
print("5. Повідомлення")
print("6. Свій приклад документу")
what_exam = input("Ваш вибір:")
if what_exam == "1":
    pth_doc = exam_dir / "Анкета для банку.docx"
    what = "анкета для банку"
elif what_exam == "2":
    pth_doc = exam_dir / "Аркуш випробувань.docx"
    what = "аркуш випробувань"
elif what_exam == "3":
    pth_doc = exam_dir / "Витяг з наказу.docx"
    what = "витяг з наказу"
elif what_exam == "4":
    pth_doc = exam_dir / "Опис справи.docx"
    what = "опис справи"
elif what_exam == "5":
    pth_doc = exam_dir / "Повідомлення.docx"
    what = "повідомлення"
elif what_exam == "6":
    def open_doc():
        file = filedialog.askopenfile(mode='r', title="Вибір прикладу документа",
                                      filetypes=[('Виберіть свій приклад', '*.docx')])
        if file:
            filepath_doc = os.path.abspath(file.name)
            return filepath_doc
    pth_doc = open_doc()
    what = input("як назвати документ:")
else:
    print("ok")
    exit()
word_template = pth_doc

#User imput
print("Виберіть лист з якого печатать з переліку нижче")
sheet_names = pd.ExcelFile(excel_path).sheet_names
print(sheet_names)
sheet = input("Який лист використати:")
print('Чи потрібні бали?')
des = input('Введіть так або ні:\n')

df = pd.read_excel(excel_path, dtype={"Реєстраційни номер":str,"Номер зно":str,"Рік зно":str, "Конкурсний бал":str},
                   sheet_name = sheet).fillna(value=' ')

# Main for define start
try:
    df["kod1"] = pd.Index(df["Назва групи"])
    df["nomer"] = pd.Index(df["Реєстраційни номер"])
    df["name1"] = pd.Index(df["Прізвище"])
    df["name2"] = pd.Index(df["Ім'я"])
    df["name3"] = pd.Index(df["По батькові"])
    df["adresa"] = pd.Index(df["Адреса"])
    df["mob_number"] = pd.Index(df["Контактний номер"])
    df["form_b"] = pd.Index(df["Бютжет чи контракт"])
    df["gr_num"] = pd.Index(df["Номер групи"])
    df["stupen"] = pd.Index(df["Освітній ступінь"])
    df["spc"] = pd.Index(df["Спеціальність"])
    df["num_pass"] = pd.Index(df["ДПО.Номер"])
    df["seria_pass"] = pd.Index(df["ДПО.Серія"])
    df["vydan"] = pd.Index(df["ДПО.Ким виданий"])
    df["nakaz"] = pd.Index(df["Наказ про зарахування"])
    df["ser_sv"] = pd.Index(df["Серія документа"])
    df["num_sv"] = pd.Index(df["Номер документа"])
    df["kym_vydany"] = pd.Index(df["Ким видано"])
    df["zno_num"] = pd.Index(df["Номер зно"])
    df["zno_rik"] = pd.Index(df["Рік зно"])
    df["forma_nav"] = pd.Index(df["Форма навчання"])
    df["doc_of"] = pd.Index(df["ДПО"])
    df["typ_doc"] = pd.Index(df["Тип документа"])
    df["typ_doc_dod"] = pd.Index(df["Додаток до типу документу"])
    df["prot_num"] = pd.Index(df["Номер протоколу"])
    # Main define end

    # Date define
    df["data_sv"] = pd.to_datetime(df["Дата видачі документа"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["data_prot"] = pd.to_datetime(df["Дата протоколу"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["zayava_vid"] = pd.to_datetime(df["Дата подачі заяви"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["data"] = pd.to_datetime(df["ДПО.Дата видачі"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["data_vstup"] = pd.to_datetime(df["Дата вступу"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["data_nakaz"] = pd.to_datetime(df["Дата наказу"], errors='coerce').dt.strftime('%d.%m.%Y')
    df["d"] = datetime.datetime.today().strftime("%d")
    df["m"] = format_datetime(datetime.datetime.today(), "MMMM", locale='uk_UA')
    df["Y"] = datetime.datetime.today().strftime("%Y")
    # Date define end
except KeyError:
    pass

# Baly to words and define
if des == "так":
    df["baly_ukr"] = pd.Index(df["ЗНО.Українська мова та література"])
    df["baly_mat"] = pd.Index(df["ЗНО.Математика"])
    df["baly_istor"] = pd.Index(df["ЗНО.Історія України"])
    df["sr_bal"] = pd.Index(df["Конкурсний бал"])
    df["baly_ukr_mov"] = pd.Index(df["ЗНО.Українська мова"])
    df["baly_geo"] = pd.Index(df["ЗНО.Географія"])
    def f(row):
       return num2words(row["ЗНО.Українська мова та література"])
       return num2words(row["ЗНО.Математика"])
       return num2words(row["ЗНО.Історія України"])
       return num2words(row["ЗНО.Українська мова"])
       return num2words(row["ЗНО.Географія"])
       return num2words(row["Конкурсний бал"])

    df["baly_ukr_slova"] = df["ЗНО.Українська мова та література"].apply(num2words, lang='uk')
    df["baly_geo_slova"] = df["ЗНО.Географія"].apply(num2words, lang='uk')
    #df["baly_ukr_mov_slova"] = df["ЗНО.Українська мова"].apply(num2words, lang='uk')
    #df["baly_mat_slova"] = df["ЗНО.Математика"].apply(num2words, lang='uk')
    #df["baly_istor_slova"] = df["ЗНО.Історія України"].apply(num2words, lang='uk')
    df["sr_bal_slova"] = df["Конкурсний бал"].apply(num2words, lang='uk')
else:
    print('ok')
    pass
#df["baly_istor_slova"] = df["ЗНО Історія України"].apply(num2words, lang='uk')
#df["reg_cof"] = pd.Index(df["Регіональний коефіцієнт"])
#df["baly_reg_cof_slova"] = num2words(df["reg_cof"],  lang='uk')
#df["galuz_cof"] = pd.Index(df["Галузевий коефіцієнт"])
#df["baly_galuz_cof_slova"] = num2words(df["galuz_cof"],  lang='uk')
#df["sils_cof"] = pd.Index(df["Сільський коефіцієнт"])
#df["baly_sils_cof_slova"] = num2words(df["sils_cof"],  lang='uk')
#df["dod_bal"] = pd.Index(df["Додаткові бали за успішне закінчення підготовчих курсів"])
#df["dod_bal_slova"] = num2words(df["dod_bal"],  lang='uk')
# Baly to words and define end

#Output
for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template)
    doc.render(record)
    output_dir = base_dir / f"{record['spc']}"
    output_dir.mkdir(exist_ok=True)
    output_path = output_dir / f"{record['name1'],record['name2'],record['name3']}-{what}.docx"
    doc.save(output_path)
    #os.startfile(f"{output_path}", "print")


    print(output_path)