import pathlib

from sympy import symbols, Eq, solve
from tkinter import *
from tkinter import filedialog
import os
import pandas as pd

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

df = pd.read_excel(excel_path,sheet_name = "Лист1").fillna(value='1')

print(df["ЗНО.Українська мова та література"])
for index, row in df.iterrows():
    ukr = pd.to_numeric("ЗНО.Українська мова та література", errors='coerce')
    mat = pd.to_numeric("ЗНО.Математика", errors='coerce')
    istor = pd.to_numeric("ЗНО.Історія України", errors='coerce')
    balll = pd.to_numeric("Конкурсний бал", errors='coerce')
    x, y = symbols('x, y')
    eq1 = Eq(((0.3 * row[ukr] + 0.5 * row[mat] + 0.2 * row[istor]) * x), row[balll])
    sol1 = solve((eq1), (x))
    floats = [float(x) for x in sol1]
    for floats in range(100):
        s1 = (int(floats))
    print(s1)
    if s1 > 1.06:
        x = 1.04
        y = 1.02
    elif s1 < 1.06 >= 1.04:
        x = 1.04
        y = 1
    else:
        pass
    if s1 <= 1.02 >= 1:
        x = 1
        y = 1.02
    else:
        pass
d = {'galyzev':[x], 'reg':[y]}
df1 = pd.DataFrame(data=d)


