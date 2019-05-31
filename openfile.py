import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import workbook

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename(title="Открыть файл генератора", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
# show an "Open" dialog box and return the path to the selected file

wb = workbook
f_open = False
# Проверяем, закрыт ли файл генератора
Excel = win32com.client.Dispatch("Excel.Application")

wb32 = Excel.Workbooks.Open(filename)
print(wb32)
wb32.Close()

try:
    my_file = open(filename, "r+")  # or "a+", whatever you need
    my_file.close()
    if '.xlsx' in filename:
        wb = load_workbook(filename=filename)
        f_open = True
except IOError:
    print("Нельзя работать с открытым файлом!!!")
