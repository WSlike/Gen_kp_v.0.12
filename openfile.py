import win32com.client
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import workbook
from transliterate import translit
from datetime import datetime, timedelta


utc3 = timedelta(hours=3)  # часовой пояс +3 часа

Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
filename = askopenfilename(title="Открыть файл генератора", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
# show an "Open" dialog box and return the path to the selected file

wb = workbook
f_open = False
# Проверяем, закрыт ли файл генератора


try:
    if '.xlsx' in filename:
        # работаем с книгой
        wb = load_workbook(filename=filename)
        # Когда и кто последний раз редактировал дркумент

        date_modified = (wb.properties.modified + utc3).strftime("%Y%m%d_%H%M%S")
        lastModifiedBy = wb.properties.lastModifiedBy
        save_name = 'prGen_' + date_modified + '_' + translit(lastModifiedBy, reversed=True).replace(' ',
                                                                                                     '_') + '.xlsx'
        Excel = win32com.client.Dispatch("Excel.Application")

        if Excel.Workbooks.Count != 0:
            for i in range(1, Excel.Workbooks.Count + 1):
                if Excel.Workbooks(i).Name in save_name:

                    if not Excel.Workbooks(i).Saved:
                        print("Файл редактировали, но не сохраняли")
                        date_now = datetime.now().strftime("%Y%m%d_%H%M%S")
                        date_modified = date_now
                        lastModifiedBy = wb.properties.lastModifiedBy
                        save_name = 'prGen_' + date_modified + '_' + \
                                    translit(lastModifiedBy, reversed=True).replace(' ', '_') + '.xlsx'
                        filename = Excel.Workbooks(i).Path + save_name
                        Excel.Workbooks(i).SaveAs(filename)
                        Excel.Workbooks(i).Close()
                        wb.close()
                        wb = load_workbook(filename=filename)

                    elif Excel.Workbooks(i).Saved:
                        print("Файл сохраняли")
                        Excel.Workbooks(i).Close()
                    break

        f_open = True
        print(filename)
        print(save_name)

except IOError:
    print("Нельзя работать с открытым файлом!!!")
