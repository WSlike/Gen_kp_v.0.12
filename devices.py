import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment


class Transaction:
    mfc = offset = len = timeout = 0

    def __init__(self, mfc=None, offset=None, lenght=None, timeout=None):
        self.mfc = mfc
        self.offset = offset
        self.lenght = lenght
        self.timeout = timeout


# Открываем файл excel
print('=======Читаем настройки внешних устройств========')
filename = 'pr_gen'
wb = openpyxl.load_workbook(filename=filename + '.xlsx')

ws = wb.worksheets[7]
dev = {}  # список {Название устройства: Код устройства}
j = 0
dev_row = []  # массив с начальными строками устройств
for i in range(4, ws.max_row):
    dev_key = ws.cell(i, 4).value
    if dev_key not in dev and (dev_key is not None) and (dev_key != 'Название устройства'):
        dev[dev_key] = ws.cell(i, 3).value
        dev_row.append(i)
dev_row.append(ws.max_row + 2)

print('Используемые устройства:')
print(dev)
for key in dev:
    print(' ' + key)

# Заполнение массива транзакций
Modbus = []
Device = []
for i in range(len(dev_row) - 1):
    first = dev_row[i] + 5
    last = dev_row[i + 1] - 2
    Device = []
    index_tr = 0
    for j in range(first, last+1):
        mfc = ws.cell(j, 6).value  # функция
        offset = ws.cell(j, 7).value  # смещение
        reg_type = ws.cell(j, 8).value  # тип данных
        col = ws.cell(j, 9).value  # количество данных
        timeout = ws.cell(j, 10).value  # таймаут опроса
        index_tr += 1
        if reg_type == 'float' or reg_type == 'long':
            len_tr = col * 2  # количесвто регистров
        else:
            len_tr = col  # количесвто регистров

        Transaction.mfc = mfc
        Transaction.offset = offset
        Transaction.len_tr = len_tr
        Transaction.timeout = timeout
        Device.append(Transaction)
        print(ws.cell(dev_row[i], 4).value, 'Запрос №', index_tr, 'Функция =', mfc, 'смещение =', offset,
              'кол-во регистров', len_tr, 'таймаут запроса =', timeout)

    Modbus.append(Device)

wb.close()
