from openfile import wb, f_open


class Transaction:

    def __init__(self, name=None, mfc=None, offset=None, col=None, len_tr=None, timeout=None):
        self.name = name
        self.mfc = mfc
        self.offset = offset
        self.col = col
        self.len_tr = len_tr
        self.timeout = timeout


if f_open:
    print('==========Читаем таблицу внешних устройств============')

    ws_dev = wb.worksheets[7]
    dev = {}  # список {Название устройства: Код устройства}
    j = 0
    dev_row = []  # массив с начальными строками устройств
    for i in range(4, ws_dev.max_row):
        dev_key = ws_dev.cell(i, 4).value
        if dev_key not in dev and (dev_key is not None) and (dev_key != 'Название устройства'):
            dev[dev_key] = ws_dev.cell(i, 3).value
            dev_row.append(i)
    dev_row.append(ws_dev.max_row + 2)

    print('Используемые устройства:')
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
            dev_mfc = ws_dev.cell(j, 6).value  # функция
            dev_offset = ws_dev.cell(j, 7).value  # смещение
            reg_type = ws_dev.cell(j, 8).value  # тип данных
            dev_col = ws_dev.cell(j, 9).value  # количество данных
            dev_timeout = ws_dev.cell(j, 10).value  # таймаут опроса

            if reg_type == 'float' or reg_type == 'long':
                dev_len_tr = dev_col * 2  # количесвто регистров
            else:
                dev_len_tr = dev_col  # количесвто регистров
            t = Transaction(name=ws_dev.cell(dev_row[i], 4).value, mfc=dev_mfc, offset=dev_offset, col=dev_col,
                            len_tr=dev_len_tr, timeout=dev_timeout)
            Device.append(t)
            index_tr += 1

        Modbus.append(Device)

    for i in range(len(Modbus)):
        for j in range(len(Modbus[i])):
            transaction = Modbus[i][j]
            # print('\tЗапрос =', j, '\tФункция =', transaction.mfc, '\tсмещение =', transaction.offset,
            #     '\tкол-во регистров =', transaction.len_tr, '\tтаймаут запроса =', transaction.timeout)
