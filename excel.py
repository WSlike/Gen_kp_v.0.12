"""\\\\\\\\\\\\\\\РАБОТА С ТАБЛИЦЕЙ\\\\\\\\\\\\\\\\"""
import openpyxl
import os
from openpyxl.styles import PatternFill, Border, Side, Alignment


"""================Заполняем настройки================="""

# Открываем файл excel
print('Открываем файл...')
filename = 'pr_gen.xlsx'
wb = openpyxl.load_workbook(filename=filename)

print('Читаем Количесвто сигналов и Настройки протокола...')
# Считаем количество ТС
ws = wb.worksheets[1]
if ws.cell(2, 2).value is '+':
    discrete_count = ws.max_row - 5
else:
    discrete_count = 0

# Считаем количество ТУ
ws = wb.worksheets[2]
if ws.cell(2, 2).value is '+':
    discrete_output_count = ws.max_row - 5
else:
    discrete_output_count = 0

# Считаем количество ТИ
ws = wb.worksheets[3]
if ws.cell(2, 2).value is '+':
    input_count = ws.max_row - 5
else:
    input_count = 0

# Считаем количество ТР
ws = wb.worksheets[4]
if ws.cell(2, 2).value is '+':
    output_count = ws.max_row - 5
else:
    output_count = 0

# Считаем количество модулей
ws = wb.worksheets[0]
module_count = 0
for i in range(16):
    if ws.cell(10 + i, 4).value is None:
        break
    if module_count < ws.cell(10 + i, 4).value:
        module_count = ws.cell(10 + i, 4).value


# Считываем настройки 104 протокола
ws = wb.worksheets[0]
IEC_104 = {'ASDU': ws.cell(29, 2).value, 'k': ws.cell(30, 2).value, 'w': ws.cell(31, 2).value,
           'timeoutK': ws.cell(32, 2).value, 'startAddressTS': ws.cell(33, 2).value,
           'startAddressTI': ws.cell(34, 2).value, 'startAddressTF': ws.cell(35, 2).value,
           'startAddressTU': ws.cell(36, 2).value, 'startAddressTRI': ws.cell(37, 2).value,
           'startAddressTRF': ws.cell(38, 2).value, 'commandMaxCount': ws.cell(39, 2).value,
           'eventMaxCount': ws.cell(40, 2).value, 'inOutPacketMaxCount': ws.cell(41, 2).value}


"""================Заполняем БД ВУ================="""
print('Заполняем БД ВУ...')
bd = wb.worksheets[8]

"""==Заголовки=="""

print(' Заголовки')
bd['C1'] = wb.worksheets[0].cell(3, 2).value  # Название объекта телемеханизации
modules = ''    # Расположение модулей в контроллере
for i in range(module_count):
    modules = modules + str(wb.worksheets[0].cell(10+i, 2).value) + ', '
bd['C3'] = modules
extra1 = 'ASDU ' + str(IEC_104['ASDU'])  # Дополнительная информация по контроллеру и протоколу обмена
bd['C5'] = extra1
extra2 = 'k=' + str(IEC_104['k']) + ', ' + 'w=' + str(IEC_104['w']) + ', ' + 'T0=' + \
         str(IEC_104['timeoutK'])
bd['L5'] = extra2

"""==Физические ТС=="""

print(' Физические ТС')
ts = wb.worksheets[1]
index = 9
# №
fill = PatternFill(fill_type='solid', start_color='9bc2e6', end_color='9bc2e6')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(discrete_count):
    cell = bd.cell(index+i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ts.cell(6+i, 2).value

# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(discrete_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ts.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(discrete_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(discrete_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'M_SP_NA_1 (1)'

# Адрес объекта
for i in range(discrete_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = IEC_104['startAddressTS'] + i

# Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР)
for k in range(3):
    for i in range(discrete_count):
        cell = bd.cell(index+i, 6+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# Расшифровка значения
for i in range(discrete_count):
    cell = bd.cell(index + i, 10)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ts.cell(6+i, 13).value

# № физического канала
for i in range(discrete_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(ts.cell(6+i, 4).value) + "." + str(ts.cell(6+i, 5).value)

# Примечание,
for k in range(2):
    for i in range(discrete_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Физические ТИ=="""

print(' Физические ТИ')
index = index + discrete_count
ti = wb.worksheets[3]

# №
fill = PatternFill(fill_type='solid', start_color='ccffcc', end_color='ccffcc')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(input_count):
    cell = bd.cell(index + i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6 + i, 2).value


# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(input_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(input_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(input_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'M_ME_NB_1 (11)'

# Адрес объекта
for i in range(input_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = IEC_104['startAddressTI'] + i

# Нижний диапазон
for i in range(input_count):
    cell = bd.cell(index+i, 6)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 10).value

# Верхний диапазон
for i in range(input_count):
    cell = bd.cell(index+i, 7)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 11).value

# Ед. измерения
for i in range(input_count):
    cell = bd.cell(index+i, 8)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 17).value


# Значение по умолчанию (для ТР), Расшифровка значения
for k in range(2):
    for i in range(input_count):
        cell = bd.cell(index+i, 9+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# № физического канала
for i in range(input_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(ti.cell(6+i, 4).value) + "." + str(ts.cell(6+i, 5).value)

# Примечание
for k in range(2):
    for i in range(input_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Физические ТУ=="""

print(' Физические ТУ')
tu = wb.worksheets[2]
index = index + input_count
# №
fill = PatternFill(fill_type='solid', start_color='1f4e78', end_color='1f4e78')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 2).value

# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'C_SC_NA_1 (45)'

# Адрес объекта
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = IEC_104['startAddressTU'] + i

# Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР)
for k in range(4):
    for i in range(discrete_output_count):
        cell = bd.cell(index+i, 6+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# Расшифровка значения
for i in range(discrete_output_count):
    cell = bd.cell(index + i, 10)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 11).value

# № физического канала
for i in range(discrete_output_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(tu.cell(6+i, 4).value) + "." + str(tu.cell(6+i, 5).value)

# Примечание,
for k in range(2):
    for i in range(discrete_output_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Физические ТР=="""

print(' Физические ТР')
tr = wb.worksheets[4]
index = index + discrete_output_count
# №
fill = PatternFill(fill_type='solid', start_color='fabf8f', end_color='fabf8f')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(output_count):
    cell = bd.cell(index+i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tr.cell(6+i, 2).value

# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(output_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tr.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(output_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(output_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'C_SE_NC_1 (50)'

# Адрес объекта
for i in range(output_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = IEC_104['startAddressTRF'] + i

# Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР), Расшифровка значения
for k in range(5):
    for i in range(output_count):
        cell = bd.cell(index+i, 6+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# № физического канала
for i in range(output_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(tr.cell(6+i, 4).value) + "." + str(tr.cell(6+i, 5).value)

# Примечание,
for k in range(2):
    for i in range(output_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

print('База готова')
wb.save(filename=filename)
wb.close()
print('Файл сохранен')

os.system(filename)
