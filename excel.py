"""\\\\\\\\\\\\\\\РАБОТА С ТАБЛИЦЕЙ\\\\\\\\\\\\\\\\"""
import openpyxl
import settings
import devices
import os
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.comments import Comment
from datetime import datetime


def fill_ts():
    ts = wb.worksheets[1]

    # №
    fill = PatternFill(fill_type='solid', start_color='9bc2e6', end_color='9bc2e6')
    bord_side = Side(border_style='thin', color='00000000')
    bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 1)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = ts.cell(6 + i, 2).value

    # Наименование логического параметра
    align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 2)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = ts.cell(6 + i, 3).value

    # Тип параметра
    align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 3)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = 'физический'

    # Функция ASDU
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 4)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = 'M_SP_NA_1 (1)'

    # Адрес объекта
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 5)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = settings.IEC_104['startAddressTS'] + i

    # Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР)
    for k in range(4):
        for i in range(settings.discrete_count):
            cell = bd.cell(index + i, 6 + k)
            cell.fill = fill
            cell.border = bord
            cell.alignment = align
            cell.value = '-'

    # Расшифровка значения
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 10)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = ts.cell(6 + i, 13).value

    # № физического канала
    for i in range(settings.discrete_count):
        cell = bd.cell(index + i, 11)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = str(ts.cell(6 + i, 4).value) + "." + str(ts.cell(6 + i, 5).value)

    # Примечание,
    for k in range(2):
        for i in range(settings.discrete_count):
            cell = bd.cell(index + i, 12 + k)
            cell.fill = fill
            cell.border = bord
            cell.alignment = align
            cell.value = '-'
"""================Заполняем БД ВУ================="""
# Открываем файл excel
print('================Заполняем БД ВУ=================')
filename = 'pr_gen'
wb = openpyxl.load_workbook(filename=filename + '.xlsx')

bd = wb.worksheets[8]

# очищаем базу
clr = bd['A9:M500']
fill = PatternFill()
bord_side = Side()
bord = Border()
align = Alignment()
for row in clr:
    for cell in row:
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = None
"""==Заголовки=="""

print(' Заголовки')
bd['C1'] = wb.worksheets[0].cell(3, 2).value  # Название объекта телемеханизации
modules = ''    # Расположение модулей в контроллере
for i in range(settings.module_count):
    modules = modules + str(wb.worksheets[0].cell(10+i, 2).value) + ', '
bd['C3'] = modules
extra1 = 'ASDU ' + str(settings.IEC_104['ASDU'])  # Дополнительная информация по контроллеру и протоколу обмена
bd['C5'] = extra1
extra2 = 'k=' + str(settings.IEC_104['k']) + ', ' + 'w=' + str(settings.IEC_104['w']) + ', ' + 'T0=' + \
         str(settings.IEC_104['timeoutK'])
bd['L5'] = extra2

"""==Физические ТС=="""
index = 9
print(' Физические ТС')
fill_ts()


"""==Физические ТИ=="""

print(' Физические ТИ')
index = index + settings.discrete_count
ti = wb.worksheets[3]

# №
fill = PatternFill(fill_type='solid', start_color='ccffcc', end_color='ccffcc')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.input_count):
    cell = bd.cell(index + i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6 + i, 2).value


# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(settings.input_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.input_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(settings.input_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'M_ME_NB_1 (11)'

# Адрес объекта
for i in range(settings.input_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = settings.IEC_104['startAddressTI'] + i

# Нижний диапазон
for i in range(settings.input_count):
    cell = bd.cell(index+i, 6)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 10).value

# Верхний диапазон
for i in range(settings.input_count):
    cell = bd.cell(index+i, 7)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 11).value

# Ед. измерения
for i in range(settings.input_count):
    cell = bd.cell(index+i, 8)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = ti.cell(6+i, 17).value


# Значение по умолчанию (для ТР), Расшифровка значения
for k in range(2):
    for i in range(settings.input_count):
        cell = bd.cell(index+i, 9+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# № физического канала
for i in range(settings.input_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(ti.cell(6+i, 4).value) + "." + str(ti.cell(6+i, 5).value)

# Примечание
for k in range(2):
    for i in range(settings.input_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Физические ТУ=="""

print(' Физические ТУ')
tu = wb.worksheets[2]
index = index + settings.input_count
# №
fill = PatternFill(fill_type='solid', start_color='1f4e78', end_color='1f4e78')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 2).value

# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'C_SC_NA_1 (45)'

# Адрес объекта
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = settings.IEC_104['startAddressTU'] + i

# Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР)
for k in range(4):
    for i in range(settings.discrete_output_count):
        cell = bd.cell(index+i, 6+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# Расшифровка значения
for i in range(settings.discrete_output_count):
    cell = bd.cell(index + i, 10)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tu.cell(6+i, 11).value

# № физического канала
for i in range(settings.discrete_output_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(tu.cell(6+i, 4).value) + "." + str(tu.cell(6+i, 5).value)

# Примечание,
for k in range(2):
    for i in range(settings.discrete_output_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Физические ТР=="""

print(' Физические ТР')
tr = wb.worksheets[4]
index = index + settings.discrete_output_count
# №
fill = PatternFill(fill_type='solid', start_color='fabf8f', end_color='fabf8f')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.output_count):
    cell = bd.cell(index+i, 1)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tr.cell(6+i, 2).value

# Наименование логического параметра
align = Alignment(horizontal='left', vertical='center', wrap_text=True)
for i in range(settings.output_count):
    cell = bd.cell(index+i, 2)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = tr.cell(6+i, 3).value

# Тип параметра
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
for i in range(settings.output_count):
    cell = bd.cell(index+i, 3)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'физический'

# Функция ASDU
for i in range(settings.output_count):
    cell = bd.cell(index+i, 4)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = 'C_SE_NC_1 (50)'

# Адрес объекта
for i in range(settings.output_count):
    cell = bd.cell(index+i, 5)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = settings.IEC_104['startAddressTRF'] + i

# Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР), Расшифровка значения
for k in range(5):
    for i in range(settings.output_count):
        cell = bd.cell(index+i, 6+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'

# № физического канала
for i in range(settings.output_count):
    cell = bd.cell(index+i, 11)
    cell.fill = fill
    cell.border = bord
    cell.alignment = align
    cell.value = str(tr.cell(6+i, 4).value) + "." + str(tr.cell(6+i, 5).value)

# Примечание,
for k in range(2):
    for i in range(settings.output_count):
        cell = bd.cell(index+i, 12+k)
        cell.fill = fill
        cell.border = bord
        cell.alignment = align
        cell.value = '-'


"""==Интерфейсные ТС=="""
index = index + settings.output_count
# №
fill = PatternFill(fill_type='solid', start_color='9bc2e6', end_color='9bc2e6')
bord_side = Side(border_style='thin', color='00000000')
bord = Border(bottom=bord_side, left=bord_side, top=bord_side, right=bord_side)
align = Alignment(horizontal='center', vertical='center', wrap_text=True)
ind_ts = 1
for i in range(len(devices.Modbus)):
    for j in range(len(devices.Modbus[i])):
        transaction = devices.Modbus[i][j]
        if transaction.mfc == 1 or transaction.mfc == 2:
            for k in range(transaction.len_tr):
                # №
                cell = bd.cell(index + k, 1)
                cell.fill = fill
                cell.border = bord
                cell.alignment = align
                cell.value = ind_ts

                # Наименование логического параметра
                align = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell = bd.cell(index + k, 2)
                cell.fill = fill
                cell.border = bord
                cell.alignment = align
                cell.value = transaction.name + ' ТС №' + str(k+1)


                # # Тип параметра
                # align = Alignment(horizontal='center', vertical='center', wrap_text=True)
                # cell = bd.cell(index + i, 3)
                # cell.fill = fill
                # cell.border = bord
                # cell.alignment = align
                # cell.value = 'физический'
                #
                # # Функция ASDU
                # cell = bd.cell(index + i, 4)
                # cell.fill = fill
                # cell.border = bord
                # cell.alignment = align
                # cell.value = 'C_SE_NC_1 (50)'
                #
                # # Адрес объекта
                # cell = bd.cell(index + i, 5)
                # cell.fill = fill
                # cell.border = bord
                # cell.alignment = align
                # cell.value = settings.IEC_104['startAddressTRF'] + i
                #
                # # Нижний диапазон, Верхний диапазон, Ед. измерения, Значение по умолчанию (для ТР), Расшифровка значения
                # for k in range(5):
                #     cell = bd.cell(index + i, 6 + k)
                #     cell.fill = fill
                #     cell.border = bord
                #     cell.alignment = align
                #     cell.value = '-'
                #
                # # № физического канала
                # cell = bd.cell(index + i, 11)
                # cell.fill = fill
                # cell.border = bord
                # cell.alignment = align
                # cell.value = str(tr.cell(6 + i, 4).value) + "." + str(tr.cell(6 + i, 5).value)
                #
                # # Примечание,
                # for k in range(2):
                #     cell = bd.cell(index + i, 12 + k)
                #     cell.fill = fill
                #     cell.border = bord
                #     cell.alignment = align
                #     cell.value = '-'
                ind_ts += 1

            index += transaction.len_tr
print('База готова')

wb.save(filename=filename + '.xlsx')
wb.template = False
wb.close()
print('Файл сохранен')
os.system(filename + '.xlsx')

""" заготовка под сохранение с текущей датой и временем
date = datetime.now().timetuple()
new_filename = filename + '_' + str(date.tm_year) + '.' + str(date.tm_mon) + '.' + str(date.tm_mday) + '_' + \
               str(date.tm_hour) + '-' + str(date.tm_min) + '-' + str(date.tm_sec)
print(new_filename)
wb.save(filename=new_filename + '.xlsx')
wb.close()
print('Файл сохранен')
os.system(new_filename)
"""