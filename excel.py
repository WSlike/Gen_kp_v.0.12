import openpyxl
import xml.etree.ElementTree as ET

wb = openpyxl.load_workbook(filename='Gen_Уренгой-Пур-Пэ_КП15.xlsx')

ws = wb.worksheets[0]
IEC_104 = {'ASDU': ws.cell(5, 1).value,
           'k': ws.cell(5, 2).value,
           'w': ws.cell(5, 3).value,
           'startTS': ws.cell(5, 4).value,
           'startTI': ws.cell(5, 5).value,
           'startTF': ws.cell(5, 6).value,
           'startTU': ws.cell(5, 7).value,
           'startTR': ws.cell(5, 8).value
           }

print(ws.title)
print(IEC_104)

a = ET.Element('a')
b = ET.SubElement(a, 'b')
c = ET.SubElement(a, 'c')
d = ET.SubElement(c, 'd')
ET.dump(a)