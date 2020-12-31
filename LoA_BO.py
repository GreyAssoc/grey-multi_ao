from docxtpl import DocxTemplate
import openpyxl
import pathlib
import xlrd
from pick import pick

dir = pathlib.Path().absolute()
ao_list = ["AO 1", "AO 2", "AO 3", "AO 4", "AO 5"]
#ao = input("Select the AO you wish to produce a document for: ", ao_list[0], ao_list[1], ao_list[2], ao_list[3], ao_list[4])

#User selects AO
title = "Select the AO you wish to produce a document for:"
options = ["AO 1", "AO 2", "AO 3", "AO 4", "AO 5"]
selected_ao, index = pick(options, title)
print(selected_ao)
print(index)


data = '00.0 - DATA Pemberton.Steven_Caroline - 54 The Chase.xlsx'
wb = openpyxl.load_workbook(filename=data, data_only=True)
#wb = xlrd.open_workbook('00.0 - DATA Pemberton.Steven_Caroline - 54 The Chase.xlsx')

#print(data)
#ws = wb[str("'"+selected_ao+"'")]
ws = wb[selected_ao]

letter_name = ws['A1'].value
notice_name = ws['B2'].value
fhlh = ws['C2'].value
ao_i_we = ws['D2'].value
ao_my_our = ws['E2'].value
ao_his_her = ws['F2'].value
ao_he_she = ws['G2'].value
ao_do_does = ws['H2'].value
ao_have_has = ws['I2'].value
ao_him_her = ws['J2'].value
property_horz = ws['K2'].value
correspond_horz = ws['L2'].value

print(letter_name)
print(notice_name)
print(fhlh)
print(ao_i_we)


