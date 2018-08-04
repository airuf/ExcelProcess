# -.*- coding: UTF-8 -*-
import xlrd


class AHS_xlrd():
    def __init__(self, name, path):
        self.name = name
        self.path = path
        self.members = 0
        self.m_dic = {}

    def count_members(self, name, shop, address, amount, date):
        if name in self.m_dic:
            self.m_dic['name'].append({'shop': shop, 'address': address, 'amount':amount, 'date':date})
        else:
            self.m_dic.setdefault('name', {'shop': shop, 'address': address, 'amount':amount, 'date':date})

    # def performance_members(self):
    #

path1 = 'C:/Users/LMAN/Desktop/PYTHON TEST/exclT/1.xlsx'

workbook = xlrd.open_workbook(path1)
sheet = workbook.sheet_by_index(0)
# print(sheet.name,sheet.nrows,sheet.ncols)
rows = sheet.row_values(3)
# print("rows: ", rows)



# print(workbook.sheet_names())
# print(workbook)


