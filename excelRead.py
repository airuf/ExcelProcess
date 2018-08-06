# -.*- coding: UTF-8 -*-
import xlrd

class BasicInfo:
    def __init__(self, name, path):
        self.name = name;
        self.

class AHS_xlrd:
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

排名	    rank
门店名称	store name
经办人         manager
成交单数        number of orders
商家服务费       merchant service fee
店员服务费       clerk service fee
金蛋金额        gloden egg amount
店员收入合计      total clerk's income
门店服务费合计     total store service fee
所属总账户       master account

## 基本信息 ##
## clerk info ##
# 名字 name
# 手机 phone number
# 订单 orders

## 单子信息 ##
## order infor ##
# 订单号 number of order
# 下单价 contract price
# 成交价 final price
# 商家服务费 store service fee
# 店长服务费 store manager service fee
# 店员服务费 clerk service fee
# 金蛋金额 gloden amount
# 下单时间 order time
# 订单状态 order state
# 成交时间 transcation time
# 验货状态 inspection time
# 发货单号 number of deilvery
# 实付旧机款 actual payment
# 旧机款结算账户 settlement account

## 店铺 ##
## store ##
# 排名 rank
# 名称 name
# 经办人 manager
# 单子总数 number of orders
# 店铺服务费 merchant service fee
# 金蛋总金额gloden egg fee

## 所属总账户 ##
## master account ##
# 名字 name

class order:
    def __init__(self, order_number):
class MasterAccount():
    def __init__(self, name):
        self.name = name;







