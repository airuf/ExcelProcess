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

# 排名	    rank
# 门店名称	store name
# 经办人         manager
# 成交单数        number of orders
# 商家服务费       merchant service fee
# 店员服务费       clerk service fee
# 金蛋金额        gloden egg amount
# 店员收入合计      total clerk's income
# 门店服务费合计     total store service fee
# 所属总账户       master account

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


## 单子信息 ##
## order infor ##
# 订单号 number of order
# 下单价 contract price
# 成交价 final price
# 商家服务费 store service fee
# 店长服务费 store manager service fee
# 店员服务费 clerk service fee
# 金蛋金额 gloden egg amount
# 下单时间 order time
# 订单状态 order state
# 成交时间 transcation time
# 验货状态 inspection time
# 发货单号 number of deilvery
# 实付旧机款 actual payment
# 旧机款结算账户 settlement account

class OrderInfo:
    def __init__(self, order_number):
        self.order_number.key = order_number
        self.contract_price = False
        self.final_price = False
        self.store_service_fee = False
        self.store_manager_service_fee = False
        self.clerk_service_fee = False
        self.gloden_egg_amount = False
        self.order_time = False
        self.order_state = False
        self.transcation_time = False
        self.inspection_time = False
        self.number_of_deilvery = False
        self.actual_payment = False
        self.settlement_account = False

    def set_contract_price(self, contract_price):
        self.contract_price = contract_price

    def set_final_price(self, final_price):
        self.final_price = final_price

    def set_store_service_fee(self, store_service_fee):
        self.store_service_fee = store_service_fee

    def set_store_manager_service_fee(self, store_manager_service_fee):
        self.store_manager_service_fee = store_manager_service_fee

    def set_clerk_service_fee(self, clerk_service_fee):
        self.clerk_service_fee = clerk_service_fee

    def set_gloden_egg_amount(self, gloden_egg_amount):
        self.gloden_egg_amount = gloden_egg_amount

    def set_order_time(self, order_time):
        self.order_time = order_time

    def set_order_state(self, order_state):
        self.order_state = order_state

    def set_transcation_time(self, transcation_time):
        self.transcation_time = transcation_time

    def set_inspection_time(self, inspection_time):
        self.inspection_time = inspection_time

    def set_number_of_delivery(self, number_of_delivery):
        self.number_of_delivery = number_of_delivery

    def set_actual_payment(self, actual_payment):
        self.actual_payment = actual_payment

    def set_settlement_account(self, settlement_account):
        self.settlement_account = settlement_account

    def is_order_integrity(self):
        for key,value in vars(self).items():
            if value is False:
                raise AssertionError("the %s is null", key)



## 基本信息 ##
## clerk info ##
# 名字 name
# 手机 phone number
# 订单 orders
#
class ClerkInfo():
    def __init__(self, name):
        self.name = name
        self.phone_number = 0
        self.orders = []

    def create_order(self, order_number):
        self.orders.append(OrderInfo(order_number))

    def set_order_data(self, order_number):


## 所属总账户 ##
## master account ##
# 名字 name
class MasterAccount():
    def __init__(self, name):
        self.name = name







