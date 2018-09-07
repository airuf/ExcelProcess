# -.*- coding: UTF-8 -*-
import xlrd

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
# 订单号 number of order                       rows[0]

# 所属总账户名称 master account name           rows[37]
# 所属门店名称 store name                      rows[42]
# 经办人姓名 manager name                      rows[51]

# 下单价 contract price                        rows[6]
# 成交价 final price                           rows[10]
# 商家服务费 store service fee                 rows[15]
# 店长服务费 store manager service fee         rows[16]
# 店员服务费 clerk service fee                 rows[17]
# 金蛋金额 gloden egg amount                   rows[20]
# 订单状态 order state                         rows[30]
# 验货状态 inspection state                    rows[31]
# 下单时间 order time                          rows[53]
# 成交时间 transcation time                    rows[54]
# 发货单号 number of deilvery                  rows[55]
# 实付旧机款 actual payment                    rows[56]
# 旧机款结算账户 settlement account            rows[57]

class OrderInfo:
    def __init__(self, order_number):
        self.order_number = order_number

        self.master_account_name = False
        self.store_name = False
        self.manager_name = False

        self.contract_price = False
        self.final_price = False
        self.store_service_fee = False
        self.store_manager_service_fee = False
        self.clerk_service_fee = False
        self.gloden_egg_amount = False
        self.order_state = False
        self.order_time = False
        self.inspection_time = False
        self.transcation_time = False
        self.number_of_deilvery = False
        self.actual_payment = False
        self.settlement_account = False

####
    def set_contract_price(self, contract_price):
        self.contract_price = contract_price

####
    def set_master_accout_name(self, master_account_name):
        self.master_account_name = master_account_name

    def set_store_name(self, store_name):
        self.store_name = store_name

    def set_manager_name(self, manager_name):
        self.manager_name = manager_name

####
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

    def set_order_state(self, order_state):
        self.order_state = order_state

    def set_inspection_state(self, inspection_state):
        self.inspection_state = inspection_state

    def set_order_time(self, order_time):
        self.order_time = order_time

    def set_transcation_time(self, transcation_time):
        self.transcation_time = transcation_time

    def set_number_of_delivery(self, number_of_delivery):
        self.number_of_delivery = number_of_delivery

    def set_actual_payment(self, actual_payment):
        self.actual_payment = actual_payment

    def set_settlement_account(self, settlement_account):
        self.settlement_account = settlement_account

    def is_order_integrity(self):
        for key,value in vars(self).items():
            if value is False:
                raise AssertionError("the #" + key + "# is null")





## 基本信息 ##
## clerk info ##
# 名字 name
# 手机 phone number
# 订单 orders
#
# 成交单数合计    numbers of orders
# 商家服务费合计   all store service fee
# 店员服务费合计   all clerk service fee
# 金蛋金额合计    all gloden egg amounts
# 店员收入合计    all income


class ClerkInfo():
    def __init__(self, name):
        self.name = name
        self.phone_number = 0
        self.orders = {}
        self.number_of_orders = 0
        self.all_store_service_fee = 0
        self.all_clerk_service_fee = 0
        self.all_gloden_egg_amounts = 0
        self.all_income = 0

    def add_create_order(self, order_number):
        self.orders.setdefault(order_number, OrderInfo(order_number))

    def total_result(self):
        # add up to all orders
        self.number_of_orders = len(self.orders)
        store_service_fee = 0
        clerk_service_fee = 0
        gloden_egg_amounts = 0
        income = 0
        for order in self.orders.values():
            store_service_fee += order.store_service_fee
            clerk_service_fee += order.clerk_service_fee
            ##TODO:
            if isinstance(order.gloden_egg_amount, str):
                gloden_egg_amounts += 0
            else:
                gloden_egg_amounts += order.gloden_egg_amount
            income = clerk_service_fee + gloden_egg_amounts
        self.all_store_service_fee = store_service_fee
        self.all_clerk_service_fee = clerk_service_fee
        self.all_gloden_egg_amounts = gloden_egg_amounts
        self.all_income = income

## 所属门店名称 ##
## store name  ##       rows[42]
class StoreInfo(object):
    def __init__(self, object):
        pass

    def __init__(self, name):
        self.name = name
        self.all_income = 0
        self.clerks = {}
        self.num_of_clerks = 0

    def add_create_clerk(self, name):
        self.clerks.setdefault(name, ClerkInfo(name))

    def get_num_clerks(self):
        self.num_of_clerks = len(self.clerks)
        num_of_clerks = self.num_of_clerks
        return num_of_clerks

    def total_result(self):
        # add up to all clerks income
        # self.number_of_orders = len(self.orders)
        store_service_fee = 0
        clerk_service_fee = 0
        gloden_egg_amounts = 0
        income = 0
        for clerks in self.clerks.values():
            clerks.total_result()
            store_service_fee += clerks.all_store_service_fee
            clerk_service_fee += clerks.all_clerk_service_fee
            gloden_egg_amounts += clerks.all_gloden_egg_amounts
            income += clerks.all_income
        self.all_store_service_fee = store_service_fee
        self.all_clerk_service_fee = clerk_service_fee
        self.all_gloden_egg_amounts = gloden_egg_amounts
        self.all_income = income

## 所属总账户 ##
## master account ##
# 名字 name
class MasterAccount(object):
    def __init__(self, name):
        self.name = name
        self.store = {}

    def add_create_store(self, name):
        self.store.setdefault(name, StoreInfo(name))



def test_me():
    #####  step 1  建立总公司  #####
    HY = MasterAccount('福建泉州市华远电讯有限公司')

    #####          step 2  扫描单子         #####
    #####          step 3  单子分店         #####
    #####          step 3  单子分人         #####
    #####          step 4  单子存到人里     #####
    ## get data ##

    path1 = 'C:/Users/lin/Desktop/ExcelProcess/data/7月总体数据.xlsx'
    path2 = 'C:/Users/LMAN/Desktop/PYTHON TEST/exclT/7月总体数据.xlsx'  ## win 7
    workbook = xlrd.open_workbook(path2)
    sheet = workbook.sheet_by_index(0)

    ### 遍历总体数据，找出限定条件内的  单子， 放进职员的 单子表中 #######
    ### 限定条件 ###  sheet.row_values(i)[num]###
    ### 37 公司名称 ####
    ### 30 订单状态  ！=交易取消    ####  6 下单价 > 5    ######
    ####  职员表  ######
    ####  51  经办人姓名    #### 42 门店名称  ### 0 订单号 number of order   #####
    x= 0
    for i in range(1, sheet.nrows):
        if sheet.row_values(i)[37] == HY.name:
            if sheet.row_values(i)[30] != "交易取消" and sheet.row_values(i)[6] > 5:
                HY.add_create_store(sheet.row_values(i)[42])
                HY.store[sheet.row_values(i)[42]].add_create_clerk(sheet.row_values(i)[51])
                HY.store[sheet.row_values(i)[42]].clerks[sheet.row_values(i)[51]].add_create_order(sheet.row_values(i)[0])

                _order = HY.store[sheet.row_values(i)[42]].clerks[sheet.row_values(i)[51]].orders[sheet.row_values(i)[0]]

                _order.master_account_name = sheet.row_values(i)[37]
                _order.store_name = sheet.row_values(i)[42]
                _order.manager_name = sheet.row_values(i)[51]

                _order.contract_price = sheet.row_values(i)[6]
                _order.final_price = sheet.row_values(i)[10]
                _order.store_service_fee = sheet.row_values(i)[15]
                _order.store_manager_service_fee = sheet.row_values(i)[16]
                _order.clerk_service_fee = sheet.row_values(i)[17]
                _order.gloden_egg_amount = sheet.row_values(i)[20]
                _order.order_state = sheet.row_values(i)[30]
                _order.order_time = sheet.row_values(i)[31]
                _order.inspection_time = sheet.row_values(i)[53]
                _order.transcation_time = sheet.row_values(i)[54]
                _order.number_of_deilvery = sheet.row_values(i)[55]
                _order.actual_payment = sheet.row_values(i)[56]
                _order.settlement_account = sheet.row_values(i)[57]

                # print(HY.store[sheet.row_values(i)[42]].clerks[sheet.row_values(i)[51]])
                x += 1
                # if sheet.row_values(i)[51] in clerk_order_dic.keys():
                #     clerk_order_dic[sheet.row_values(i)[51]] += 1
            else:
                a = 1
                # print(i, sheet.row_values(i)[51], sheet.row_values(i)[42], sheet.row_values(i)[30])
    print("x", x)



    # print("rows: ", rows)

    ## 单子信息 ##
    ## order infor ##
    # 订单号 number of order                       rows[0]

    # 所属总账户名称 master account name           rows[37]
    # 所属门店名称 store name                      rows[42]
    # 经办人姓名 manager name                      rows[51]

    # 下单价 contract price                        rows[6]
    # 成交价 final price                           rows[10]
    # 商家服务费 store service fee                 rows[15]
    # 店长服务费 store manager service fee         rows[16]
    # 店员服务费 clerk service fee                 rows[17]
    # 金蛋金额 gloden egg amount                   rows[20]
    # 订单状态 order state                         rows[30]
    # 验货状态 inspection state                    rows[31]
    # 下单时间 order time                          rows[53]
    # 成交时间 transcation time                    rows[54]
    # 发货单号 number of deilvery                  rows[55]
    # 实付旧机款 actual payment                    rows[56]
    # 旧机款结算账户 settlement account            rows[57]


    # path1 = 'C:/Users/lin/Desktop/ExcelProcess/data/7月总体数据.xlsx'
    path2 = 'C:/Users/LMAN/Desktop/PYTHON TEST/exclT/ExcelProcess/data/7月总体数据.xlsx'
    workbook = xlrd.open_workbook(path2)
    sheet = workbook.sheet_by_index(0)
    # print(sheet.name, sheet.nrows, sheet.ncols)
    # rows = sheet.row_values(0)
    # for index, value in enumerate(rows):
    #     print("索引：" + str(index), ", 值：" + value)



    ## save order ##
    x = 0
    y = 0
    clerk_list = []
    path2 = 'C:/Users/LMAN/Desktop/PYTHON TEST/exclT/ExcelProcess/data/华远数据转存.xlsx'
    path1 = 'C:/Users/lin/Desktop/ExcelProcess/data/华远数据转存.xlsx'
    workbook2 = xlrd.open_workbook(path2)
    sheet2 = workbook2.sheet_by_index(0)
    # print(sheet2.name, sheet2.nrows, sheet2.ncols)
    rows2 = sheet2.row_values(0)
    clerk_order_dic = {}
    for i in range(2, sheet2.nrows):
        #print(sheet2.row_values(i)[2])
        clerk_list.append(sheet2.row_values(i)[2])
        clerk_order_dic.setdefault(sheet2.row_values(i)[2], 0)
        #print(len(clerk_list))
    store_list = ['九一手机城', '晋江泉安店', '九一华为专卖店', '南安新华厅', '商务手机城', '九一oppo专卖店', \
                          '华远田安店', '中骏世界城华为体验店', '华远浮桥店', '永春华为店', '泉港南埔厅', '智天智能oppo体', \
                          '智能体验城', '晋江泉安新店', '九一百源店', '智天智能水头店', '洛江双阳店', '南安官桥店']


    sum = 0
    for i in clerk_order_dic.values():
        sum += i
    # print("sum : ",sum)
    # print(clerk_order_dic)


    for i in range(2, sheet2.nrows):
        #print(sheet2.row_values(i)[2])
        if sheet2.row_values(i)[2] in clerk_order_dic.keys():
            if sheet2.row_values(i)[3] != clerk_order_dic[sheet2.row_values(i)[2]]:
                a =1
                #print("#####下面是问题单#####")
            #print('名字',   sheet2.row_values(i)[2],  '统计的',   sheet2.row_values(i)[3],  '原生数据里的',   clerk_order_dic[sheet2.row_values(i)[2]])
    print(x)
    return HY


