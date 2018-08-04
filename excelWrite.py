dan1 = {'shop': '1_shop', 'date':'2018-7-22', 'amount':1}
dan2 = {'shop': '2_shop', 'date':'2018-7-23', 'amount':2}
dan3 = {'shop': '3_shop', 'date':'2018-7-24', 'amount':3}
total_dan = {'total_amount': 0}
member1 = [dan1, dan2]
member1.append(dan3)
member1.append(total_dan)
dic = {'num1':member1}
if 'num2' in dic:
    print (dic)
else:
    dic.setdefault('num2', dan2)
    print(dic)
    dic['num2'].append()