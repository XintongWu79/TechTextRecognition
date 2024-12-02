import xlwings as xw
import jieba

app = xw.apps.active
wb = app.books.active
sht = wb.sheets.active

rng = sht.range('f2:f6029')
value_arr = rng.value

txt = ['科技','技术','网络', '电子','医药','医疗','药业','制药','研究','能源','金属','生物','生化','化工','化学','精密','精细','数字','信息','仪器','仪表','设备','环保','环境','工业','材料','水电','电力','光电','通信','软件','电器','石化','计算机']

result = []

for i in value_arr:
    if i is not None:
        i_str = i.encode('utf-8')
        i_jieba = jieba.lcut(i_str, cut_all=True)
        inter = set(txt).intersection(set(i_jieba))
        if len(inter) > 0:
            result.append(1)
        else:
            result.append(0)
    else:
        result.append('')

sht.range('l2').options(transpose=True).value = result