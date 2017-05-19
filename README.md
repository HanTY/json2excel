# json2excel
解决每次需要把json转换为excel的痛苦。支持的格式：
case a:
a={
  'sheet1':[
    {'name':'123', 'age':18},
    {'name':'456', 'age':20},
  ]
}
case b:
b={
  'sheet1':{"2017-02-19":
    [
      {'name':'123', 'age':18},
      {'name':'456', 'age':20},
    ]
  }
}
sheet,为sheet名称。
结果为a.xlsx 和 b.xlsx
