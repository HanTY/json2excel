# -*- coding: utf-8 -*-
"""
json 到excel 的转换程序
"""
import xlsxwriter
import datetime
import pytz
import json
from decimal import Decimal


class Json2xlsx(object):
    def __init__(self, data=None, filename=None):
        if data is None:
            data = {}
        if not isinstance(data, dict):
            raise 'data type is not a dict'
        if filename is None:
            filename = ''
        self.filename = filename
        self.data = data
        self.xlsx = self._xlsx()
        # Add a number format for cells with money.
        # self.money_format = self.xlsx.add_format({'num_format': '#,##0'})
        # Add an Excel date format.
        self.date_format = self.xlsx.add_format({'num_format': 'yyyy-mm-dd'})


    def transform(self):
        for k, v in self.data.items():
            tmp_worksheet = self.xlsx.add_worksheet(k[:30])
            self.write_sheet(tmp_worksheet, v)
        return True

    def xlsx_close(self):
        self.xlsx.close()

    def _xlsx(self):
        workbook = xlsxwriter.Workbook(self.filename)
        return workbook

    def write_sheet(self, sheet, data):
        data = self._serialize_data(data)
        if data:
            for row, item in enumerate(data):
                for col, value in enumerate(item):
                    if isinstance(value, datetime.datetime):
                        sheet.write_datetime(row, col, value, self.date_format)
                    elif isinstance(value, datetime.date):
                        sheet.write_datetime(row, col, value, self.date_format)
                    elif isinstance(value, str):
                        sheet.write_string(row, col, value)
                    elif isinstance(value, unicode):
                        sheet.write_string(row, col, value)
                    elif isinstance(value, Decimal):
                        sheet.write_number(row, col, value)
                    elif value is None:
                        value = ''
                        sheet.write_string(row, col, value)
                    else:
                        sheet.write_number(row, col, value)
            return True
        else:
            return False

    def _serialize_data(self, data):
        if isinstance(data, dict):
            return self._serialize_dict_data(data)
        elif isinstance(data, list) or isinstance(data, tuple):
            return self._serialize_list_data(data)
        else:
            return None

    def _serialize_dict_data(self, data, first_column_name=u'日期'):
        result = []
        columns = [first_column_name]
        for k, v in data.items():
            if isinstance(v, list) or isinstance(v, tuple):
                for item in v:
                    for item_k in item.keys():
                        if item_k not in columns:
                            columns.append(item_k)
            elif isinstance(v, dict):
                for item_k in item.keys():
                    if item_k not in columns:
                        columns.append(item_k)
        result.append(columns)
        for k, v in data.items():
            if isinstance(v, list) or isinstance(v, tuple):
                if v:
                    for item in v:
                        tmp = []
                        tmp.append(k)
                        for query in columns[1:]:
                            tmp.append(item.get(query, 0))
                        result.append(tmp)
                else:
                    tmp = []
                    tmp.append(k)
                    for query in columns[1:]:
                        tmp.append(0)
                    result.append(tmp)
            elif isinstance(v, dict):
                tmp = []
                tmp.append(k)
                for query in columns[1:]:
                    tmp.append(v.get(query, 0))
                result.append(tmp)
        return result

    def _serialize_list_data(self, data):
        result = []
        columns = []
        for item in data:
            for k in item.keys():
                if k not in columns:
                    columns.append(k)
        result.append(columns)
        for item in data:
            tmp = []
            for i in columns:
                tmp.append(item.get(i, 0))
            result.append(tmp)
        return result


# if __name__ == '__main__':
#     a=open('outs_20170518.json')
#     data = json.load(a)
#     a.close()
#     b=Json2xlsx(data, filename='test.xlsx')
#     b.transform()
#     b.xlsx_close()
