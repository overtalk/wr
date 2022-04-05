# -*- coding: UTF-8 -*-

import argparse
import openpyxl

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", type=str, required=False,
                default='./wr1.xlsx', help='excel path')
    args = parser.parse_args()

    print("[ARGS] excel(%s)" % (args.excel))

    workbook = openpyxl.load_workbook(filename=args.excel)

    sheet = workbook["Sheet1"]

    print(type(sheet.dimensions), sheet.dimensions)

    dimensions = str(sheet.dimensions)
    splited_data = dimensions.split(":")

    start = splited_data[0]
    end = splited_data[-1]
    print()
    raw_data = {}
    tmp_data_by_cat1 = {}
    cell = sheet["%s:%s"%(start, end)]

    datas = tuple(sheet.rows)
    for index in range(len(datas)):
        items = datas[index]
        for item in items:
            print(item.value)
        break

    # for i in cell:
    #     for j in i:
    #         print(j.value)
