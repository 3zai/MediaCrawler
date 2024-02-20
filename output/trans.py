import json
import glob
import xlwt

file_list = glob.glob("../data/xhs/*contents*.json")
wb = xlwt.Workbook()

for i in range(len(file_list)):
    ws = wb.add_sheet(f"Sheet{i + 1}",cell_overwrite_ok=True)
    with open(file_list[i]) as f:
        data = json.load(f)
        row_0_col_index = 0
        for key, val in data[0].items():
            ws.write(0, row_0_col_index, key)
            row_0_col_index += 1

        for j in range(len(data)):
            row_index = j + 1
            col_index = 0
            for k in data[j]:
                ws.write(row_index, col_index, data[j][k])
                col_index += 1

wb.save("111.xls")
