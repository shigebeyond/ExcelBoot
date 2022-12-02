# Pandas与openpyxl库 结合
# https://blog.csdn.net/weixin_41261833/article/details/120169556

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment

data = {
    "姓名": ["张三", "李四"],
    "性别": ["男", "女"],
    "年龄": [15, 25],
}
df = pd.DataFrame(data)


wb = Workbook()
ws = wb.active

# df 转 openpyxl
for row in dataframe_to_rows(df, index=False, header=True):
    ws.append(row)

wb.save("pandas.xlsx")