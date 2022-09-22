import openpyxl
from openpyxl.chart import LineChart, Reference


workbook = openpyxl.load_workbook('file/data.xlsx')
sheet = workbook['薪水']
# 创建折线图的图标对象
chart = LineChart()
# 数据的引用范围
data = Reference(worksheet=sheet, min_row=2, max_row=5, min_col=1, max_col=13)
# 类别的引用范围 min_row-> 开始行号， max_row-> 结束行号， min_col-> 开始列， max_col-> 结束列
categories = Reference(sheet, min_row=1, min_col=2, max_col=13)
# 将数据与类别添加到图标当中
chart.add_data(data, from_rows=True, titles_from_data=True)
chart.set_categories(categories)
# 将图表插入到工作表中，从A8列开始插入图表
sheet.add_chart(chart, 'A8')
workbook.save('data.xlsx')
workbook.close()
