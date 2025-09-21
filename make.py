from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference

# 새 워크북과 시트
wb = Workbook()
ws = wb.active
ws.title = "순열 실행시간"

# 데이터 (n, 실행시간)
data = [
    ["n", "실행시간"],
    [1, 0],
    [2, 0],
    [3, 0],
    [4, 0],
    [5, 0],
    [6, 0.001],
    [7, 0.003],
    [8, 0.031],
    [9, 0.291],
    [10, 3.233],
]
for row in data:
    ws.append(row)

# 꺾은선형(Line) 차트
chart = LineChart()
chart.title = "순열 알고리즘의 n-실행시간 그래프"
chart.style = 2                # 기본 스타일 (엑셀과 유사)
chart.y_axis.title = "실행 시간"
chart.x_axis.title = "n"

# 데이터 범위
data_ref = Reference(ws, min_col=2, min_row=1, max_row=11)
cats_ref = Reference(ws, min_col=1, min_row=2, max_row=11)

chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)

# 시트에 그래프 추가
ws.add_chart(chart, "E5")

wb.save("순열_그래프.xlsx")

