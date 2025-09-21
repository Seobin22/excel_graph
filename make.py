from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.axis import DateAxis
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
import os
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.axis import NumericAxis
from openpyxl.chart import BarChart


dir='C:\\Users\\Owner\\Desktop\\graph_save\\'
filename='perm'+'.xlsx'
if os.path.exists(dir+filename):
    os.remove(dir+filename)
x='n'
y='실행시간'
#1. 데이터 (n, 실행시간)
data={x:[x for x in range(1,11)],y:[0,0,0,0,0,0.001,0.003,0.031,0.291,3.233]}
#data={x:[0.1*x for x in range(1000)],y:[x for x in range(1000)]}
df=pd.DataFrame(data)

# 2. 엑셀 워크북 및 워크시트 생성
wb = Workbook()
ws = wb.active
ws.title = "3차시"

for row in dataframe_to_rows(df, index=False, header=True):
    ws.append(row)
    

# 꺾은선형(Line) 차트
chart =LineChart()
chart.title = "순열 알고리즘의 n-실행시간 그래프"
# chart.style = 12                # 기본 스타일 (엑셀과 유사)
chart.y_axis.title = "실행 시간"
chart.x_axis.title = "n"
# --- 🎯 여기가 데이터 레이블을 추가하는 핵심 부분 ---
chart.dLbls = DataLabelList()
chart.dLbls.showVal = True              # 값 표시 (가장 일반적)
chart.dLbls.showSerName = False         # 계열 이름 표시 (선택)
chart.dLbls.showCatName = False         # 항목 이름 표시 (선택)
chart.dLbls.separator = ' '
chart.dLbls.dLblPos = 't'
chart.legend.position = 'b'# 구분자 설정 (선택)
# chart.y_axis = NumericAxis(title="실행 시간")
# chart.x_axis = NumericAxis(title="n")
# chart.x_axis.number_format = '0'        # X축은 정수로 표시
# chart.y_axis.number_format = '0.000'   # Y축은 소수점 3자리까지 표시

# 데이터 범위
mr=len(data[x])+1

x_ref = Reference(ws, min_col=1, min_row=2, max_row=mr)#x축
y_ref = Reference(ws, min_col=2, min_row=1, max_row=mr)#y축

chart.add_data(y_ref, titles_from_data=True)
chart.set_categories(x_ref)

# 범례(오른쪽의 계열 표시)를 없애는 코드
# chart.legend = None

# series = chart.series[0] 
# # 선 색을 파란색 계열(0070C0)로 설정
# series.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="0070C0")) 


# 시트에 그래프 추가
ws.add_chart(chart, "E5")


wb.save(dir+filename)

