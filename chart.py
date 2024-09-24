from openpyxl.workbook import Workbook
from openpyxl import load_workbook
from openpyxl.chart import PieChart, PieChart3D, Reference, BarChart, BarChart3D, LineChart, LineChart3D

# load existing spreadsheet
wb = load_workbook('hello.xlsx')

# Create an active worksheet
ws = wb.active

# Determine type of chart
chart = BarChart3D()

# Designate labels and data
labels = Reference(ws, min_col=1, max_col=1, min_row=2, max_row=10)
data = Reference(ws, min_col=3, min_row=1, max_row=10)

# Put it all together
chart.add_data(data, titles_from_data=True)
chart.set_categories(labels)

# Add a title
chart.title = "Employee Salaries"

# Place the chart on the spreadsheet
ws.add_chart(chart, "E2")



# Save an excel spreadsheet
wb.save('BarChart3D.xlsx')
print("File was saved!")