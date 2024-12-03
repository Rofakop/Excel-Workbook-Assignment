import openpyxl
from openpyxl.chart import BarChart, Reference, Series

# Create a new workbook and select the active sheet
wb = openpyxl.Workbook()
sheet = wb.active

# Data to be added to the Excel sheet
data = {"Agricultural vehicle": 11938, "Cars": 12911, "Bus": 11636, "Van": 11394, "Bike": 17456, "Others": 15140}

# Add headers and data to the sheet
sheet.append(["Vehicle Type", "Count"])  # Adding header row
for key, value in data.items():
    sheet.append([key, value])  # Appending rows as [key, value]

# Create a reference object for the chart data
data_ref = Reference(sheet, min_col=2, min_row=2, max_col=2, max_row=len(data))
categories_ref = Reference(sheet, min_col=1, min_row=2, max_row=len(data))

# Create the chart
chartObj = BarChart()
chartObj.title = 'Vehicle Accident Data'
chartObj.x_axis.title = 'Vehicle Type'
chartObj.y_axis.title = ' Accident Count'
chartObj.add_data(data_ref, titles_from_data=False)
chartObj.set_categories(categories_ref)

# Add the chart to the sheet
sheet.add_chart(chartObj, 'C10')

# Save the workbook
wb.save('Excel/Paul_visualization.xlsx')
print('Done')



