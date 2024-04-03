import openpyxl as xl
from openpyxl.chart import BarChart, LineChart, Reference

def process_worksheet(input_filename, output_filename, discount_percentage, graph_type, x_column, y_column):
    wb = xl.load_workbook(input_filename)
    sheet1 = wb['Sheet1']

    discount_factor = 1 - (discount_percentage / 100)  # Convert discount percentage to discount factor

    for row in range(2, sheet1.max_row + 1):
        cell = sheet1.cell(row, 4)
        discounted_price = cell.value * discount_factor
        discounted_price_cell = sheet1.cell(row, 6)
        discounted_price_cell.value = discounted_price

    # Set the title "Discounted price" for the 6th column (assuming first row contains headers)
    sheet1.cell(row=1, column=6).value = "Discounted price"

    x_values = Reference(sheet1, min_row=2, max_row=sheet1.max_row, min_col=x_column, max_col=x_column)
    y_values = Reference(sheet1, min_row=2, max_row=sheet1.max_row, min_col=y_column, max_col=y_column)

    if graph_type == 'bar':
        chart = BarChart()
    elif graph_type == 'line':
        chart = LineChart()

    chart.add_data(y_values)
    chart.set_categories(x_values)
    chart.x_axis.title = 'Sales Volume'  # Set x-axis label
    chart.y_axis.title = 'Profit'  # Set y-axis label
    sheet1.add_chart(chart, 'G2')

    wb.save(output_filename)

# Rest of the code remains the same

# Ask the user to input the initial file name
input_file_name = input("Enter the name of the initial Excel file (with extension .xlsx): ")

# Ask the user to input the new file name for the modified file
output_file_name = input("Enter the name of the new Excel file to save (with extension .xlsx): ")

# Ask the user to input the discount percentage
discount_percentage = float(input("Enter the discount percentage (e.g., 10 for 10%): "))

# Ask the user to choose the graph type
graph_type = input("Choose the type of graph (bar/line): ").lower()

# Validate user input for graph type
while graph_type not in ['bar', 'line']:
    print("Invalid input! Please choose either 'bar' or 'line'.")
    graph_type = input("Choose the type of graph (bar/line): ").lower()

# Ask the user to input the x-axis column number
x_column = int(input("Enter the column number for the x-axis data: "))

# Ask the user to input the y-axis column number
y_column = int(input("Enter the column number for the y-axis data: "))

# Call the process_worksheet function with the provided parameters
process_worksheet(input_file_name, output_file_name, discount_percentage, graph_type, x_column, y_column)
