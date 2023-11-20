from openpyxl import load_workbook

def load_excel_data(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    max_row = ws.max_row
    all_rows = []

    # Iterate over rows
    for row in range(2, max_row):
        # Create an array for the current row
        current_row = []

        # Iterate over columns from B to Y
        for col in range(1, 26):  # 26 corresponds to the column index for Y
            cell_value = ws.cell(row=row, column=col).value
            current_row.append(cell_value)

        # Append the current row to the list of all rows
        all_rows.append(current_row)

    return all_rows

if __name__ == "__main__":
    file_path = 'data.xlsx'
    data = load_excel_data(file_path)

    # Print or use the array of rows as needed
    for row in data:
        print(row)
