def print_row_cells(sheet, min, max):
    for row in sheet.iter_rows(min_row=min, max_row=max, values_only=True):
        print(row)


# gets an entire row of cells
def get_row(sheet, row_number):
    cells = []
    for col in sheet.iter_cols(max_row=1):
        for cell in col:
            cells.append(cell)
    return cells
