def print_row_cells(sheet, min, max):
    for row in sheet.iter_rows(min_row=min, max_row=max, values_only=True):
        print(row)


# gets an entire row of cells
# A1,B1,C1,D1,...
def get_row(sheet, row_number):
    cells = []
    for row in sheet.iter_cols(max_row=1):
        for cell in row:
            cells.append(cell)
    return cells

# gets an entire column of cells
# A1,A2,A3,A4,...


def get_column_cells(sheet, column_number):
    arr_of_cells = []
    for cells in sheet.iter_cols(min_col=column_number, max_col=column_number):
        for cell in cells:
            arr_of_cells.append(cell.value)
    return arr_of_cells


# build a dictionary from a list of cells
# maping of dictionary is value ==> column number
def build_dictonary_from_cell_array(cell_array):
    cell_dict = {}
    for cell in cell_array:
        cell_dict[cell.value] = cell.column
    return cell_dict


# append an array of data cells, to a specified row and column in the provided
# sheet
def append_data_to_sheet(sheet, data_array, col_to_append_to, starting_row):
    for cell in data_array:
        sheet.cell(row=starting_row, column=col_to_append_to).value = cell
        starting_row += 1
