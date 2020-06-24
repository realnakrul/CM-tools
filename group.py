from sys import argv
import tools
from pprint import pprint
from openpyxl import Workbook
from openpyxl import load_workbook


def main():
    if tools.check_input(argv[1]):
        matrix_input = argv[1]
        matrix_output = f'grp_{matrix_input}'
        print(f'Output file: {matrix_output}')
        # matrix = pandas.read_excel(matrix_input)
        # print(matrix)
        wb = load_workbook(matrix_input)
        sheet = tools.select_sheet(wb)
        # print(sheet.title)
        # print(sheet.max_column)
        # print(sheet.max_row)
        raw_matrix_list = tools.read_sheet(sheet)
        clean_matrix_list = tools.clean_list(raw_matrix_list)
        devices = tools.get_devices(clean_matrix_list)
        # pprint(clean_matrix_list)
        # tools.split_interfaces([0, 1], clean_matrix_list)
        group_list = tools.group_by_device(devices, clean_matrix_list)
        split_list = tools.split_interfaces([0, 1], clean_matrix_list)
        tools.write_to_excel(matrix_output, sheet.title, split_list)


if __name__ == '__main__':
    main()
