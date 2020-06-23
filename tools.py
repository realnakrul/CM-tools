from os import path
import xlrd
# import xlwt
from openpyxl import Workbook


def check_input(file):
    result = path.exists(file)
    print(f'Input file: {file}')
    if result:
        print('Found.')
    else:
        print('Not found.')
    return result


def select_sheet(book):
    print('Available sheets: ')
    for ind, sheet in enumerate(book.sheetnames, 1):
        print(f'{ind}. {sheet}')
    correct = False
    while not correct:
        try:
            sheet_ind = int(input(f'Select sheet (1 - {len(book.sheetnames)}): '))
            # sheet = book.sheet_by_index(sheet_ind - 1)
            sheet = book.worksheets[sheet_ind - 1]
            correct = True
        except ValueError:
            print('Select sheet number')
        except IndexError:
            print(f'Select sheet number in range (1 - {len(book.sheetnames)})')
        except Exception as e:
            print(type(e), e)
    return sheet


def read_sheet(sheet):
    matrix_list = []
    row_list = []
    #for row in range(0, sheet.nrows):
    for row in sheet.iter_rows():
        row_list = []
        # print(row)
        #for cell in sheet.row(row):
        for cell in row:
            # print(cell.value)
            row_list.append(cell.value)
        matrix_list.append(row_list)
    # for row in wb.sheet_by_index(0).rows
    return matrix_list


def clean_list(raw_list, header=True):
    clean = []
    if header:
        _ = raw_list.pop(0)
    for line in raw_list:
        if line[1]:
            clean.append(line)
    return clean


def group_by_device(devices, matrix, add_name=True):
    result = []
    print(f'Grouping by device with device header = {add_name}')
    for device in devices:
        if add_name:
            result.append([device])
        for line in matrix:
            if device in line[0]:
                result.append(line)
    return result


def write_to_excel(file, sheet_name, data):
    book = Workbook(write_only=True)
    book.create_sheet(sheet_name)
    sheet = book[sheet_name]
    '''for line_ind, line in enumerate(data, 1):
        for cel_ind, cell in enumerate(line, 1):
            print(line_ind, cel_ind, cell)
            sheet.cell(line_ind, cel_ind, cell)'''
    for line in data:
        sheet.append(line)
    book.save(file)


def get_devices(matrix):
    result = set()
    for line in matrix:
        *device, _ = line[0].split()
        result.add(' '.join(device))
    print(f'{len(result)} devices found:')
    for d in result:
        print(d)
    return list(result)


def split_interfaces(matrix):
    result = []
    for line in matrix:
        print(line)
        new_line = []
        for ind, cell in enumerate(line):
            print(ind, cell)
            if ind == 0 or ind == 1:
                *device, interface = cell.split()
                new_line.extend([' '.join(device), interface])
                print('AAA', new_line)
            else:
                new_line.append(cell)
        print('new ', new_line)
        result.append(new_line)
    print(result)





