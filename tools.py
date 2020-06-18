from os import path
import xlrd
import xlwt


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
    for ind, sheet in enumerate(book.sheet_names(), 1):
        print(f'{ind}. {sheet}')
    correct = False
    while not correct:
        try:
            sheet_ind = int(input(f'Select sheet (1 - {len(book.sheet_names())}): '))
            sheet = book.sheet_by_index(sheet_ind - 1)
            correct = True
        except ValueError:
            print('Select sheet number')
        except IndexError:
            print(f'Select sheet number in range (1 - {len(book.sheet_names())})')
        except Exception as e:
            print(type(e), e)
    return sheet


def read_sheet(sheet):
    matrix_list = []
    row_list = []
    for row in range(0, sheet.nrows):
        row_list = []
        # print(sheet.row(row))
        for cell in sheet.row(row):
            # print(cell.value)
            row_list.append(cell.value)
        matrix_list.append(row_list)
    # for row in wb.sheet_by_index(0).rows
    return matrix_list


def clean_list(raw_list):
    clean = []
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
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    for line_ind, line in enumerate(data):
        for cel_ind, cell in enumerate(line):
            sheet.write(line_ind, cel_ind, cell)
    workbook.save('grp.xls')


def get_devices(matrix):
    result = set()
    for line in matrix:
        *device, _ = line[0].split()
        result.add(' '.join(device))
    print(f'{len(result)} devices found:')
    for d in result:
        print(d)
    return list(result)





