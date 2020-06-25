from os import path
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
    # for row in range(0, sheet.nrows):
    for row in sheet.iter_rows():
        row_list = []
        # print(row)
        # for cell in sheet.row(row):
        for cell in row:
            # print(cell.value)
            row_list.append(cell.value)
        matrix_list.append(row_list)
    # for row in wb.sheet_by_index(0).rows
    return matrix_list


def clean_list(raw_matrix, b_index: int, header=True):
    """
    Removes all matrix entries without B device name (Device A open interfaces)
    :param raw_matrix: Connectivity matrix with Device A open interfaces and header
    :param b_index: B device column in connectivity matrix
    :param header: Specifies if connectivity matrix has header in first row
    :return: Connectivity matrix without header and Device A open interfaces
    """
    clean_matrix = []
    if header:
        _ = raw_matrix.pop(0)
    for line in raw_matrix:
        if line[b_index]:
            clean_matrix.append(line)
    return clean_matrix


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


def split_interfaces(device_columns: list, matrix):
    """
    Splits single list item with device and interface into two list items, assuming that interface is a substring after
    last space.
    :param device_columns: is a list of indexes - specifies which list items to split
    """
    result = []
    for line in matrix:
        new_line = []
        for ind, cell in enumerate(line):
            if ind in device_columns:
                *device, interface = cell.split()
                new_line.extend([' '.join(device), interface])
            else:
                new_line.append(cell)
        result.append(new_line)
    return result


def populate_b(matrix):
    """
    Populates B device SFP, patch cord and rack from REVERSE record
    :param matrix: Clean connectivity matrix in FORWARD and REVERSE (Engineer) format without B SFP, Patch cord and rack
    :return: Connectivity matrix in FORWARD and REVERSE format with populated B SFP, Patch cord and rack
    """
    result = []
    for a_line in matrix:
        ab_line = []
        a_line_a_name = a_line[0]
        a_line_a_interface = a_line[1]
        a_line_b_name = a_line[2]
        a_line_b_interface = a_line[3]
        a_line_a_sfp = a_line[4]
        a_line_a_patch = a_line[5]
        a_line_a_rack = a_line[6]
        a_line_comment = a_line[7]
        # print('A: ', a_line)
        for b_line in matrix:
            b_line_a_name = b_line[0]
            b_line_a_interface = b_line[1]
            # b_line_b_name = b_line[2]
            # b_line_b_interface = b_line[3]
            b_line_a_sfp = b_line[4]
            b_line_a_patch = b_line[5]
            b_line_a_rack = b_line[6]
            # b_line_comment = b_line[7]
            if a_line_b_name == b_line_a_name and \
                    a_line_b_interface == b_line_a_interface:
                # print('B: ', b_line)
                ab_line = [a_line_a_name, a_line_a_interface, a_line_b_name, a_line_b_interface,
                           a_line_a_sfp, a_line_a_patch, a_line_a_rack,
                           b_line_a_sfp, b_line_a_patch, b_line_a_rack, a_line_comment]
        if not ab_line:
            ab_line = [a_line_a_name, a_line_a_interface, a_line_b_name, a_line_b_interface,
                       a_line_a_sfp, a_line_a_patch, a_line_a_rack,
                       '', '', '', a_line_comment]
        result.append(ab_line)
        # print('C: ', ab_line)
    return result
