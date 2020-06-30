from os import path
from openpyxl import Workbook
from collections import namedtuple


Link = namedtuple('Link', 'a_name, a_interface, b_name, b_interface, a_sfp, a_patch, a_rack, \
                           b_sfp, b_patch, b_rack, comment')


def check_input(file):
    print('-'*80)
    result = path.exists(file)
    print(f"Checking if input file \'{file}\' exist: {result}")
    return result


def select_sheet(book):
    sheets = []
    print('Available sheets: ')
    for ind, sheet in enumerate(book.sheetnames, 1):
        print(f'{ind}. {sheet}')
    print(f'{len(book.sheetnames)+1}. ALL')
    correct = False
    # print(book.worksheets)
    while not correct:
        try:
            sheet_ind = int(input(f'Select sheet (1 - {len(book.sheetnames) + 1}): '))
            if sheet_ind == len(book.sheetnames)+1:
                sheets = book.worksheets
                correct = True
            elif sheet_ind <= len(book.sheetnames)+1:
                sheets.append(book.worksheets[sheet_ind - 1])
                correct = True
        except ValueError:
            print('Select sheet number')
        except IndexError:
            print(f'Select sheet number in range (1 - {len(book.sheetnames)+1})')
        except Exception as e:
            print(type(e), e)
    return sheets


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


def clean_list(raw_matrix, b_index: int):
    """
    Removes all matrix entries without B device name (Device A open interfaces)
    :param raw_matrix: Connectivity matrix with Device A open interfaces and header
    :param b_index: B device column in connectivity matrix
    :param header: Specifies if connectivity matrix has header in first row
    :return: Connectivity matrix without header and Device A open interfaces
    """
    clean_matrix = []
    correct_input = False
    print('Cleaning matrix')
    print('Header row MUST be removed for processing (will be added back later)')
    print('First row:')
    print('| ', end='')
    for cell in raw_matrix[0]:
        print(cell, end=' | ')
    print()
    while not correct_input:
        answer = input('Is it header (y/n): ')
        if answer == 'y':
            correct_input = True
            header = True
        elif answer == 'n':
            correct_input = True
            header = False
        else:
            print('y or n')
    print(f'\tRemoving header = {header}')
    if header:
        _ = raw_matrix.pop(0)
    for line in raw_matrix:
        if line[b_index]:
            clean_matrix.append(line)
    print(f'\t{len(raw_matrix)-len(clean_matrix)} open interfaces removed')
    return clean_matrix


def group_by_device(devices: list, matrix: list, add_name=True):
    """
    Groups matrix rows by device name
    :param devices: Unique list of devices in the matrix
    :param matrix: Connectivity matrix in ENGINEER format
    :param add_name: Controls if device name required as a caption before group
    :return: Connectivity matrix grouped by device
    """
    result = []
    print('Grouping by device')
    print(f'\tSet device name as a header = {add_name}')
    for device in devices:
        if add_name:
            result.append([device])
        for line in matrix:
            if device in str(line[0]):
                result.append(line)
    return result


def add_to_sheet(book, sheet_name, data: list):
    """
    Adds excel sheet to workbook and fills it with data
    :param book: Excel workbook
    :param sheet_name: Excel sheet name
    :param data: List to write to excel sheet
    :return: Excel workbook with given sheet
    """
    header = ['Device A name', 'Device A interface', 'Device B name', 'Device B interface', 'Device A SFP',
              'Device A patch cord', 'Device A rack', 'Device B SFP', 'Device B patch cord', 'Device B rack', 'Comment']

    print('Adding header')
    data.insert(0, header)
    print(f"Adding excel sheet \'{sheet_name}\' to result workbook")
    book.create_sheet(sheet_name)
    sheet = book[sheet_name]
    for line in data:
        sheet.append(line)
    return book


def get_devices(matrix):
    result = set()
    for line in matrix:
        *device, _ = line[0].split()
        result.add(' '.join(device))
    '''print(f'{len(result)} devices found:')
    for d in result:
        print(d)'''
    return list(result)


def get_unique_values(matrix: list, index_columns: list):
    """
    Gets unique values from a given matrix column
    :param matrix: Connectivity matrix
    :param index_columns: List of column indexes to get unique values
    :return: List of unique values
    """
    result = set()
    for i in index_columns:
        for line in matrix:
            result.add(str(line[i]).strip())
    '''print(f'{len(result)} devices found:')
    for d in result:
        print(d)'''
    return list(result)


def split_interfaces(matrix, device_columns: list):
    """
    Splits single list item with device and interface into two list items, assuming that interface is a substring after
    last space.
    :param device_columns: is a list of indexes - specifies which list items to split
    :param matrix: Connectivity matrix with device and interface as one cell
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


def get_reverse(matrix: list, link_str: str):
    """
    Searches for REVERSE link for given link and matrix
    :param matrix: Clean connectivity matrix
    :param link_str: FORWARD link
    :return: REVERSE link
    """
    forward = Link(*link_str)
    result = ''
    for line in matrix:
        reverse = Link(*line)
        if forward.a_name == reverse.b_name and forward.b_name == reverse.a_name and \
                forward.a_interface == reverse.b_interface and forward.b_interface == reverse.a_interface:
            result = line
    return result


def engineer_format(matrix: list):
    """
    Creates ENGINEER format matrix from TECHNICIAN format
    :param: matrix: Clean connectivity matrix
    :return: Connectivity matrix in ENGINEER format
    """
    print('Creating ENGINEER format matrix')
    reverse_list = []
    result = []
    for line in matrix:
        forward = Link(*line)
        result.append(line)
        if not get_reverse(matrix, line):
            reverse_list.append([forward.b_name, forward.b_interface, forward.a_name, forward.a_interface,
                                 forward.b_sfp, forward.b_patch, forward.b_rack,
                                 forward.a_sfp, forward.a_patch, forward.a_rack, forward.comment])
    print(f'\t{len(reverse_list)} reverse connections added')
    if reverse_list:
        result += reverse_list
    # print(result)
    return result


def technician_format(matrix: list):
    """
    Creates TECHNICIAN format matrix from ENGINEER format
    :param: matrix: Clean connectivity matrix
    :return: Connectivity matrix in TECHNICIAN format
    """
    print('Creating TECHNICIAN format matrix')
    # reverse_list = []
    result = []
    for line in matrix:
        # forward = Link(*line)
        if not get_reverse(result, line):
            result.append(line)
    print(f'\t{len(matrix)-len(result)} reverse connections removed')
    return result


def consistency_check(matrix: list):
    """
    Checks if interfaces are not duplicated or linked to itself
    :param matrix: Clean connectivity matrix
    :return: List of warning if detected
    """
    result = []
    for first_ind, first_line in enumerate(matrix):
        first_link = Link(*first_line)
        if first_link.a_name == first_link.b_name and first_link.a_interface == first_link.b_interface:
            print(f'Warning: {first_link.a_name} {first_link.a_interface} links to itself. '
                  f'Line {first_ind}')
            result.append(f'Warning: {first_link.a_name} interface {first_link.a_interface} links to itself. '
                          f'Line {first_ind}')
        for second_ind, second_line in enumerate(matrix):
            second_link = Link(*second_line)
            if first_ind != second_ind and first_link.a_name == second_link.a_name and \
                    first_link.a_interface == second_link.a_interface:
                print(f'Warning: {first_link.a_name} {first_link.a_interface} duplicated. '
                      f'Lines {first_ind} and {second_ind}')
                result.append(f'Warning: {first_link.a_name} {first_link.a_interface} duplicated. '
                              f'Lines {first_ind} and {second_ind}')
            elif first_ind != second_ind and first_link.b_name == second_link.b_name and \
                    first_link.b_interface == second_link.b_interface:
                print(f'Warning: {first_link.b_name} {first_link.b_interface} duplicated. '
                      f'Lines {first_ind} and {second_ind}')
                result.append(f'Warning: {first_link.b_name} {first_link.b_interface} duplicated. '
                              f'Lines {first_ind} and {second_ind}')
    return result
