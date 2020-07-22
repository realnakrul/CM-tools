from sys import argv
import tools
from openpyxl import Workbook
from openpyxl import load_workbook
from draw_network_graph import draw_topology
import os
from pprint import pprint

os.environ["PATH"] += os.pathsep + r'C:\Program Files (x86)\Graphviz2.38\bin'

ACTIONS = ['Enforce Engineer format and get summary', 'Enforce Technician format', 'Group by device']


def select_action(avail_actions: list):
    print('Available actions (matrix will be cleaned and validated for any action):')
    result = ''
    for ind, action in enumerate(avail_actions, 1):
        print(f'{ind}. {action}')
        correct = False
    while not correct:
        try:
            action_ind = int(input(f'Select action (1 - {len(avail_actions)}): '))
            if action_ind <= len(avail_actions):
                result = avail_actions[action_ind - 1]
                correct = True
        except ValueError:
            print('Select action number')
        except IndexError:
            print(f'Select action number in range (1 - {len(avail_actions)})')
        except Exception as e:
            print(type(e), e)
    return result


def main():
    if tools.check_src_file(argv[1]):
        src_file = argv[1]
        dst_file = f'cmt_{src_file}'
        print(f'Output file: {dst_file}')
        tools.check_dst_file(dst_file)
        print('-'*80)
        src_book = load_workbook(src_file)
        dst_book = Workbook(write_only=True)
        sheets = tools.select_sheet(src_book)
        print('-' * 80)
        action = select_action(ACTIONS)
        print(action)
        for sheet in sheets:
            print('-'*80)
            print(f'Current worksheet: {sheet.title}')
            print(f'Action: {action}')
            raw_matrix_list = tools.read_sheet(sheet)
            clean_matrix_list = tools.clean_list(raw_matrix_list)
            devices = tools.get_unique_values(clean_matrix_list, [0, 2])
            if not tools.consistency_check(clean_matrix_list, devices):
                racks = tools.get_unique_values(clean_matrix_list, [6, 9])
                if action == 'Enforce Engineer format and get summary':
                    clean_matrix_list = tools.engineer_format(clean_matrix_list)
                    summary_list = tools.rack_to_rack_summary(racks, clean_matrix_list)
                    dst_book = tools.add_to_sheet(dst_book, sheet.title, clean_matrix_list, 'connectivity')
                    dst_book = tools.add_to_sheet(dst_book, sheet.title + ' Sum', summary_list, 'summary')
                elif action == 'Enforce Technician format':
                    clean_matrix_list = tools.technician_format(clean_matrix_list)
                    topology = tools.get_topoly(clean_matrix_list)
                    pprint(topology)
                    '''topology = {('CR SW 01', 'MGMT01'): ('MGMT SW 01', 'E1/1'),
                                ('CR SW 02', 'MGMT01'): ('MGMT SW 01', 'E1/2')}'''
                    draw_topology(topology)
                    dst_book = tools.add_to_sheet(dst_book, sheet.title, clean_matrix_list, 'connectivity')
                elif action == 'Group by device':
                    clean_matrix_list = tools.group_by_device(devices, clean_matrix_list)
                    dst_book = tools.add_to_sheet(dst_book, sheet.title, clean_matrix_list, 'connectivity')
            else:
                print('Processing impossible, skipping')
        print('-'*80)
        if dst_book.sheetnames:
            print(f'Saving result workbook as {dst_file}')
            dst_book.save(dst_file)
            tools.add_filters(dst_file)


if __name__ == '__main__':
    main()
