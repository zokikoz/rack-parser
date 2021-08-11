#!/usr/bin/env python3
# Matching inventory file with SM ID

import sys
import csv
import os

dc_invent_db = []
sm_id_db = []

def check_row(dc_row, sm_row, sm_id, type, count):
    if sm_id != '\\N':
        if dc_row.lower() == sm_row.lower():
            count['total'] += 1
            print(f'\r{sm_id}: found {type} {sm_row} in {dc_row} '
                f"({count['total']} matches: {count['serial']} by serial, {count['label']} by label)", end='')
            print(' '*5, end='') if os.name == 'nt' else print('\033[K', end='')
            sys.stdout.flush()
            if type == 'serial':
                count['serial'] += 1
            else:
                count['label'] += 1
            return True
    return False

if not sys.argv[2:]:
    print('Usage: match-sm-id.py dc_invent.csv sm_id.csv')
    sys.exit()

with open(sys.argv[1], mode='r') as csv_file:
    count = 0
    print('Loading DC inventory ', end='')
    csv_reader = csv.DictReader(csv_file, delimiter=';')
    for row in csv_reader:
        count +=1 
        dc_invent_db.append(row)
    print(f'({count})')

with open(sys.argv[2], mode='r') as csv_file:
    count = 0
    print('Loading SM ID ', end='')
    csv_reader = csv.DictReader(csv_file, delimiter=';')
    for row in csv_reader:
        count +=1
        if not row['sn']: row['sn'] = 'empty'
        if not row['dev_name']: row['dev_name'] = 'empty'
        if row['dev_name'] == 'N': row['dev_name'] = 'empty'
        sm_id_db.append(row)
    print(f'({count})')

count = {'total': 0, 'serial': 0, 'label': 0}

with open('result-smid.csv', 'w') as csv_file:
    ln = 0
    field_names = list(dc_invent_db[0].keys())
    field_names.append('SM ID')
    csv_writer = csv.DictWriter(csv_file, delimiter=';', fieldnames=field_names)
    csv_writer.writeheader()
    for dc_row in dc_invent_db:
        for sm_row in sm_id_db:
            if check_row(dc_row['S/N'], sm_row['sn'], sm_row['sm_id'], 'serial', count) \
            or check_row(dc_row['Label'], sm_row['dev_name'], sm_row['sm_id'], 'label', count):
                dc_row['SM ID'] = sm_row['sm_id']
                break
        csv_writer.writerow(dc_row)
        ln += 1
    print(f'\nDone ({ln} lines result)')
