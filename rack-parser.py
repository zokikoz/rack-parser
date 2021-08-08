#!/usr/bin/env python3
# Rack parser to flat file

import re
import json
import argparse
from openpyxl import load_workbook

# Search string for rack id (format: AA1.DC1.NN)
# AA - Address
# DC - Data center
# NN - Rack number
rack_id_regex = r'([A-Z][A-Z]\d)\.([A-Z][A-Z]\d)\.(\w\w)'
address_book = {}

parser = argparse.ArgumentParser(description='Converts rack unit view to flat inventory file')
parser.add_argument('source', type=argparse.FileType('rb'),
                    help='Rack view XLSX file')
parser.add_argument('-x', '--row', default=1000, metavar='X', type=int,
                    help='Maximum row to scan (default 1000)')
parser.add_argument('-y', '--column', default=1000, metavar='Y', type=int,
                    help='Maximum column to scan (default 1000)')
parser.add_argument('-b', '--buffer', default=100, metavar='N', type=int,
                    help='Break after N empty cells (default 100)')
parser.add_argument('-a', '--addr', metavar='json', type=argparse.FileType('r'),
                    help='Address book JSON file')
args = parser.parse_args()

def is_num(n):
    """Checks value is an integer"""
    try:
        if n and n is not True:
            int(n)
            return True
        else: return False
    except ValueError:
        return False

def is_merge(row, column):
    """Checks cell is merged"""
    cell = ws.cell(row, column)
    for merged_cell in ws.merged_cells.ranges:
        if (cell.coordinate in merged_cell):
            return True
    return False

def bottom_border(x, y, size):
    """Checks that the bottom border is found"""
    # First merged cell always has a border, so they need to be ignored
    if is_merge(x, y) is True and size == 1:
        # It can lead to an error if cell only horizontal merged. Need to figure it out
        return False
    elif is_merge(x,y) is False and ws.cell(row=x, column=y).border.bottom.style:
        return True
    elif is_merge(x,y) is True and ws.cell(row=x, column=y).border.bottom.style:
        return True
    else:
        return False

def search_rack(rack_x, rack_y):
    """Searching devices from rack id coordinates"""
    for x in range(rack_x, rack_x+60):
        value = ws.cell(row=x, column=rack_y-1).value
        # Finding units count
        if is_num(value):
            rack_unit = value
            label = get_label(x, rack_y)
            if label:
                vendor = get_info(x, rack_y+1)
                model = get_info(x, rack_y+2)
                serial = get_info(x, rack_y+3)
                dev = prepare_device(vendor, model, serial)
                if dev:
                    print(f"{data_center} - {address} - {dev['model']} - {dev['serial']} - {label['name']} - {rack_num} - {rack_unit} - {label['size']}")
            # Stoping on last unit
            if int(value) == 1: break 

def get_label(x, y):
    """Getting label name and size"""
    # Checks that unit has top border in label, this means the device starts here
    if ws.cell(row=x, column=y).border.top.style:
        label = {'size': 1, 'name': ''}
        for x in range(x, x+40):
            value = ws.cell(row=x, column=y).value
            if value:
                label['name'] += str(value)
            if not bottom_border(x, y, label['size']): 
                label['size'] += 1
            else: 
                break
        if label['name']: return label
        return False
    else: 
        return False

def get_info(x, y):
    """Getting label attributes"""
    #Getting to the top border. Sometimes different labels can have one attribute for all, located above 
    if not ws.cell(row=x, column=y).border.top.style:
        for x in reversed(range(x-40, x)):
            if ws.cell(row=x, column=y).border.top.style:
                break
    #Getting down
    for x in range(x, x+40):
        size = 1
        value = ws.cell(row=x, column=y).value
        if value:
            return value
        if not bottom_border(x, y, size):
            size += 1
        else: 
            return False
    return False

def prepare_device(vendor, model, serial):
    if vendor or model or serial:
        rack_device = {}
        if not vendor: vendor = ''
        if not model: model = ''
        if not serial: serial = ''
        rack_device['model'] = str(vendor).strip() + ' ' + str(model).strip()
        rack_device['serial'] = str(serial).strip()
        return rack_device
    else:
        return False

def set_address(addr_id):
    for id, address in address_book.items():
        if id == addr_id:
            return address
    return ''


if args.addr:
    with open(args.addr.name) as json_file:
        address_book = json.load(json_file)

wb = load_workbook(args.source)
for ws in wb:
    # Reading all worksheets in book
    print(ws.title)
    break_count_row = 1
    for x in range(1, args.row):
        row_not_empty = False
        break_count_col = 1
        for y in range(1, args.column):
            value = ws.cell(row=x, column=y).value
            if value:
                row_not_empty = True
                # Search for rack id
                output = re.search(rack_id_regex, str(value))
                if output:
                    data_center = 'ЦОД-' + output.group(1) + '_' + output.group(2)
                    address = set_address(output.group(1))
                    rack_num = output.group(3)
                    print(output.group(0))
                    search_rack(x, y)
            else:
                break_count_col += 1
                if break_count_col > args.buffer: break
        # Checks for an empty stop buffer
        break_count_row += 1
        if row_not_empty: break_count_row = 1
        if break_count_row > args.buffer: break
