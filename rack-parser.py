#!/usr/bin/env python3
# Rack parser to flat file

import re
import sys
import argparse
from openpyxl import load_workbook

# Default scan range
x_deep = 100
y_deep = 100

# Rewrite later, false positive with True values
def is_num(n):
    try:
        if n:
            int(n)
            return True
        else: return False
    except ValueError:
        return False

def is_merge(row, column):
    cell = ws.cell(row, column)
    for mergedCell in ws.merged_cells.ranges:
        if (cell.coordinate in mergedCell):
            return True
    return False

def bottom_border(x,y,size):
    if is_merge(x,y) is True and size == 1:
        return False
    elif is_merge(x,y) is False and ws.cell(row=x, column=y).border.bottom.style:
        return True
    elif is_merge(x,y) is True and ws.cell(row=x, column=y).border.bottom.style:
        return True
    else:
        return False

def search_rack(rack_x, rack_y):
    #rack_devices = {}
    for x in range(rack_x,rack_x+60):
        value = ws.cell(row=x, column=rack_y-1).value
        if is_num(value):
            rack_unit = value
            label = get_label(x,rack_y)
            if label:
                vendor = get_info(x,rack_y+1)
                model = get_info(x,rack_y+2)
                serial = get_info(x,rack_y+3)
                if vendor or model or serial:
                    print(f"{rack_unit}: {label['size']} - {label['name']} - {vendor} {model} - {serial}")
            if int(value) == 1: break 

def get_label(x,y):
    if ws.cell(row=x, column=y).border.top.style:
        label = {'size': 1, 'name': ''}
        for x in range(x,x+20):
            value = ws.cell(row=x, column=y).value
            if value:
                label['name']+=value
            if not bottom_border(x,y,label['size']): 
                label['size']+=1
            else: 
                break
        return label
    else: 
        return False

def get_info(x,y):
    #Getting to the top border
    if not ws.cell(row=x, column=y).border.top.style:
        for x in reversed(range(x-10,x)):
            if ws.cell(row=x, column=y).border.top.style:
                break
    #Getting down
    for x in range(x,x+20):
        size = 1
        value = ws.cell(row=x, column=y).value
        if value:
            return value
        if not bottom_border(x,y,size):
            size+=1
        else: 
            return False
    return False

def xlsx_load():
    if sys.argv[1:]:
        wb = load_workbook(sys.argv[1])
        return wb
    else:
        sys.exit('Usage: rack-parser.py <filename.xlsx>')

wb = xlsx_load()
for ws in wb:
    # Reading all worksheets in book
    print(ws.title)
    for x in range(1,x_deep):
        for y in range(1,y_deep):
            value = ws.cell(row=x, column=y).value
            if value:
                output = re.search(r'[A-Z][A-Z]\d\.[A-Z][A-Z]\d\.\w+', str(value))
                if output:
                    print(value)
                    search_rack(x,y)

