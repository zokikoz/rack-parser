#!/usr/bin/env python3
# Rack parser to flat file

import re
from openpyxl import load_workbook

# Rewrite later, false positive with True values
def is_num(n):
    try:
        if n:
            int(n)
            return True
        else: return False
    except ValueError:
        return False

def is_merge(row, column, ws):
    cell = ws.cell(row, column)
    for mergedCell in ws.merged_cells.ranges:
        if (cell.coordinate in mergedCell):
            return True
    return False

def bottom_border(x,y,size,ws):
    if is_merge(x,y,ws) is True and size == 1:
        return False
    elif is_merge(x,y,ws) is False and ws.cell(row=x, column=y).border.bottom.style:
        return True
    elif is_merge(x,y,ws) is True and ws.cell(row=x, column=y).border.bottom.style:
        return True
    else:
        return False

def search_unit(rack_x, rack_y, ws):
    #rack_devices = {}
    for x in range(rack_x,rack_x+60):
        value = ws.cell(row=x, column=rack_y-1).value
        if is_num(value):
            rack_unit = value
            label = get_label(x,rack_y,ws)
            if label:
                vendor = get_vendor(x,rack_y+1,ws)
                print(f"{rack_unit}: {label['size']} - {label['name']} - {vendor}")
            if int(value) == 1: break 

def get_label(u_x,rack_y,ws):
    if ws.cell(row=u_x, column=rack_y).border.top.style:
        label = {'size': 1, 'name': ''}
        for u_x in range(u_x,u_x+20):
            value = ws.cell(row=u_x, column=rack_y).value
            if value:
                label['name']+=value
            if not bottom_border(u_x,rack_y,label['size'],ws): 
                label['size']+=1
            else: 
                break
        return label
    else: 
        return False

def get_vendor(x,y,ws):
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
        if not bottom_border(x,y,size,ws):
            size+=1
        else: 
            return False
    return False

wb = load_workbook('./rack template.xlsx')
for ws in wb:
    print(ws.title)
    for x in range(1,61):
        for y in range(1,21):
            value = ws.cell(row=x, column=y).value
            if value:
                output = re.search(r'[A-Z][A-Z]\d\.[A-Z][A-Z]\d\.\w+', str(value))
                if output:
                    print(value)
                    search_unit(x,y, ws)

