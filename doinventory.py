#!/usr/bin/env python3
from pprint import pprint
from csv import DictReader, DictWriter, writer
import openpyxl as xl
from openpyxl.styles import PatternFill



def doinventory():
    
    inventory = get_inventory()
    check = 0
    serialsfound = []
    serialsnotfound = []

    for line in inventory:
        line = line.split(",")
        if line[3]:
            check = check + 1
            checkingNone = (checkonlinesystems(line[3]))
            if checkingNone == None:
                serialsnotfound.append(line[3])
            else:
                serialsfound.append(checkingNone)

    print("FOUND:")
    print('+'*80)
    print(serialsfound)
    print('+'*80)
    print()
    print("NOT FOUND:")
    print('-'*80)
    print(serialsnotfound)
    print('-'*80)
    print()
    print('_'*80)
    print(f'Serial numbers checked in Inventory list: {check}')
    print(f'Online PDQ Inventory serial numbers -VS- NHLEM Inventry list: {len(serialsfound)}')
    print(f'Missing serial numbers: {check - len(serialsfound)}')
    print('_'*80)

    with open("missingSerials.csv", "w") as f:
        for item in serialsnotfound:
            f.write(item)

    with open("serialsfound.csv", "w") as f:
        for item in serialsnotfound:
            f.write(item)
    colorcode(serialsfound)


def checkonlinesystems(check):
    
    onlinesys =  get_online_systems()
    serialsfound = []
    serialsnotfound = []
    found = 0
    
    for serial in onlinesys:
        serial = serial.split(",")
        
        if serial[3] != check:
            serialsnotfound.append(check)
        elif serial[3] == check:
            found = found + 1
            return check


def get_inventory():
    with open("May_inventory.csv", "r") as f:
        lines = f.readlines()
        return lines


def get_online_systems():
    with open("OnlineSystems.csv", "r") as f:
        lines = f.readlines()
        return lines
        
def colorcode(found):
    wb = xl.load_workbook("May_inventory.xlsx")
    ws = wb["May_inventory"]
    fill_green = PatternFill(patternType="solid", fgColor="00FF00")
    fill_red = PatternFill(patternType="solid", fgColor="FF0000")

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=4, max_col=4):
        for cell in row:
            if cell.value in (x for x in found):
                cell.fill = fill_green

            else:
                cell.fill = fill_red
            
    wb.save("May_inventory2.xlsx")

    print('From colorcode:')
    print(found)

if __name__ == "__main__":
    doinventory()
    colorcode(found=['SERIALNUMBER'])
