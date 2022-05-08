#!/usr/bin/env python3

from csv import DictReader, DictWriter, writer
from os import path, system
from pprint import pprint

import openpyxl as xl
from openpyxl.styles import PatternFill
from tqdm import tqdm


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
    print('-'*80)
    print(f'Serial numbers in Inventory list: {check}')
    print(f'Serial numbers FOUND against PDQ Inventry list: {len(serialsfound)}')
    print(f'Serial numbers NOT FOUND against PDQ Inventry list: {check - len(serialsfound)}')
    print('-'*80)
    print()
    print('Writing found and not found serial numbers to CSV file...')

    with open("serialsnotfound.csv", "w") as f:
        for item in serialsnotfound:
            f.write(item)
    with open("serialsfound.csv", "w") as f:
        for item in serialsnotfound:
            f.write(item)
    print('_'*80)
    return serialsfound


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
    inventory = input('Inventory file name: ')
    create_copy_of_inventory(f'{inventory}.xlsx')
    with open(f"{inventory}.csv", "r") as f:
        lines = f.readlines()
    return lines


def get_online_systems():
    online_systems = input('Online Systems file name: ')
    with open(f'{online_systems}.csv', "r") as f:
        lines = f.readlines()
        return lines
        
def colorcode(found):
    file = create_copy_of_inventory()
    print('_'*80)
    wb = xl.load_workbook(f'{file}')
    ws = wb["Finished_Inventory"]
    fill_green = PatternFill(patternType="solid", fgColor="00FF00")
    fill_red = PatternFill(patternType="solid", fgColor="FF0000")
    pbar =  tqdm(range(len(found)),desc='Color coding in progress: ')
    for item in found:
        # print(f'Looking for {item}')
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=4, max_col=4):
            for cell in row:
                # print(f'Checking rows and cols for item against {cell.value}')
                if cell.value == item:
                    # print(f'Found {item} : cell {cell.value}')
                    cell.fill = fill_green
                    # print('Attempting to colorcode cell Green')
                    wb.save(file)
    pbar.close()    
    # sorted_inventory(file)
    print('Finished Inventory')


# def sorted_inventory(file):
#     pass
#     wb = xl.load_workbook(file)
#     ws = wb["Finished_Inventory"]
#     ws.sort_values(by="Serial Number", cellColor)
#     wb.save(file)
#     print('Sorted Inventory')
#     print('_'*80)



def create_copy_of_inventory(file):
    wb = xl.load_workbook(file)
    ws = wb["May_inventory"]
    file_new = f'{file}_new.xlsx'
    wb.save(file_new)
    
    


if __name__ == "__main__":
    serialsfound = doinventory()
    colorcode(serialsfound)
    print('Colorcoded Inventory\nPlease check items not found in local file')
