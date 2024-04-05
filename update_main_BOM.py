#!/bin/python3
from pyexcel_ods3 import get_data, save_data
from pathlib import Path
import json
import csv
from loguru import logger
from argparse import ArgumentParser


class Component:
    name = None
    library = None
    value = None
    mouser_partnr = None

    def __init__(self, name, library, value, mouser_partnr) -> None:
        '''
        Typecasts and trims input data
        '''
        self.name = str(name).strip()
        self.library = str(library).strip().upper() if str(library).strip() else None
        self.value = str(value).strip().upper()
        self.mouser_partnr = str(mouser_partnr).strip()

    def __str__(self) -> str:
        return "name: [%s] library: [%s] value: [%s] partnr: [%s]" % (self.name, self.library, self.value, self.mouser_partnr)



class Item:
    qty = None
    library = None
    value = None

    def __init__(self, qty, library, value) -> None:
        self.qty = int(qty)
        self.library = str(library).strip().upper() if str(library) else None
        self.value = str(value).strip().upper()

        if self.value[0] == ".":
            self.value = "0" + self.value

    def __str__(self) -> str:
        return "%s:%s->%s" % (self.library, self.value, self.qty)


class Components:
    items = {}
    idx = {}

    def append(self, component : Component):
        component_hash = hash(component)
        self.items[component_hash] = component

        if (component.value, component.library) in self.idx:
            logger.error("Duplicate component V:[%s] L:[%s]" % (component.value, component.library))
            quit()

        self.idx[(component.value, component.library)] = component_hash

           
    def get(self, value, library):
        idx_first = (value, library)
        idx_second = (value, None)
        if idx_first in self.idx:
            return self.items[self.idx[idx_first]]
        elif idx_second in self.idx:
            return self.items[self.idx[idx_second]]
        else:
            return None

    def __str__(self) -> str:
        result = ""
        for key, i in self.items.items():
            result += "%s\n" % str(i)

        return result


MAPPING = "Mapping"
QTY_BOARDS = "QtyBoards"
BOM = "BOM"

components = Components()
items = []
quantities = {}
qty_boards = {}


def update_quantities():
    logger.info("Updating quantities")
    errors = []
    global items
    global quantities
    item : Item

    for item in items:
        component = components.get(item.value, item.library)
        if component is None:
            errors.append(item)
        if component not in quantities:
            quantities[component] = 0
        
        quantities[component] += int(item.qty)

    if errors:
        for item in errors:
            logger.error("Error: No component entry in mapping tab for V:%s L:%s" % (item.value, item.library))
        quit()



def update_BOM_sheet(sheets):
    data = [[x.mouser_partnr, qty] for x, qty in quantities.items()]
    sheets[BOM] = data





def display_quantities():
    for key,qty in quantities.items():
        print("%s %s" % (key.mouser_partnr, qty))



        
    
def load_components(component_sheet):
    global components
    for row in component_sheet[1:]:
        if not row:
            break # We're done

        name, library, value, mouser_partnr = row[0:4]

        # Filter out invalid components
        if not library and not value:
            logger.warning("Throwing out: %s" % row)
            continue

        components.append(Component(name, library, value,mouser_partnr))

def load_items(sheets, sheet_name):
    if sheet_name not in qty_boards:
        logger.warning("No entry of {} in {} sheet", sheet_name, QTY_BOARDS)
        return

    if qty_boards[sheet_name] == 0:
        logger.warning("Quantity of sheet {} is set to zero. Items in this sheet will not be used for the BOM", sheet_name)
        return

    board_qty = qty_boards[sheet_name]

    global items
    for row in sheets[sheet_name][1:]:
        if not row:
            break

        qty = board_qty * int(row[1])
        items.append(Item(qty=qty, value=row[3], library=row[4]))


def load_qty_boards(qty_boards_sheet):
    global qty_boards
    for row in qty_boards_sheet:
        if not row:
            break
        qty_boards[row[0]] = row[1]

    



def get_collated_data(file):
    ST_SEARCHING=0
    ST_COLLATED_DETECTED=1
    ST_PARSING=2
    ST_DONE=3

    state = ST_SEARCHING
    result = []
    
    with open(file, newline='') as csvfile:
        bomreader = csv.reader(csvfile, dialect="unix")
        for row in bomreader:
            if state == ST_SEARCHING:
                if not row:
                    continue
                if row[0] == "Collated Components:":
                    state = ST_COLLATED_DETECTED
                    continue

            if state == ST_COLLATED_DETECTED:
                if not row:
                    continue
                if row[0] == "Item":
                    state = ST_PARSING

            if state == ST_PARSING:
                if row:
                    result.append(row[0:5])
                else:
                    state = ST_DONE
        
            if state == ST_DONE:
                break

    return result

    

# === Start ===

parser = ArgumentParser()
parser.add_argument("--filename", required=True)
args = parser.parse_args()
            

sheets = get_data(args.filename)

load_components(sheets[MAPPING])
load_qty_boards(sheets[QTY_BOARDS])

# Get all .csv files
pathlist = Path(".").glob('**/*.csv')
for path in pathlist:
    data = get_collated_data(path)
    sheet_name = str(path).split("/")[0]
    sheets[sheet_name] = data

    logger.info("Loading items from sheet [{}]", sheet_name)
    load_items(sheets,sheet_name)



update_quantities()
display_quantities()
update_BOM_sheet(sheets)




save_data(args.filename, sheets)
     