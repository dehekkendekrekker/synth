#!/bin/python3
from pyexcel_ods3 import get_data, save_data
from pathlib import Path
import json
import csv
from loguru import logger


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
        print("%s--%s" % (value, library))
        idx_first = (value, library)
        idx_second = (value, None)
        if idx_first in self.idx:
            return self.items[self.idx[idx_first]]
        elif idx_second in self.idx:
            return self.items[self.idx[idx_second]]
        else:
            logger.error("Error: No component entry for V:%s L:%s" % (value, library))
            quit()

    def __str__(self) -> str:
        result = ""
        for key, i in self.items.items():
            result += "%s\n" % str(i)

        return result


MAPPING = "Mapping"

components = Components()
items = []
quantities = {}



def update_quantities():
    logger.info("Updating quantities")
    global items
    item : Item
    for item in items:
        #print(item)
        component = components.get(item.value, item.library)
        if component not in quantities:
            quantities[component] = 0
        
        quantities[component] += item.qty
        
    
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

def load_items(items_sheet):
    
    global items
    for row in items_sheet[1:]:
        #print(row)
        if not row:
            break

        items.append(Item(qty=row[1], value=row[3], library=row[4]))

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

    


            

odsbom = get_data("BOM.ods")

load_components(odsbom[MAPPING])

logger.info(components)
print(components.idx)



# Get all .csv files
pathlist = Path(".").glob('**/*.csv')
for path in pathlist:
    data = get_collated_data(path)
    tabname = str(path).split("/")[0]
    odsbom[tabname] = data

    print("Loading items from sheet [{}]", tabname)
    load_items(odsbom[tabname])



update_quantities()





save_data("BOM.ods", odsbom)
     