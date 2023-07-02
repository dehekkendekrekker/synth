#!/bin/python3
from pyexcel_ods3 import get_data, save_data
from pathlib import Path
import json
import csv



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

# Get all .csv files
pathlist = Path(".").glob('**/*.csv')
for path in pathlist:
    data = get_collated_data(path)
    tabname = str(path).split("/")[0]
    odsbom[tabname] = data


save_data("BOM.ods", odsbom)
     