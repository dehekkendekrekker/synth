#!/bin/python3
import pandas as pd
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

   

sheets={}
            
with pd.ExcelFile("BOM.ods") as xls:

    # Get all .csv files
    pathlist = Path(".").glob('**/*.csv')
    for path in pathlist:
        data = get_collated_data(path)
        tabname = str(path).split("/")[0]
        sheet = pd.read_excel(xls, tabname, engine="odf")
        sheet.update(data)
        sheets[tabname] = sheet

with pd.ExcelWriter("BOM2.ods") as xls:
    for name, sheet in sheets.items():
        sheet.to_excel(xls, sheet_name=name)



        
    

    
     

