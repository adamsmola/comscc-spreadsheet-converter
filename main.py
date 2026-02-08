from openpyxl import load_workbook
import json

def process_vehicles(sheet):
    CURRENT_YEAR = 2026

    COL_MAKE = 1
    COL_MODEL = 2
    COL_START_YEAR = 3
    COL_END_YEAR = 4
    COL_SHOWROOM_BASE_WEIGHT = 5
    COL_FACTORY_HP = 6
    COL_FACTORY_TQ = 7
    COL_SUSP_INDEX = 11 

    ROW_START = 13
    ROW_END = 501 

    entries = []
    index = ROW_START
    for row in sheet.iter_rows(min_row=ROW_START, max_row=ROW_END, values_only=True):

        data = {}
        data['id'] = index
        data['make'] = row[COL_MAKE]
        data['model'] = row[COL_MODEL]
        data['start_year'] = row[COL_START_YEAR]
        data['end_year'] = row[COL_END_YEAR]
        data['showroom_weight'] = row[COL_SHOWROOM_BASE_WEIGHT]
        data['factory_hp'] = row[COL_FACTORY_HP]
        data['factory_tq'] = row[COL_FACTORY_TQ]
        data['susp_index'] = row[COL_SUSP_INDEX]
        
        entries.append(data)
        index += 1

    return entries


def process_sheet(sheet, start_row, end_row):

    COL_POINTS = 1       # Col A
    COL_DESCRIPTION = 2  # Col B

    entries = []
    index = start_row
    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, values_only=True):
        if row[1] == "Dyno": continue; 
        data = {}
        data['id'] = index
        data['description'] = str(row[COL_DESCRIPTION]).encode("ascii", "ignore").decode()
        data['points'] = row[COL_POINTS]
        entries.append(data)
        index += 1

    return entries

def process_engine(sheet):
    START_ROW = 9 
    END_ROW = 58
    return process_sheet(sheet, START_ROW, END_ROW)    

def process_drivetrain(sheet):
    START_ROW = 9 
    END_ROW = 30
    return process_sheet(sheet, START_ROW, END_ROW)

def process_suspension(sheet):
    START_ROW = 9 
    END_ROW = 38    
    return process_sheet(sheet, START_ROW, END_ROW) 

def process_brakes(sheet):
    START_ROW = 9 
    END_ROW = 19 
    return process_sheet(sheet, START_ROW, END_ROW)

def process_exterior(sheet):
    START_ROW = 9 
    END_ROW = 29
    return process_sheet(sheet, START_ROW, END_ROW)

def process_tires(sheet):
    START_ROW = 9
    END_ROW = 77
    return process_sheet(sheet, START_ROW, END_ROW)


def main():
    
    wb = load_workbook('touring.xlsx', read_only=True, data_only=True)

    sheet = wb["Vehicles"]
    data = process_vehicles(sheet)
    with open("vehicles.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Engine"]
    data = process_engine(sheet)
    with open("engine.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Drivetrain"]
    data = process_drivetrain(sheet)
    with open("drivetrain.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Suspension"]
    data = process_suspension(sheet)
    with open("suspension.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Brakes"]
    data = process_brakes(sheet)
    with open("brakes.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Exterior"]
    data = process_exterior(sheet)
    with open("exterior.json", "w") as json_file:
        json.dump(data, json_file, indent=4) 

    sheet = wb["Tires"]
    data = process_tires(sheet)
    with open("tires.json", "w") as json_file:
        json.dump(data, json_file, indent=4)     

 




if __name__ == "__main__":
    main()
