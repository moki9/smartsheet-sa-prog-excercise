import csv
import logging

import pandas as pd
import smartsheet

from pathlib import Path

from dotenv import load_dotenv

logging.getLogger('smartsheet.log').setLevel(logging.CRITICAL)

def create_or_get_sheet(client,sheet_spec):
    # return seet of a sheet if already exists
    response = client.Sheets.list_sheets(include_all=True)
    for sheet in response.data:
        if sheet.name == sheet_spec.name:
            sheet = client.Sheets.get_sheet(sheet.id)
            return sheet
    # otherwise, create a new one 
    new_sheet = client.Home.create_sheet(sheet_spec)
    return new_sheet.result  



if __name__ == "__main__":

    config_path = Path("config/devtoken")
    load_dotenv(dotenv_path=config_path)

    csv_path = Path("data/data.csv")
    rows = []
    with open(csv_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            rows.append(row)

    column_names = ["country", "state", "city", "total arr"]

    columns = []
    for i, col in enumerate(column_names):
        column = {
            "title": col,
            "primary": True if (i==0) else False,
            "type": "TEXT_NUMBER" 
        }

        columns.append(column)

    sheet_spec = smartsheet.models.Sheet({
            "name": "ARR per location",
            "columns": columns
        })

    client = smartsheet.Smartsheet()
    client.errors_as_exceptions(True)

    sheet = create_or_get_sheet(client, sheet_spec)

    columns = [column for column in sheet.columns]
            
    # Load data from csv file
    df = pd.read_csv("./data/data.csv")

    # Grouping data by country, state, and city and calculating total ARR for each group
    grouped = df.groupby(['country', 'state', 'city']).agg({'arr': 'sum'}).reset_index()
    grouped.set_index(['country', 'state', 'city'], inplace=True)


    for i, dframe in grouped.items():
    
        new_rows = []
        hierarchy = ["country", "state", "city"]
        prev_country = prev_state = prev_city = ""

        for location, arr in dframe.items():
            new_row = smartsheet.models.Row()
            new_row.to_top = True
            (country, state, city) = location
            data_dict = {
                "country": country, 
                "state": state,
                "city": city,
                "total arr": arr
            }
            for key, value in data_dict.items():
                column_id = next((col.id for col in columns if col.title == key), None)
                if key in hierarchy:
                    if key == "country":
                        tmp = value
                        if value == prev_country:
                            value = ""
                        prev_country = tmp
                    elif key == "state":
                        tmp = value
                        if value == prev_state:
                            value = ""
                        prev_state = tmp
                    elif key == "city":
                        tmp = value
                        if value == prev_city:
                            value = ""
                        prev_city = value
                    
                new_row.cells.append({
                    'column_id': column_id,
                    'value': value
                })
            new_rows.append(new_row)
    
    print(f"Added rows: {len(new_rows)} ")
    response = client.Sheets.add_rows(sheet.id, new_rows)
    print(f"New Sheet Data {response}")
