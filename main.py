import csv
import logging

from collections import defaultdict

import pandas as pd
import smartsheet

from pathlib import Path

from dotenv import load_dotenv

logging.getLogger('smartsheet.log').setLevel(logging.CRITICAL)


def get_column_id(columns, key):
    return next((col.id for col in columns if col.title == key), None)

def flatten(tree):
    """
    Flatten a nested dictionary representing a tree structure into a list of dictionaries.

    Args:
        tree (dict): A nested dictionary representing a tree structure where each key represents a country, each value is a dictionary representing states, and each value in the states dictionary is a list of cities.

    Returns:
        list: A flattened list of dictionaries representing the countries, states, and cities from the input tree. Each dictionary has two keys: "item" which contains the name of the country, state, or city, and "kind" which indicates whether it is a country, state, or city.
    """
    flattened = []
    is_parent = True
    for country, states in tree.items():
        flattened.append({"item": country, "kind": "country"})
        for state, cities in states.items():
            flattened.append({"item": state, "kind": "state"})
            for city in cities:
                flattened.append({"item": city, "kind": "city"})
    return flattened

def create_or_get_sheet(client, sheet_spec):
    """
    Checks if a sheet with a specific name already exists in a client's account.
    If it does, the function returns the existing sheet. If it doesn't, the function creates a new sheet using the provided sheet specification.

    Args:
        client (object): The client object representing the connection to the client's account.
        sheet_spec (object): The sheet specification object containing the details of the sheet to be created or retrieved.

    Returns:
        object: The sheet object representing the created or retrieved sheet.
    """

    response = client.Sheets.list_sheets(include_all=True)
    for sheet in response.data:
        if sheet.name == sheet_spec.name:
            sheet = client.Sheets.get_sheet(sheet.id)
            return sheet

    new_sheet = client.Home.create_sheet(sheet_spec)
    return new_sheet.result

def get_cell_value_by_row_and_column(row, columns, column):
    """
    Retrieves the value of a specific cell in a row based on the column title.

    Args:
        row (smartsheet.models.Row): The row object from which to retrieve the cell value.
        columns (list): A list of column objects that contain the column titles and IDs.
        column (str): The title of the column for which to retrieve the cell value.

    Returns:
        str: The display value of the cell in the specified column.
    """
    column_id = get_column_id(columns=columns, key=column)
    cell = row.get_column(column_id)
    return cell.display_value

def add_rows(client, sheet, tree):
    """
    Adds new rows to a Smartsheet sheet based on a tree structure.

    Args:
        client (smartsheet.Smartsheet): The Smartsheet client object used to interact with the Smartsheet API.
        sheet (smartsheet.models.Sheet): The Smartsheet sheet object representing the sheet where the new rows will be added.
        tree (dict): A nested dictionary representing a tree structure where each key represents a country, each value is a dictionary representing states, and each value in the states dictionary is a list of cities.

    Returns:
        None
    """
    flattened_tree = flatten(tree)
    new_rows = []
    for item in flattened_tree:
        new_row = smartsheet.models.Row()
        new_row.to_top = True 

        if type(item["item"]) is list:
            new_row.cells.append({
                    'column_id': get_column_id(columns=columns, key='Location'),
                    'value': item["item"][0]
            }) 
            new_row.cells.append({
                'column_id': get_column_id(columns=columns, key='ARR'),
                'value': item["item"][1]
            })
        else:
            new_row.cells.append({
                'column_id': get_column_id(columns=columns, key='Location'),
                'value': item["item"]
            }) 
            new_row.cells.append({
                'column_id': get_column_id(columns=columns, key='ARR'),
                'value': "" # blank
            }) 
        
        new_rows.append(new_row)
    
    response = client.Sheets.add_rows(sheet.id, new_rows)
    print(f"New Sheet Data {response}")


def indent_rows(client, sheet, columns, locations):
    """
    Indent rows in a Smartsheet based on the location values in the rows.

    Args:
        client (smartsheet.Smartsheet): An instance of the Smartsheet client.
        sheet (smartsheet.models.Sheet): The Smartsheet sheet object.
        columns (list): A list of column objects that contain the column titles and IDs.
        locations (dict): A dictionary containing the locations (countries, states, cities) to be indented.

    Code Analysis:
        - Get the sheet object using the Smartsheet API.
        - Initialize variables `country_id` and `state_id` to None.
        - Iterate over each row in the sheet.
        - Get the value of the "Location" column for the current row using the `get_cell_value_by_row_and_column` function.
        - If the location value is None, exit the loop.
        - Get the value of the "ARR" column for the current row using the `get_cell_value_by_row_and_column` function.
        - Create a new row object and set its properties.
        - Add cells to the new row object with the column ID and value.
        - Determine the parent row ID based on the location value.
        - Update the row in the Smartsheet using the Smartsheet API.
    """
    sheet = client.Sheets.get_sheet(sheet.id)
    country_id = None
    state_id = None

    for row in sheet.rows:
        print(f"{row.row_number} ->> {row.id}")
        loc = get_cell_value_by_row_and_column(row, columns=columns, column="Location")
        if loc == None:
            break

        arr = get_cell_value_by_row_and_column(row, columns=columns, column="ARR")

        new_row = smartsheet.models.Row()
        new_row.to_top = True
        new_row.id = row.id

        new_row.cells.append({
            'column_id': get_column_id(columns=columns, key='Location'),
            'value': loc
        }) 
        new_row.cells.append({
            'column_id': get_column_id(columns=columns, key='ARR'),
            'value': "" if arr == None else arr
        })

        if loc in locations["countries"]:
            country_id = row.id
            new_row.parent_id = None

        if loc in locations["states"]:
            state_id = row.id
            new_row.parent_id = country_id

        if loc in locations["cities"]:
            new_row.parent_id = state_id

        response = client.Sheets.update_rows(sheet.id, [new_row])
        print(f"Response: {response}")

def delete_existing_data(client, sheet, chunk_interval=300):
    """
    Deletes all existing rows in a Smartsheet.

    Args:
        client (smartsheet.Smartsheet): An instance of the Smartsheet client.
        sheet (smartsheet.models.sheet.Sheet): The Smartsheet sheet object representing the sheet from which to delete the rows.
        chunk_interval (int, optional): The number of rows to delete in each API call. Default value is 300.

    Returns:
        None
    """
    rows_to_delete = [row.id for row in sheet.rows]
    for x in range(0, len(rows_to_delete), chunk_interval):
        client.Sheets.delete_rows(sheet.id, rows_to_delete[x:x + chunk_interval])

def sort_by_column(client, sheet, column_id, order='DECENDING'):
    """
    Sorts the rows in a Smartsheet based on a specified column and order.

    Args:
        client (Smartsheet client): An instance of the Smartsheet client.
        sheet (Smartsheet sheet object): The Smartsheet sheet object representing the sheet to be sorted.
        column_id (int): The ID of the column to sort by.
        order (str, optional): The sort order, either 'ASCENDING' or 'DECENDING'. Default is 'DECENDING'.

    Returns:
        None
    """
    print(f"Sorting by {column_id}")
    sort_specifier = smartsheet.models.SortSpecifier({
        'sort_criteria': [smartsheet.models.SortCriterion({
            'column_id': column_id,
            'direction': order
        })]
    })
    sheet = client.Sheets.sort_sheet(sheet.id, sort_specifier)
    print(f"Sorted: {sheet}")

if __name__ == "__main__":
    config_path = Path("config/devtoken")
    load_dotenv(dotenv_path=config_path)

    csv_path = Path("data/data.csv")
    rows = []
    with open(csv_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            rows.append(row)

    column_names = ["Location", "ARR"]

    columns = []
    for i, col in enumerate(column_names):
        column = {
            "title": col,
            "primary": True if (i==0) else False,
            "type": "TEXT_NUMBER" 
        }

        columns.append(column)

    sheet_spec = smartsheet.models.Sheet({
            "name": "ARR per Location",
            "columns": columns
        })

    client = smartsheet.Smartsheet()
    client.errors_as_exceptions(True)

    sheet = create_or_get_sheet(client, sheet_spec)
    delete_existing_data(client, sheet)

    columns = [column for column in sheet.columns]
            
    # Load data from csv file
    df = pd.read_csv("./data/data.csv")

    # Grouping data by country, state, and city and calculating total ARR for each group
    grouped = df.groupby(['country', 'state', 'city']).agg({'arr': 'sum'}).reset_index()
    grouped.set_index(['country', 'state', 'city'], inplace=True)

    grouped_loc = {}
    grouped_loc["countries"] = df['country'].unique().tolist()
    grouped_loc["states"] = df['state'].unique().tolist()
    grouped_loc["cities"] = df['city'].unique().tolist()

    tree = defaultdict(lambda: defaultdict(list))

    for _, dframe in grouped.items():
        countries = states = cities = []
        for location, arr in dframe.items():
            (country, state, city) = location
            tree[country][state].append([city, arr])

    add_rows(client=client, sheet=sheet, tree=tree)    
    indent_rows(client=client, sheet=sheet, columns=columns, locations=grouped_loc)
    sort_by_column(client=client, sheet=sheet, column_id=get_column_id(columns=columns, key='Location'))


