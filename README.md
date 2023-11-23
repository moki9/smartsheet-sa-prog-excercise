# ARR By Location

This code snippet is responsible for creating a Smartsheet, loading data from a CSV file, grouping the data by country, state, and city, and adding the grouped data to the Smartsheet. It also indents the rows in the Smartsheet based on the location values and sorts the rows by the "Location" column.

Inputs

* CSV file path
* Column names for the Smartsheet
* Smartsheet API client

Flow

1. Load the Smartsheet API token from a configuration file.
2. Read the data from the CSV file and store it in a list of rows.
3. Define the column names and their properties for the Smartsheet.
4. Create a Smartsheet specification object with the sheet name and columns.
5. Create a Smartsheet API client and set it to handle errors as exceptions.
6. Create or retrieve the Smartsheet using the client and sheet specification.
7. Delete any existing data in the Smartsheet.
8. Get the columns of the Smartsheet.
9. Load the data from the CSV file into a pandas DataFrame.
10. Group the data by country, state, and city, and calculate the total ARR for each group.
11. Create a dictionary to store the unique locations (countries, states, cities).
12. Create a nested dictionary representing the tree structure of the grouped data.
13. Add the rows to the Smartsheet using the add_rows function.
14. Indent the rows in the Smartsheet based on the location values using the indent_rows function.
15. Sort the rows in the Smartsheet by the "Location" column using the sort_by_column function.

Outputs

* The Smartsheet with the grouped and indented data, sorted by the "Location" column.
