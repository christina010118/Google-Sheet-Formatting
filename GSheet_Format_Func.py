from __future__ import print_function
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import gspread
import datetime
from dateutil.relativedelta import relativedelta

#This file inlcudes all functions needed for Google sheets format


SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Create credentials and service
creds = Credentials.from_authorized_user_file("token.json", SCOPES)
service = build('sheets', 'v4', credentials=creds)

#Helper Function
def get_sheet_id_by_name(spreadsheet_id, sheet_name):
    """
    This function retrieves the sheet_id by sheet_name
    """

    spreadsheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet_metadata.get('sheets', [])

    for sheet in sheets:
        properties = sheet.get('properties', {})
        if properties.get('title', '') == sheet_name:
            return properties.get('sheetId', 0)

    return 0  #Return 0 if sheet name is not found

def calculate_quarter(month):
    """
    This function build the logic for asigning quarters each month in a year
    """
    if 1 <= month <= 3:
        return 'Q1'
    elif 4 <= month <= 6:
        return 'Q2'
    elif 7 <= month <= 9:
        return 'Q3'
    elif 10 <= month <= 12:
        return 'Q4'

def generate_quarters(start_year, end_year, start_month):
    """
    This function generates a list of quarters given start year and end year
    """
    list_of_quarters = []
    next_quarter_date = datetime.datetime(year=start_year, month=start_month, day=1)
    end_quarter_date = datetime.datetime(year=end_year, month=12, day=31)

    list_of_quarters.append(f'{calculate_quarter(start_month)} {start_year}')

    while next_quarter_date < end_quarter_date:
        next_quarter_date += relativedelta(months=+3)
        if next_quarter_date <= end_quarter_date:
            list_of_quarters.append(f'{calculate_quarter(next_quarter_date.month)} {next_quarter_date.year}')

    return list_of_quarters

def generate_ranges_for_cell_calculation(spreadsheet_id,sheet_name, start_col):

    # Find the last row in the sheet with values
    values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A:A').execute()
    last_row = len(values.get('values', []))  # Add 1 to go to the next row

    sheet_properties = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    max_column_count = sheet_properties['sheets'][0]['properties']['gridProperties']['columnCount']

    cell_ranges = []
    cell_letters = []
    for i in range(start_col, max_column_count):
        col_letter= chr(ord('A') + i - 1)
        # Create the cell range for the current column (assuming rows 6 to 19)
        col_range = f'{col_letter}6:{col_letter}{last_row}'
    
        # Append the cell range and letter to their respective lists
        cell_ranges.append(col_range)
        cell_letters.append(col_letter)
    return cell_ranges, cell_letters

#Functions
def delete_sheet_by_name(spreadsheet_id, sheet_name):

    """
    This function deleted spreadsheets in the file by sheet_name
    """
    # Get the list of sheets in the document
    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = sheet_metadata['sheets']

    sheet_id = None

    # Find the sheet by name
    for sheet_info in sheets:
        if sheet_info['properties']['title'] == sheet_name:
            sheet_id = sheet_info['properties']['sheetId']
            break

    if sheet_id is not None:
        # Create a request to delete the sheet by sheet_id
        request = {
            'requests': [
                {
                    'deleteSheet': {
                        'sheetId': sheet_id
                    }
                }
            ]
        }

        try:
            # Execute the request to delete the sheet
            response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request).execute()
            print(f"Sheet '{sheet_name}' deleted successfully.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
    else:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")

def add_rows_to_sheet(spreadsheet_id, sheet_name, num_rows_to_insert):
    """
    This function adds rows to the sheet
    """

    # Get the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    if sheet_id == 0:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    # Create the request body for inserting rows
    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": 0,  # Start at the top
                        "endIndex": num_rows_to_insert  # Number of rows to insert
                    },
                    "inheritFromBefore": False
                }
            }
        ]
    }

    try:
        # Execute the request to insert rows at the top
        response = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request_body).execute()
        print(f"{num_rows_to_insert} rows added to the sheet '{sheet_name}'.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def replace_value(spreadsheet_id, sheet_name, cell_range, new_value):
    """
    This function replacing the specified cell range with new values 
    """
    # Update the value in the specified cell range
    request = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f'{sheet_name}!{cell_range}',  # Specify the cell range
        valueInputOption='RAW',
        body={'values': [[new_value]]}
    )
    response = request.execute()

    # Check the response for errors, if needed
    if 'error' in response:
        print(f"Error updating value: {response['error']['message']}")
    else:
        print(f"Value updated successfully in {sheet_name}!{cell_range}")

def merge_cells_and_fill_labels(spreadsheet_id, sheet_name):
    """
    This function merges every three cells from column 2 to the maximum column in the sheet
    and fills them with quarter labels.
    """
    # Get the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    if sheet_id == 0:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    # Calculate the maximum column count based on the sheet properties
    sheet_properties = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    max_column_count = sheet_properties['sheets'][0]['properties']['gridProperties']['columnCount']

    merge_cells_requests = []
    cell_ranges = []

    cell_count=0
    # Loop through every three columns starting from column 2
    for start_col_index in range(1, max_column_count, 3):
        end_col_index = min(start_col_index + 2, max_column_count - 1)
        cell_count+=1

        # Convert column indices to column letters (e.g., 0 -> 'A', 1 -> 'B', etc.)
        start_column = chr(ord('A') + start_col_index)
        end_column = chr(ord('A') + end_col_index)

        cell_range = f"{start_column}4:{end_column}4"  # Updated to row 4
        cell_ranges.append(cell_range)

        merge_cells_request = {
            'mergeCells': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 3,  # Starting from row 4 (0-based index)
                    'endRowIndex': 4,   # Ending at row 5 (0-based index)
                    'startColumnIndex': start_col_index,
                    'endColumnIndex': end_col_index + 1,
                },
                'mergeType': 'MERGE_ALL'
            }
        }
        merge_cells_requests.append(merge_cells_request)

    
    # Check if the sheet name indicates a specific starting point
    if sheet_name == "Followers (OO)" or sheet_name == "Followers (Total Page - EG Entities)":
        # Starting from Q4 2022
        start_year = 2022
        start_month=10
    else:
        # Starting from Q3 2022 (default)
        start_year = 2022
        start_month=7
    end_year= start_year + cell_count//4
    quarter_labels = generate_quarters(start_year, end_year, start_month)

    # Perform the batch update to merge the cells
    batch_update_request = {
        'requests': merge_cells_requests
    }
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=batch_update_request
    )
    response = request.execute()

    if 'error' in response:
        print(f"Error merging cells: {response['error']['message']}")
    else:
        print("Cells merged and value filled successfully")

    # Call the existing fill_values function to fill the labels
    fill_values(spreadsheet_id, sheet_name, cell_ranges, quarter_labels)

def fill_values(spreadsheet_id, sheet_name, cell_ranges, values_to_fill):
    """
    This function fills the above merged cells with Quarter labels
    """

    # Get the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    if sheet_id == 0:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
        return
    repeat_cell_requests = []

    for i, cell_range in enumerate(cell_ranges):
        start_cell, end_cell = cell_range.split(':')
        repeat_cell_request = {
            'repeatCell': {
                'range': {
                    'sheetId': sheet_id,  
                    'startRowIndex': int(start_cell[1:]) - 1,  # Convert to 0-based index
                    'endRowIndex': int(end_cell[1:]),  # Convert to 0-based index
                    'startColumnIndex': ord(start_cell[0]) - ord('A'),  # Convert to 0-based index
                    'endColumnIndex': ord(end_cell[0]) - ord('A') + 1,  # Convert to 0-based index
                },
                'cell': {
                    'userEnteredValue': {
                        'stringValue': values_to_fill[i] if i < len(values_to_fill) else ''
                    },
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.0,  # Background color black
                            'green': 0.0,
                            'blue': 0.0
                        },
                        'textFormat': {
                            "bold": True,
                            'fontSize': 12,
                            'foregroundColor': {
                                'red': 1.0,  # Text color white
                                'green': 1.0,
                                'blue': 1.0
                            }
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                'fields': 'userEnteredValue,userEnteredFormat.backgroundColor,userEnteredFormat.textFormat, userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment'
            }
        }
        repeat_cell_requests.append(repeat_cell_request)

    batch_update_request = {
        'requests': repeat_cell_requests
    }

    request = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=batch_update_request
    )
    response = request.execute()

    if 'error' in response:
        print(f"Error filling values: {response['error']['message']}")
    else:
        print("Values filled successfully")

def format_sheet(spreadsheet_id, sheet_name):
    """
    This function formats the table by bolding both row and col headers, and adjusting the color of the headers
    """
    # Get the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    if sheet_id == 0:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

    # Find the specified sheet by name and get its properties
    for sheet in sheet_metadata['sheets']:
        if sheet['properties']['title'] == sheet_name:
            max_column_count = sheet['properties']['gridProperties']['columnCount']
            break  # Stop searching once you find the sheet

    # Define the requests to format the cells
    requests = [
        #bolding row 3-5
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,  
                    "startRowIndex": 2, 
                    "endRowIndex": 5,  
                    "startColumnIndex": 0,
                    "endColumnIndex": 16
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat.bold"
            }
        },
        
        #bolding column 1 all titles
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,  
                    "startRowIndex": 0, 
                    "endRowIndex": 20,  
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat.textFormat.bold"
            }
        },
        #change header col to blue 
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,  
                    "startRowIndex": 2,  # Rows 3, 4 and 5
                    "endRowIndex": 5,
                    "startColumnIndex": 0, #col 1
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.58,
                            "green": 0.71,
                            "blue": 0.92  # RGB color for the fill
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        },
        #changing header row 3 background to blue
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id, 
                    "startRowIndex": 2,  
                    "endRowIndex": 3,
                    "startColumnIndex": 0,
                    "endColumnIndex": max_column_count  
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.58,
                            "green": 0.71,
                            "blue": 0.92  # RGB color for the fill
                        }
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        },
        #changing month row
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,  
                    "startRowIndex": 2,  #Row 3, 0 index based
                    "endRowIndex": 3,
                    "startColumnIndex": 0,
                    "endColumnIndex": max_column_count  #Adjust to the number of columns in your sheet
                },
                "cell": {
                    "userEnteredFormat": {
                        "textFormat": {
                            "bold": True,
                            "fontSize": 10
                        },
                        "horizontalAlignment": "CENTER",
                        "verticalAlignment": "MIDDLE"
                    }
                },
                "fields": "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
            }
        }
    ]

    batch_update_request = {
        'requests': requests
    }


    request = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=batch_update_request
    )
    response = request.execute()

    if 'error' in response:
        print(f"Error filling values: {response['error']['message']}")
    else:
        print("Values formatted successfully")

def calculate_sums(spreadsheet_id, sheet_name, cell_ranges, cell_letters):
    """
    This function calculates the sum for all columns in the cell ranges and puts the result
    in the next row after the last row in the sheet with values.
    """

    # Find the last row in the sheet with values
    values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A:A').execute()
    last_row = len(values.get('values', [])) + 1  # Add 1 to go to the next row

    for cell_range, cell_letter in zip(cell_ranges, cell_letters):
        # Define the range for the sum calculation
        range_to_sum = f"'{sheet_name}'!{cell_range}"

        # Define the formula to calculate the sum
        formula = f'=SUM({range_to_sum})'

        # Create a request to update a single cell with the calculated sum
        request = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!{cell_letter}{last_row}', #where the calculated result will be located
            valueInputOption='USER_ENTERED',
            body={'values': [[formula]]}
        )

        try:
            # Execute the request to calculate and update the sum
            response = request.execute()
            print(f"Sum calculated and updated successfully")
        except Exception as e:
            print(f"Error calculating sum: {str(e)}")

def remove_value(spreadsheet_id, sheet_name, cell):

    """
    This function removes value in a cell
    """

    try:
        # Clear the value in the specified cell
        request = service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=f'{sheet_name}!{cell}'
        )
        response = request.execute()
        print(f"Value in {cell} cleared successfully")
    except Exception as e:
        print(f"Error clearing value in {cell}: {str(e)}")

def change_font_size(spreadsheet_id, sheet_name, cell_range, font_size):
    """
    This function changes font size in a specified cell range.
    """

    # Get the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    # Define the range for which you want to change the font size
    range_to_format = f"'{sheet_name}'!{cell_range}"

    # Create a request to update the font size of the specified cell
    request = {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": int(cell_range[1:]) - 1,  # Convert to 0-based index
                "endRowIndex": int(cell_range[1:]),        # Convert to 0-based index
                "startColumnIndex": ord(cell_range[0]) - ord('A'),  # Convert to 0-based index
                "endColumnIndex": ord(cell_range[0]) - ord('A') + 1  # Convert to 0-based index
            },
            'cell': {
                'userEnteredFormat': {
                    'textFormat': {
                        'fontSize': font_size
                    }
                }
            },
            'fields': 'userEnteredFormat.textFormat.fontSize'  # Specify the field to update
        }
    }

    batch_update_request = {
        'requests': [request]  # You should put the request in a list
    }

    request = service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=batch_update_request
    )
    response = request.execute()

    if 'error' in response:
        print(f"Error font size: {response['error']['message']}")
    else:
        print("Font size updated successfully")

def add_top_border(spreadsheet_id, sheet_name):
    """
    This function adds a top border to a row
    """
    # Find max column
    sheet_properties = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    max_column_count = sheet_properties['sheets'][0]['properties']['gridProperties']['columnCount']

    # Find the last row in the sheet with values
    values = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f'{sheet_name}!A:A').execute()
    last_row = len(values.get('values', []))

    # Convert the end column index to a letter (e.g., 16 -> P)
    end_column_letter = chr(ord('A') + max_column_count - 1)

    range_to_border = f"'{sheet_name}'!A{last_row}:{end_column_letter}{last_row}"

    # Define the border style
    border_style = {
        "style": "SOLID",
        "width": 1,  # 1 pixel width for the border
        "color": {
            "red": 0.0,   # Border color (red component)
            "green": 0.0,  # Border color (green component)
            "blue": 0.0    # Border color (blue component)
        }
    }

    # Create a request to add the top border to the specified range
    request = {
        "updateBorders": {
            "range": {
                "sheetId": get_sheet_id_by_name(spreadsheet_id, sheet_name),
                "startRowIndex": last_row ,  #last row with values
                "endRowIndex": last_row+1,    #Last row with values +1
                "startColumnIndex": 0,  # Column A
                "endColumnIndex": max_column_count - 1  # Adjusted to the last column
            },
            "top": border_style
        }
    }

    try:
        # Execute the request to add the top border
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={"requests": [request]}
        ).execute()
        print(f"Top border added successfully to {range_to_border}")
    except Exception as e:
        print(f"Error adding top border: {str(e)}")

def auto_resize_columns(spreadsheet_id, sheet_name):

    """
    This function auto resize the entire spreadsheet cells
    """

    try:
        # Get the sheet ID based on the sheet name
        sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

        if sheet_id == 0:
            print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
            return

        # Determine the total number of columns in the sheet
        sheet_properties = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        num_columns = sheet_properties['sheets'][0]['properties']['gridProperties']['columnCount']

        # Define the batch size for resizing columns
        batch_size = 10  # You can adjust this as needed

        # Resize columns in batches
        for start_index in range(0, num_columns, batch_size):
            end_index = min(start_index + batch_size, num_columns)

            # Create a request to auto-resize a batch of columns
            request = {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": start_index,
                        "endIndex": end_index
                    }
                }
            }

            # Send the request to auto-resize columns
            response = service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={'requests': [request]}
            ).execute()

        print('Columns auto-resized to fit content.')

    except Exception as e:
        print(f'Error: {str(e)}')

def change_cell_background_color_to_white(spreadsheet_id, sheet_name, cell):

    """
    This function changes the background color to white
    """
    # Create credentials and service
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    service = build('sheets', 'v4', credentials=creds)

    # Define the sheet ID based on the sheet name
    sheet_id = get_sheet_id_by_name(spreadsheet_id, sheet_name)

    if sheet_id == 0:
        print(f"Sheet '{sheet_name}' not found in the spreadsheet.")
        return

    # Define the request to update cell background color to white
    request = {
        'updateCells': {
            'rows': [
                {
                    'values': [
                        {
                            'userEnteredFormat': {
                                'backgroundColor': {
                                    'red': 1.0,  # White background color
                                    'green': 1.0,
                                    'blue': 1.0
                                }
                            }
                        }
                    ]
                }
            ],
            'fields': 'userEnteredFormat.backgroundColor',
            'start': {
                'sheetId': sheet_id,
                'rowIndex': int(cell[1:]) - 1,  # Convert to 0-based index
                'columnIndex': ord(cell[0]) - ord('A')  # Convert to 0-based index
            }
        }
    }

    try:
        # Execute the request to change the cell background color
        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [request]}
        ).execute()
        print(f"Cell background color changed to white in {cell}")
    except Exception as e:
        print(f"Error changing cell background color: {str(e)}")



def apply_common_operations_to_sheet(copied_spreadsheet_id, sheet_name):
    """
    This function applies all required formats to the sheet
    """

    # Change cell background to white
    change_cell_background_color_to_white(copied_spreadsheet_id, sheet_name, "A1")

    # Add rows to the sheet
    add_rows_to_sheet(copied_spreadsheet_id, sheet_name, 4)

    if sheet_name == "Followers (Total Page - EG Entities)":
        replace_value(copied_spreadsheet_id, 'Followers (Total Page - EG Entities)', "A3", "EG Related Entities")
    else:
        replace_value(copied_spreadsheet_id, sheet_name, "A3", "Evil Geniuses (O&O)")

    # Remove value in A5
    remove_value(copied_spreadsheet_id, sheet_name, 'A5')

    #Merge cells and fill in Quarters
    merge_cells_and_fill_labels(copied_spreadsheet_id, sheet_name)

    #Only Followers sheet need sums calculation
    if sheet_name == "Followers (OO)" or sheet_name == "Followers (Total Page - EG Entities)":
                cell_ranges= generate_ranges_for_cell_calculation(copied_spreadsheet_id, sheet_name, 2)[0]
                cell_letters= generate_ranges_for_cell_calculation(copied_spreadsheet_id, sheet_name, 2)[1]
                calculate_sums(copied_spreadsheet_id, sheet_name, cell_ranges, cell_letters)
                add_top_border(copied_spreadsheet_id, sheet_name)

    # Format the sheet
    format_sheet(copied_spreadsheet_id, sheet_name)

    # Auto resize columns
    auto_resize_columns(copied_spreadsheet_id, sheet_name)
