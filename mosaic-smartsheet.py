'''Uses Smartsheet's API to pull a sheet. Then extracts the zone numbers, then
   uses the exported csv --or modifies in-place??-- the mosaic designer layout.

   Create .csv directly from smartsheet, and then import?

   this could be a web app.

   TODO: use OAuth2 for access/authentication

   take sheet_title as argument
'''
import socket
import argparse
import csv
import logging

# import urllib
import smartsheet
import fixture_types

REMOTE_SERVER = '4wall.com'
testing_sheet_id = 8769876857776004
testing_column_id = 1182995975038852

logging.basicConfig(filename = 'mosaic-smartsheet.log', level=logging.INFO)

parser = argparse.ArgumentParser(
                    prog = 'mosaic-smartsheet',
                    description = 'generate mosaic layout from smartsheet',
                    epilog = 'some text here I guess')
parser.add_argument('sheet_name', 
                    help = '''The name of a Smartsheet sheet. Enclose with quotes if
                            the sheet name contains spaces''')

def check_for_internet(hostname):
    '''
    see if we have a working internet connection
    copied from https://stackoverflow.com/questions/20913411/test-if-an-internet-connection-is-present-in-python
    '''
    try:
        host = socket.gethostbyname(hostname)
        s = socket.create_connection((host, 443), 2)
        s.close()
        return True 
    except Exception:
        pass
    return False

def get_sheet(sheet_name):
    ss_client = smartsheet.Smartsheet()
    ss_client.errors_as_exceptions(True)
    print(f'Sheet to get: {sheet_name}')
    search_result = ss_client.Search.search(sheet_name, scopes = 'sheetNames').results
    print(f'found {len(search_result)} results.')
    for r in range(len(search_result)):
        if (search_result[r].text) == sheet_name:
            sheet_id = next(result.object_id for result in search_result if result.object_type == 'sheet')
            # redundant? ss_client.Sheets.get_sheet(sheet_id)
            sheet = ss_client.Sheets.get_sheet(sheet_id)
            print(f'Found: {sheet.name}. ID is {sheet_id}')
            cols = ss_client.Sheets.get_columns(sheet_id)
            column_id = cols.data[0].id
            # delete print(type(sheet))
            print(column_id)
            make_fixture_names(sheet, sheet_id, column_id)

        else:
            print(f'No results found for "{sheet_name}". Did you use the exact name, and put it in quotes?')


def make_fixture_names(sheet, sheet_id, column_id):
    '''
    takes sheet_id and column_id, combines the DMX Line ID and zone number into a
    single string, and returns that string, for use as a fixture_name with make_csv()'''
    print(f'column_to_dict time. Using sheet ID {sheet_id} and column_id {column_id}.')
    fixture_names = []
    #zones = []
    for row in sheet.rows:
        if row.parent_id is None:
            parent_row = row.id
            for cell in row.cells:
                if cell.column_id == column_id:
                    if cell.value is not None:
                        cable_id = cell.value
        if row.parent_id is not None:
            for cell in row.cells:
                if cell.value is not None:
                    if cell.column_id == column_id:
                        zones.append(cell.value)
                        fixture_name = cable_id + ' - ' + cell.value
                        fixture_names.append(fixture_name)
    #print(f'Cable ID {cable_id} has {len(zones)} zones.')
    return(fixture_names)

def create_layout:
    pass

def make_csv(layout_name, fixture_name):
    with open(f'{layout_name}.csv', 'w', newline = '') as csv_file:
        field_names = ['Name',
                       'Fixture number',
                       'Groups',
                       'Notes',
                       'Manufacturer ID',
                       'Model ID',
                       'Mode ID',
                       'Width',
                       'Height',
                       'X',
                       'Y',
                       'Angle']
        mosaic_writer = csv.DictWriter(csv_file, fieldnames = field_names)
        mosaic_writer.writeheader()
        mosaic_writer.writerow({'Name': 'XN5.23.05',
                                'Fixture number' : '',
                                'Groups': '',
                                'Notes': 'auto-added!',
                                'Manufacturer ID': 0,
                                'Model ID': '12',
                                'Mode ID': '65',
                                'Width': 24,
                                'Height': 24,
                                'X': 24,
                                'Y': 24,
                                'Angle': 0})
        mosaic_writer.writerow({'Name': 'XN3.21.01',
                                'Fixture number' : '',
                                'Groups': '',
                                'Notes': 'wheeee',
                                'Manufacturer ID': 0,
                                'Model ID': '12',
                                'Mode ID': '65',
                                'Width': 24,
                                'Height': 24,
                                'X': 48,
                                'Y': 24,
                                'Angle': 0})

def extract_zones():
    pass

if __name__ == '__main__':
    '''
    takes one argument, which must be the full name of the sheet.
    If name contains spaces, enclose the name in quotes
    '''
    arguments = parser.parse_args()
    sheet_name = arguments.sheet_name
    print(sheet_name)
    type(arguments)
    if check_for_internet(REMOTE_SERVER):
        print("looks like we have internet. Let's continue.")
        get_sheet(sheet_name)
    else:
        sys.exit('No internet connection found. Exiting. (try again later?)')
    
    # make_csv(sheet_name, 'X5.01.01', 'z912b') # this is the way