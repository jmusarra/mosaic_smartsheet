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
import sys

# import urllib
import smartsheet
#TODO: make the fixture_types table
import fixture_types

REMOTE_SERVER = '4wall.com'
TESTING_SHEET_ID = 8769876857776004
TESTING_COLUMN_ID = 1182995975038852

logging.basicConfig(filename = 'mosaic-smartsheet.log', level=logging.INFO)

parser = argparse.ArgumentParser(
                    prog = 'mosaic-smartsheet',
                    description = 'generate mosaic layout from smartsheet',
                    epilog = 'some text here I guess')
parser.add_argument('title',
                    help = '''The title of a Smartsheet sheet. Enclose with quotes if
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
        #TODO: actually catch and handle socket exceptions
        # "except: pass" you fuckin yuntz
        pass
    return False

def get_sheet(sheet_name):
    '''
    queries the Smartsheet API. Searches for the string provided as the argument
    to mosaic-smarthseet.py. Uses that to get us a sheet_id, then takes the first
    column as our Cable ID / zone information.
    '''
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
            #print(column_id)
            make_fixture_names(sheet, sheet_id, column_id)

        else:
            print(f'No results found for "{sheet_name}". Did you use the exact name, and put it in quotes?')


def make_fixture_names(sheet, sheet_id, column_id):
    #TODO: test the shit out of this
    '''
    takes sheet_id and column_id, combines the DMX Line ID and zone number into a
    single string, and returns that string, for use as a fixture_name with create_layout()'''
    print(f'Getting fixture names. Using sheet ID {sheet_id} and column_id {column_id}.')
    fixture_names = []
    id_count = 0
    fixture_count = 0
    for row in sheet.rows:
        if row.parent_id is None:
            #row is a parent / cable ID
            for cell in row.cells:
                if cell.column_id == column_id:
                    if cell.value is not None:
                        id_count += 1
                        cable_id = cell.value
        if row.parent_id is not None:
            for cell in row.cells:
                if cell.value is not None:
                    #row is a child / zone number
                    if cell.column_id == column_id:
                        fixture_count += 1
                        fixture_name = cable_id + ' - ' + cell.value
                        fixture_names.append(fixture_name)
    print(f'Created {len(fixture_names)} fixtures on {id_count} cable IDs.')
    create_fixture_rows(fixture_names)

def create_fixture_rows(fixture_names):
    '''
    takes the list of fixture names from Smartsheet, then adds the rest of the
    expected information to create a full row.
    Expected columns:
    '''
    default_fixture_number = ''                  
    default_fixture_width: 24
    default_fixture_height: 24
    default_angle = 0
    print('Creating Mosaic layout!')
    fixture_rows = []
    for f in fixture_names:
        #build a list of lists, with the items from make_fixture_names as the first elements
            fixture_row = [f,            #fixture name
                           '',           #fixture number - leave blank 
                           '',           #groups
                           '',           #notes
                           0,            #manufacturer id
                           12,           #model id
                           65,           #mode ID
                           24,           #width
                           24,           #height
                           24,           #x
                           24,           #y
                           0]            #angle
            fixture_rows.append(fixture_row)
    print(f'Created {len(fixture_rows)} fixture rows.')
    make_csv(fixture_rows)

def make_csv(fixture_rows):
    '''
    docstring
    '''
    layout_name = sheet_name
    with open(f'{layout_name}.csv', 'w', newline = '', encoding='cp1252') as csv_file:
        header = ['Name',
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
        mosaic_writer = csv.writer(csv_file, dialect = 'excel')
        mosaic_writer.writerow(header)
        for row in fixture_rows:
            mosaic_writer.writerow(row)
      
if __name__ == '__main__':
    '''
    takes one argument, which must be the full name of the sheet.
    If name contains spaces, enclose the name in quotes
    '''
    arguments = parser.parse_args()
    sheet_name = arguments.title
    print(sheet_name)
    type(arguments)
    if check_for_internet(REMOTE_SERVER):
        print("looks like we have internet. Let's continue.")
        get_sheet(sheet_name)
    else:
        sys.exit('No internet connection found. Exiting. (try again later?)')
# make_csv(sheet_name, 'X5.01.01', 'z912b') # this is the way
