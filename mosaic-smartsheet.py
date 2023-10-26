"""
Uses Smartsheet's API to pull a sheet. Then extracts the zone numbers, then
uses the exported csv --or modifies in-place??-- the mosaic designer layout.
   
Takes one argument, which must be the full name of the sheet.
   
If name contains spaces, enclose the name in quotes

Create .csv directly from smartsheet, and then import?

this could be a web app.

TODO: use OAuth2 for access/authentication

take sheet_title as argument
"""

__author__ = "John Musarra"
__license__ = "MIT"
__email__ = "john@mightymu.net"
__maintainer__ = "John Musarra"
__version__ = "alpha"

import socket
import argparse
import csv
import logging
import sys
from pprint import pprint

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
    copied from 
    https://stackoverflow.com/questions/20913411/test-if-an-internet-connection-is-present-in-python
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
    """
    Gets the relevant info from the specified Smartsheet.

    Queries the Smartsheet API and searches for the string provided as the argument
    to mosaic-smartsheet.py. Uses that to get us a sheet_id, then takes the first
    column as our Cable ID / zone information.
    """
    ss_client = smartsheet.Smartsheet()
    ss_client.errors_as_exceptions(True)
    print(f'Sheet to get: {sheet_name}')
    search_result = ss_client.Search.search(sheet_name, scopes = 'sheetNames').results
    print(f'found {len(search_result)} results.')
    for r in range(len(search_result)):
        if (search_result[r].text) == sheet_name:
            sheet_id = next(result.object_id for result in search_result if result.object_type == 'sheet')
            sheet = ss_client.Sheets.get_sheet(sheet_id)
            print(f'Found: {sheet.name}. ID is {sheet_id}')
            cols = ss_client.Sheets.get_columns(sheet_id)
            column_id = cols.data[0].id
            make_fixture_names(sheet, sheet_id, column_id)
        else:
            print(f'''No results found for "{sheet_name}".
                   Did you use the exact name, and put it in quotes?''')


def make_fixture_names(sheet, sheet_id, column_id):
    """
    Gets cable ID and zone, combines them into a fixture name.

    Takes sheet_id and column_id, combines the Cable ID and zone number into a
    single string, and returns that string, for use as a fixture_name with create_layout()
    """
    #TODO: test the shit out of this
    print(f'Getting fixture names. Using sheet ID {sheet_id} and column_id {column_id}.')
    fixture_groups = {}
    fixture_names = []
    id_count = 0
    cable_ids = []
    parent_ids = []
    zone_list = []
    #TODO: unfuck this:
    for row in sheet.rows:
        if row.parent_id is None:
            #row is a parent / cable ID
            for cell in row.cells:
                if cell.column_id == column_id:
                    if cell.value is not None:
                        parent_ids.append(row.id)
                        id_count += 1
                        cable_id = cell.value
                        if cable_id not in cable_ids:
                            cable_ids.append(cable_id)
                            fixture_groups[cable_id] = []
                        for p_id in parent_ids:
                            for row in sheet.rows:
                                if row.parent_id == p_id:
                                    for cell in row.cells:
                                        if cell.column_id == column_id:
                                            if cell.value is not None:
                                                zone_list.append(cell.value)
                                                fixture_groups[cable_id] = zone_list
                            zone_list = [] # reset the list of zone munbers to empty
    print(f'Got {len(zone_list)} fixtures')
    print(f'Created {len(fixture_names)} fixtures on {id_count} DMX lines.')
    #pprint(fixture_groups)
    create_fixture_rows(fixture_groups)

def make_groups(groups):
    """
    returns group_name and num_fixtures
    """
    group_names = []
    for num, name in list(enumerate(cable_ids, start = 1)):
       pass 
    groups = (1, 5)
    return groups

def create_fixture_rows(groups):
    """
    Creates the rows of fixture information to be CSV'ified.

    Takes the list of fixture names from Smartsheet, then adds the rest of the
    columns that Mosaic expects to create a full row.
    """
    
    default_fixture_number = ''
    default_fixture_width = 24
    default_fixture_height = 24
    default_angle = 0
    start_x = 0
    start_y = 0
    print('Creating Mosaic layout!')
    fixture_rows = []
    x = start_x
    y = start_y
    
    fixture_names = []
    # build all fixture names:
    for cable_id, zone in groups.items():
        #print(f'{cable_id}: {len(zone)} fixtures')
        for z in zone:
            fixture_names.append(f'{cable_id} - {z}')
    print(f'Created {len(fixture_names)} fixture names')

    # build all group_names
    group_names = []
    for num, name in enumerate(fixture_names, start = 1):
        group_names.append(f"{{{num},'{name}'}}")
    print(group_names[1])

    for f in fixture_names:
        fixture_row = [f,        #fixture name
                   '',           #fixture number - leave blank 
                   g,            #groups
                   '',           #notes
                   0,            #manufacturer id
                   12,           #model id
                   65,           #mode ID
                   24,           #width
                   24,           #height
                   x,            #x
                   y,            #y
                   0]            #angle
        x += default_fixture_width
        if x >= 481: #20 fixtures wide
            y += default_fixture_height
            x = start_x
        fixture_rows.append(fixture_row)
    print(f'Created {len(fixture_rows)} fixture rows.')
    make_csv(fixture_rows)
   
def make_csv(fixture_rows):
    """Writes the generated rows to an excel-formatted CSV file"""

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
    arguments = parser.parse_args()
    sheet_name = arguments.title
    print(sheet_name)
    type(arguments)
    if check_for_internet(REMOTE_SERVER):
        print("looks like we have internet. Let's continue.")
        get_sheet(sheet_name)
    else:
        sys.exit('No internet connection found. Exiting. (try again later?)')
