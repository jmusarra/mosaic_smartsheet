"""
Uses Smartsheet's API to pull a sheet. Then extracts the zone numbers, then
uses the writes a .csv which can be imported into mosaic designer to create a layout.
   
Takes one argument, which must be the full name of the sheet.
If name contains spaces, enclose the name in quotes

TODO: this could be a web app.

TODO: use OAuth2 for access/authentication
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

import smartsheet
#TODO: make the fixture_types table
from fixture_types import generic_types
print(generic_types['RGB8'])

REMOTE_SERVER = 'smartsheet.com'
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
        return True
    except TimeoutError:
        print("The network connection timed out. Try agan later.")
        return False
    finally:
        s.close()

def get_sheet(sheet_name):
    """
    Gets the relevant info from the specified Smartsheet.

    Queries the Smartsheet API and searches for the string provided as the argument
    to mosaic-smartsheet.py. Uses that to get us a sheet_id, then takes the first
    column of that sheet as our Cable ID / zone information.
    """
    ss_client = smartsheet.Smartsheet()
    ss_client.errors_as_exceptions(True)
    print(f'Sheet to get: {sheet_name}')
    search_result = ss_client.Search.search(sheet_name, scopes = 'sheetNames').results
    print(f'found {len(search_result)} results.')
    #for r in range(len(search_result)):
    for i, r in enumerate(search_result):
        if (search_result[i].text) == sheet_name:
            sheet_id = next(result.object_id for result in search_result if result.object_type == 'sheet')
            sheet = ss_client.Sheets.get_sheet(sheet_id)
            print(f'Found: {sheet.name}. ID is {sheet_id}')
            cols = ss_client.Sheets.get_columns(sheet_id)
            column_id = cols.data[0].id
            get_from_smartsheet(sheet, sheet_id, column_id)
        else:
            print(f'''No results found for "{sheet_name}".
                   Did you use the exact name, and put it in quotes?''')


def get_from_smartsheet(sheet, sheet_id, column_id):
    """
    Gets cable ID and zone, combines them into a fixture name.

    Takes sheet_id and column_id, combines the Cable ID and zone number into a
    single string, and returns that string, for use as a fixture_name with create_layout()
    """
    #TODO: test the shit out of this
    print(f'Getting fixture names. Using sheet ID {sheet_id} and column_id {column_id}.')
    fixture_groups = {}
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
    num_fixtures = sum(len(f) for f in fixture_groups.values())
    num_lines = len(fixture_groups.keys())
    print(f'Found {num_fixtures} fixtures on {num_lines} DMX lines.')
    #create_fixture_rows(fixture_groups)
    make_fixtures_for_group(fixture_groups)

def make_fixtures_for_group(groups):
    """
    TODO: make docstring

    it should go here
    """
    names = []
    fixture_rows = []
    row = ['', '', '', '', 0, 0, 0, 24, 24, 0, 0, 0]
    group_num = 1
    for cable_id, zone in groups.items():
        #fixtures_per_line = (cable_id, len(zone))
        #print(f'Creating {len(zone)} lil squareys for {cable_id}...')
        group_num += 2
        for i in range(len(zone)):
            row = ['', '', '', '', 0, 12, 65, 24, 24, 0, 0, 0]
            position = [24 * (i+1), 24 * group_num]
            #if position[0] >= 481:
            #    position[0] = 24
            #    position[1] += group_num
             #   group_num += 1
            #print(f'Squarey {i + 1} at {position[0]},{position[1]}')
            row[9] = position[0]
            row[10] = position[1]
            fixture_rows.append(row)
    for i, row in enumerate(fixture_rows):
        print(f'{i+1}: {row}')
    print(f'Got {len(fixture_rows)} rows')
    print(fixture_rows[0])
    # Generate fixture names, add them to the rows:
    for cable_id, zone in groups.items():
        #print(f'{cable_id}: {len(zone)} fixtures')
        for z in zone:
            name = f'{cable_id} - {z}'
            names.append(name)
    print(f'Names: {len(names)}')
    print(f'Rows: {len(fixture_rows)}')
    for i, row in enumerate(fixture_rows):
        fixture_rows[i][0] = names[i]
    make_csv(fixture_rows)

def make_csv(fixture_rows):
    """Writes the generated rows to an excel-formatted CSV file"""
    print('Making a CSV!')
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
