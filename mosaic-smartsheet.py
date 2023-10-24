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

# import urllib
import smartsheet

parser = argparse.ArgumentParser(
                    prog = 'mosaic-smartsheet',
                    description = 'generate mosaic layout from smartsheet',
                    epilog = 'some text here I guess')
parser.add_argument('sheet_name', 
                    help = '''The name of a Smartsheet sheet. Enclose with quotes if
                            the sheet name contains spaces''')

REMOTE_SERVER = '4wall.com'

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
    #print(type(search_result))
    print(f'found {len(search_result)} results.')
    for r in range(len(search_result)):
        if (search_result[r].text) == sheet_name:
            sheet_id = next(result.object_id for result in search_result if result.object_type == 'sheet')
            ss_client.Sheets.get_sheet(sheet_id)
            print(f'Sheet ID is {sheet_id}')
            sheet = ss_client.Sheets.get_sheet(sheet_id)
            print(f'Found: {sheet.name}')
            print(type(sheet))
            return sheet
            cols = ss_client.Sheets.get_columns(sheet_id)
            print(type(cols))
            print(cols[0])
    else:
        print(f'No results found for "{sheet_name}"')

def make_csv(sheet_name, dmx_line, zone_number):
    print('oh boooyyy csv!')
    with open(f'{sheet_name}.csv', 'w', newline = '') as csv_file:
        field_names = ['Name'
                       'Fixture number'
                       'Groups'
                       'Notes'
                       'Manufacturer ID'
                       'Model ID'
                       'Mode ID'
                       'Width'
                       'Height'
                       'X'
                       'Y'
                       'Angle']
        mosiac_writer = csv.writer(csv_file, dialect = 'excel')
        mosiac_writer.writerow({'Name': 'XN5.23.05',
                                'Fixture number' : 1001,
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
        csv_file.close()

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
    else:
        sys.exit('No internet connection found. Exiting. (try again later?)')
    get_sheet(sheet_name)
    make_csv(sheet_name, 'X5.01.01', 'z912b')