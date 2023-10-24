'''Uses Smartsheet's API to pull a sheet. Then extracts the zone numbers, then
   uses the exported csv --or modifies in-place??-- the mosaic designer layout.

   Create .csv directly from smartsheet, and then import?

   this could be a web app.

   TODO: use OAuth2 for access/authentication

   take sheet_title as argument
'''
import socket
import argparse

import urllib
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

def get_sheet(sheet):
    print(f'Sheet to get: {sheet}')

def make_csv(dmx_line, zone_number):
    pass

def extract_zones():
    pass

if __name__ == '__main__':
    arguments = parser.parse_args()
    sheet = arguments.sheet_name
    print(sheet)
    type(arguments)
    if check_for_internet(REMOTE_SERVER):
        print("looks like we have internet. Let's continue.")
    else:
        sys.exit('No internet connection found. Exiting. (try again later?)')
    get_sheet(sheet)