'''Uses Smartsheet's API to pull a sheet. Then extracts the zone numbers, then
   uses the exported csv --or modifies in-place??-- the mosaic designer layout.

   Create .csv directly from smartsheet, and then import?

   this could be a web app.

   TODO: use OAuth2 for access/authentication

   take sheet_title as argument
'''
import socket
import argparse

import smartsheet

parser = argparse.ArgumentParser()

REMOTE_SERVER = 'https://4wall.com'

def check_for_internet():
    '''
    see if we have a working internet connection
    copied from https://stackoverflow.com/questions/20913411/test-if-an-internet-connection-is-present-in-python
    '''
    try:
        host = socket.gethostbyname(hostname)
        s = socket.create_connection(host, 80), 2
        s.close()
        return True 
    except Exception:
        pass
    return False

def pull_from_smartsheet():

    pass

def extract_zones():
    pass


if __name__ == '__main__':
    parser.parse_args()