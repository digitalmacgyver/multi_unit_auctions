#!/usr/bin/env python

'''
Handle bids from multi-unit auctions recorded in a Google sheet.

4. Produce notifications for people who've been outbid.
5. Produce notifications at end of auction summarizing won items.

'''

import csv
import operator
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import texttable

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SHEET_ID = '10Q-6Nz1Eg5QO00Pu6bMkBAApLhx8uA2o4j7XQbgasPw'
LOOT_RANGE = 'Tokens!A1:I'

def auth():
    """Get login credentials done (opens browser tab for interactive
    credential auth.

    Reads from sheet_id with optional bid_range

    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    return service

def get_sheet( service, sheet_id, sheet_range ):
    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SHEET_ID,
                                range=LOOT_RANGE).execute()
    values = result.get('values', [])

    return values

def process_loot( sheet ):
    '''Returns a list of dict, where each dict is a bid with key/value pairs like:

    [u'available', 'item', 'rarity', 'trade', 'units', 'price', 'category', 'trade_value', 'desc' ]

    '''

    types = {
        'available' : int,
        'units' : int,
        'price' : float,
        'trade_value' : float
    }

    headers = sheet[0]

    items = []

    for row in sheet[1:]:
        item = {}
        for i in range( len( headers ) ):
            if headers[i] in types:
                conversion = types[headers[i]]
            else:
                conversion = str

            item[headers[i]] = conversion( row[i] )

        items.append( item )

    return items


def seshat( words ):
    return "Seshat: [i][color=darkblue] " + words + " [/color][/i]"

def yumphak( words ):
    return "Yumphak: [b][color=purple] " + words + " [/color][/b]"

def loot_tables( items ):
    #print items

    wondrous = list( reversed( sorted( [ i for i in items if i['category'] == 'Wondrous' ], key=operator.itemgetter( 'trade_value' ) ) ) )

    i = None
    rows = [ [ 'Roll', "Yumphak's Label", "Seshat's Notes" ] ]
    for i in range( len( wondrous ) ):
        slot = 100 - i
        w = wondrous[i]
        rows.append( [ slot, w['item'], w['desc'] ] )

    for i in range( 100 - len( wondrous ), 0, -1 ):
        if i == 13:
            rows.append( [ i, 'Roll again in this table', "Or, you could always take 2 things from the Wondrous Item table, it's not too late!" ] )
        else:
            rows.append( [ i, 'Roll again in this table', '' ] )
            
        
    dt = texttable.Texttable( max_width=69 )
    dt.set_cols_align( [ 'l', 'l', 'l' ] )
    dt.set_deco( texttable.Texttable.HEADER | texttable.Texttable.HLINES )
    dt.add_rows( rows )

    w_title = '[center][size=5] Worthless Mathom Loot Table - 3600 Forge Credits Per Try [/size][/center]\n'
    
    w_front = '''
May I suggest you instead pick any two Wondrous items instead?

Oh well... if you insist.  Don't say I didn't warn you if you can't unload this junk on some rich noble.

As you peruse the catalog below, please focus on my accurate and truthful descriptions of the items.  

I'm, uh, [b]bound[/b] to include Grade 3 Yumphak's labels - he really likes to title his little works - but I think you'll find my notes much more illuminating.

Remember you can pick any non-Roll Again item less than or equal to your roll.
    '''

    w_front = seshat( w_front ) + "\n"
    
    print w_title, w_front, "[code]", dt.draw(), "[/code]"


    mundane = list( reversed( sorted( [ i for i in items if i['category'] == 'Mundane' ], key=operator.itemgetter( 'trade_value' ) ) ) )

    i = None
    rows = [ [ 'Roll', "Yumphak's Label" ] ]
    for i in range( len( mundane ) ):
        slot = 100 - i
        m = mundane[i]
        rows.append( [ slot, m['item'] ] )

    for i in range( 100 - len( mundane ), 0, -1 ):
        rows.append( [ i, 'Roll again in this table' ] )
            
        
    dt = texttable.Texttable( max_width=69 )
    dt.set_cols_align( [ 'l', 'l' ] )
    dt.set_deco( texttable.Texttable.HEADER )
    dt.add_rows( rows )

    m_title = ''' [center][size=5] Wondrous Item Loot Table - 60 Forge Credits Per Try [/size][/center]

Seshat: [i][color=darkblue]Remember, you can pick any non-Roll Again item less than or equal to the value you roll.[/color][/i]

'''
    
    print m_title, "[code]", dt.draw(), "[/code]"
    
    

def main():
    # Get the auction sheet and current bids.
    service = auth()

    sheet = get_sheet( service, SHEET_ID, LOOT_RANGE )

    items = process_loot( sheet )

    loot_tables( items )


if __name__ == '__main__':
    main()
