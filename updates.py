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
AUCTION_SHEET_ID = '1b-cwze2D5X4WaheAWIXycDiR6ZGG0XDvhXEVCoAqxKY'
#BID_RANGE = 'No. 6!A2:M'
#BID_RANGE = 'No. 7!A2:M'
BID_RANGE = 'No. 8!A2:M'

GOAL = 6815

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
    result = sheet.values().get(spreadsheetId=AUCTION_SHEET_ID,
                                range=BID_RANGE).execute()
    values = result.get('values', [])

    return values

def process_bids( sheet, cancelled=False ):
    '''Returns a list of dict, where each dict is a bid with key/value pairs like:

    [u'item', u'bidder_url', u'bidder_name', u'quantity', u'max_bid', u'bid_order', u'cancelled', u'lost', u'pending', u'pseudonym', u'won_quantity', u'current_price', u'old_won_quantity' ]

    If the cancelled parameter is True, we include all bids, otherwise
    we screen out bids where the cancelled field is populated.

    '''

    types = {
        'quantity' : int,
        'max_bid' : float,
        'bid_order' : int,
        'won_quantity' : int,
        'old_won_quantity' : int,
    }

    headers = sheet[0]

    bids = []

    for row in sheet[1:]:
        bid = {}
        for i in range( len( headers ) ):
            if headers[i] in types:
                conversion = types[headers[i]]
            else:
                conversion = str

            bid[headers[i]] = conversion( row[i] )

        bids.append( bid )

    return bids

def get_deals( prices ):
    limits = {
        '1000 GP Bar' : 15.01,
        '2020 3x Treasure Chip' : 9.01,
        '2020 ONYX C/UC/R Set' : 75.01,
        '2020 or 2019 UR of Choice' : 100,
        "Alchemist's Ink" : 3.01,
        "Alchemist's Parchment" : 3.51,
        "Aragonite" : 22.51,
        "Darkwood Plank" : 1,
        "Dwarven Steel" : 3.01,
        "Elven Bismuth" : 12.01,
        "Minotaur Hide" : 2,
        "Mystic Silk" : 1,
        "Philosopher's Stone" : 1,
        "Wish Ring" : 225
    }

    result = ""
    onyx_count = 0
    for onyx in [ i for i in prices.keys() if i.startswith( '2020 ONYX UR ') ]:
        if float( prices[onyx] ) < 100:
            onyx_count += 1
    if onyx_count > 0:
        result += "[b]%d ONYX URs under $100[/b]\n" % ( onyx_count )

    dt = texttable.Texttable()
    dt.set_cols_align( ['l', 'r'] )
    dt.set_cols_dtype( [ 't', 'f' ] )
    dt.set_precision( 2 )
    dt.set_deco( texttable.Texttable.HEADER )
    deals = []

    if float( prices['2020 or 2019 UR of Choice'] ) < 95.01:
        deals.append( [ "PyPs", float( prices['2020 or 2019 UR of Choice'] ) ] )

    for item in sorted( prices.keys() ):
        if 'UR' in item:
            continue
        if item in limits and float( prices[item] ) < limits[item]:
            deals.append( [ item, float( prices[item] ) ] )

    if deals:
        dt.add_rows( deals, header=False )
        result = "Plenty of deals left based on current bids of:\n" + result + "[code]" + dt.draw() + "[/code]"

    return result


def report_changes( bids ):
    bidders = sorted( { b['pseudonym'] : True for b in bids }.keys() )

    prices = { b['item'] : b['current_price'] for b in bids if b['current_price'] != '' }

    changed = { b['item'] : True for b in bids if b['old_won_quantity'] == -1 and b['cancelled'] == '' }

    message = "[b]NOTE: If this auction doesn't fund by December 7th I will need to close it early as I need time to collect payment and place the order before the 3x Treasure Chips early order reward is still available.[/b]\n\n Updated winning bids for:\n\n"

    for item in sorted( changed.keys() ):
        message += "%s\n" % ( item )

    total_value = 0
    for b in bids:
        if b['cancelled'] == '' and b['won_quantity'] > 0:
            total_value += int( b['won_quantity'] ) * float( b['current_price'] )



    message += "\n\n%0.0f%% Funded\n\n" % ( 100*float( total_value ) / GOAL )

    message += get_deals( prices )

    print message


def main():
    # Get the auction sheet and current bids.
    service = auth()

    sheet = get_sheet( service, AUCTION_SHEET_ID, BID_RANGE )

    bids = process_bids( sheet )

    report_changes( bids )


if __name__ == '__main__':
    main()
