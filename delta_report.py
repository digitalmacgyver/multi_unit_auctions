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

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
AUCTION_SHEET_ID = '1b-cwze2D5X4WaheAWIXycDiR6ZGG0XDvhXEVCoAqxKY'
#BID_RANGE = 'No. 7!A2:M'
BID_RANGE = 'No. 8!A2:M'

AUCTION_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250702'

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
        'old_won_quantity' : int
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


def report_changes( bids ):
    bidders = sorted( { b['pseudonym'] : True for b in bids }.keys() )

    prices = { b['item'] : b['current_price'] for b in bids if b['current_price'] != '' }

    for bidder in bidders:
        bb = [ b for b in bids if ( b['pseudonym'] == bidder and b['cancelled'] == '' ) ]

        outbid = ""
        winning = ""

        issue_report = False

        for bid in sorted( bb, key=operator.itemgetter( 'item' ) ):
            won = bid['won_quantity']
            old_won = bid['old_won_quantity']
            if won < old_won or ( old_won == -1 and won < bid['quantity'] ):
                issue_report = True
                outbid += "%s : %d of %d (currently winning %d at $%s with max bid $%0.02f)\n" % ( bid['item'], bid['quantity']-won, bid['quantity'], won, prices[bid['item']], bid['max_bid'] )
            else:
                if won > old_won:
                    issue_report = True
                winning += "%s : winning %d of %d at $%s (max bid $%s)\n" % ( bid['item'], won, bid['quantity'], prices[bid['item']], bid['max_bid'] )

        if outbid != "":
            outbid = "You've been outbid on:\n" + outbid
            if winning != "":
                winning = "Status of your other bids:\n" + winning
        elif winning != "":
            winning = "Bid summary - you are winning:\n" + winning

        if issue_report:
            userid = bid['bidder_url'].split( '=' )[-1]
            contact_url = "https://truedungeon.com/component/uddeim/?task=new&recip=%s" % ( userid )
            print '='*80
            print "%s\nAuction update for %s\n\nIn auction: %s\n\n%s\n\n%s" % ( contact_url, bidder, AUCTION_URL, outbid, winning )


def main():
    # Get the auction sheet and current bids.
    service = auth()

    sheet = get_sheet( service, AUCTION_SHEET_ID, BID_RANGE )

    bids = process_bids( sheet )

    report_changes( bids )


if __name__ == '__main__':
    main()
