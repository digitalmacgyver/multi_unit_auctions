#!/usr/bin/env python

'''
Handle bids from multi-unit auctions recorded in a Google sheet.

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

# No. 4
#BID_RANGE = 'No. 4!A2:M'
#CURRENT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250559'
#NEXT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250572'
#AUCTION_NO = 4
#PAYMENT_DATE = 'October 23rd'
#SHIPPING_DISCOUNT = 8
#SHIPPING_COST = 8

# No. 5
#BID_RANGE = 'No. 5!A2:M'
#CURRENT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250572'
#NEXT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250584'
#AUCTION_NO = 5
#PAYMENT_DATE = 'October 26th'
#SHIPPING_COST = 8
#SHIPPING_DISCOUNT = 8

# No. 6
#BID_RANGE = 'No. 6!A2:M'
#CURRENT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250584'
#NEXT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250617'
#AUCTION_NO = 6
#PAYMENT_DATE = 'November 2nd'
#SHIPPING_COST = 8
#SHIPPING_DISCOUNT = 3

# No. 7
#BID_RANGE = 'No. 7!A2:M'
#CURRENT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250617'
#NEXT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250702'
#AUCTION_NO = 7
#PAYMENT_DATE = 'November 22nd'
#SHIPPING_COST = 8
#SHIPPING_DISCOUNT = 3

# No. 8
BID_RANGE = 'No. 8!A2:M'
CURRENT_URL = 'https://truedungeon.com/forum?view=topic&catid=584&id=250702'
AUCTION_NO = 8



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


def report_end( bids ):
    bidders = sorted( { b['pseudonym'] : True for b in bids }.keys() )

    prices = { b['item'] : b['current_price'] for b in bids if b['current_price'] != '' }

    for bidder in bidders:
        bb = [ b for b in bids if ( b['pseudonym'] == bidder and b['cancelled'] == '' ) ]

        won_message = ""
        lost_message = ""

        won_total = 0

        userid = bb[0]['bidder_url'].split( '=' )[-1]
        contact_url = "https://truedungeon.com/component/uddeim/?task=new&recip=%s" % ( userid )

        for bid in sorted( bb, key=operator.itemgetter( 'item' ) ):
            won = bid['won_quantity']
            lost = bid['quantity'] - won
            price = float( prices[bid['item']] )

            if won > 0:
                won_message += "%s : %d at $%0.02f = $%0.02f\n" % ( bid['item'], won, price, won*price )
                won_total += won*price

            if lost > 0:
                lost_message += "%s : %d with max_bid $%0.02f\n" % ( bid['item'], lost, bid['max_bid'] )

        if won_message != '':
            won_message = "You won the following:\n\n" + won_message
            won_message += 'foo'
        if lost_message != '':
            lost_message = 'foo'

        if won_message != '' or lost_message != '':
            print "="*80

            print "%s\nAuction Cancelled - %s.\n\nUnfortunately Auction No. 8 did not meet its funding goal by the deadline of midnight December 7th.\n\nBecause this auction includes 2020 treasure chips, which are only available as part of the 8k order during a limited preorder period, I am canceling this auction.\n\nThis is because here is no longer time to reach the goal, collect funds, and place the order before the deadline for 2020 treasure chips.\n\nThank you for your bids, I'm sorry this auction didn't succeed (if it were super close to the goal I would have gone ahead anyway, but it's over $200 away from the goal).\n\nI will likely start up another auction post PAX South.\n" % ( contact_url, bidder )
            


def main():
    # Get the auction sheet and current bids.
    service = auth()

    sheet = get_sheet( service, AUCTION_SHEET_ID, BID_RANGE )

    bids = process_bids( sheet )

    report_end( bids )


if __name__ == '__main__':
    main()
