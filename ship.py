#!/usr/bin/env python

'''
Handle bids from multi-unit auctions recorded in a Google sheet.

5. Produce notifications at end of auction summarizing won items.

'''

import csv
import datetime
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

WON_RANGE = 'Shipping!A1:M'

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
                                range=WON_RANGE).execute()
    values = result.get('values', [])

    return values

def process_bids( sheet, cancelled=False ):
    '''Returns a list of dict, where each dict is a bid with key/value pairs like:

    [u'item', u'bidder_url', u'bidder_name', u'quantity', u'max_bid', u'bid_order', u'cancelled', u'lost', u'pending', u'pseudonym', u'won_quantity', u'current_price', u'old_won_quantity' ]

    If the cancelled parameter is True, we include all bids, otherwise
    we screen out bids where the cancelled field is populated.

    '''

    types = {
        'auction' : int,
        'won_quantity' : int,
        'won_bid' : float,
        'won_total' : float,
    }

    headers = sheet[0]

    bids = []

    for row in sheet[1:]:
        #print row
        bid = {}
        for i in range( len( headers ) ):
            if headers[i] in types:
                conversion = types[headers[i]]
            else:
                conversion = str

            try:
                bid[headers[i]] = conversion( row[i] )
            except Exception as e:
                bid[headers[i]] = None
                #raise Exception( "Header: %d of %s\nLed to\n%s" % ( i, headers[i], e ) )
                
        bids.append( bid )

    return bids


def report_end( bids ):
    bidders = sorted( { b['bidder_name'] : True for b in bids }.keys() )

    for bidder in bidders:
        address = ""
        item_counts = {}
        won_message = []
        shipping = {}
            
        bb = [ b for b in bids if b['bidder_name'] == bidder ]
        
        userid = bb[0]['bidder_url'].split( '=' )[-1]
        contact_url = "https://truedungeon.com/component/uddeim/?task=new&recip=%s" % ( userid )

        for bid in sorted( bb, key=operator.itemgetter( 'item' ) ):
            won_message.append( "Auction %s: %s" % ( bid['auction'], bid['message'] ) )

            if bid['item'] in item_counts:
                item_counts[bid['item']] += bid['won_quantity']
            else:
                item_counts[bid['item']] = bid['won_quantity']

            if bid['address']:
                address = bid['address']

        dt = texttable.Texttable()
        dt.set_cols_align( ['r', 'l'] )
        dt.set_cols_dtype( ['t', 't'] )
        dt.set_deco( texttable.Texttable.HEADER )
        line_items = [ [ 'Qty', 'Item' ] ]
        for ic in sorted( item_counts.keys() ):
            line_items.append( [ item_counts[ic], ic ] )
        dt.add_rows( line_items )

        # DEBUG - this is a nightmare - skip it - it will be faster to just print the ~50 labels one by one.
        '''
        shipping['Order ID (required)'] = bidder
        shipping['Order Date'] = datetime.datetime.today().strftime( '%m/%d/%Y' )
        # Dictates insurance?
        shipping['Order Value'] = 0
        shipping['Requested Service'] = 'Standard Shipping'

        # I need to ensure my addresses are formatted as:
        # full name\nstreet\ncity, state zip
        address_parts = address.split( '\n' )
        shipping['Ship To - Name'] = address_parts[0]
        shipping['Ship To - Address 1'] = address_parts[1]
        city, rest = address_parts[2].split( ', ' )
        state = rest[:2]
        zipcode = rest[3:]
        shipping['Ship To - State/Province'] = state
        shipping['Ship To - City'] = city
        shipping['Ship To - Postal Code'] = zipcode
        '''
        
        details = "\n".join( sorted( won_message ) )

        print "-"*80, "\n", contact_url, "\n", "Auction Items for: %s\n\n%s\n\n%s\n\nAuction Breakdown:\n%s\n\n" % ( bidder, address, dt.draw(), details )



def stamps_csv():
    # THIS DOESN'T DO ANYTHING
    '''Stamps CSV requires information in particular order format.

    Order ID (required),Order Date,Order Value,Requested Service,Ship To - Name,Ship To - Company,Ship To - Address 1,Ship To - Address 2,Ship To - Address 3,Ship To - State/Province,Ship To - City,Ship To - Postal Code,Ship To - Country,Ship To - Phone,Ship To - Email,Total Weight in Oz,Dimensions - Length,Dimensions - Width,Dimensions - Height,Notes - From Customer,Notes - Internal,Gift Wrap?,Gift Message
    123-456,6/14/1984,9.99,Standard Shipping,Joe Recipient,ExampleCo,123 Main Street,,,CA,Los Angeles,12345,US,555-555-5555,email@example.com,13,11.75,8.75,11.75,Example notes from customer,Example internal notes,TRUE,Example gift message text
    '''

    fields = [
        'Order ID (required)',
        'Order Date',
        'Order Value',
        'Requested Service',
        'Ship To - Name',
        'Ship To - Company',
        'Ship To - Address 1',
        'Ship To - Address 2',
        'Ship To - Address 3',
        'Ship To - State/Province',
        'Ship To - City',
        'Ship To - Postal Code',
        'Ship To - Country',
        'Ship To - Phone',
        'Ship To - Email',
        'Total Weight in Oz',
        'Dimensions - Length',
        'Dimensions - Width',
        'Dimensions - Height',
        'Notes - From Customer',
        'Notes - Internal',
        'Gift Wrap?',
        'Gift Message'
    ]

    
    
def main():
    # Get the auction sheet and current bids.
    service = auth()

    sheet = get_sheet( service, AUCTION_SHEET_ID, WON_RANGE )

    bids = process_bids( sheet )

    report_end( bids )


if __name__ == '__main__':
    main()
