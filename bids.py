#!/usr/bin/env python

'''
Handle bids from multi-unit auctions recorded in a Google sheet.

DONE 1. Compute the total amount.
DONE 2. Compute the winners and bid amounts for each item.
DONE 3. Produce a summary of all current winning bids and prices.
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

# No. 3
#CURRENT_NO = 3
#BID_RANGE = 'No. 3!A2:K'
#WON_RANGE = 'No. 3!K3:M'

# No. 4
#CURRENT_NO = 4
#BID_RANGE = 'No. 4!A2:K'
#WON_RANGE = 'No. 4!K3:M'

# No. 5
#CURRENT_NO = 5
#BID_RANGE = 'No. 5!A2:K'
#WON_RANGE = 'No. 5!K3:M'

# No. 6
#CURRENT_NO = 6
#BID_RANGE = 'No. 6!A2:K'
#WON_RANGE = 'No. 6!K3:M'

# No. 7
#CURRENT_NO = 7
#BID_RANGE = 'No. 7!A2:K'
#WON_RANGE = 'No. 7!K3:M'
#GOAL = 7500

# No. 8
#CURRENT_NO = 8
#BID_RANGE = 'No. 8!A2:K'
#WON_RANGE = 'No. 8!K3:M'
#GOAL = 6815

# No. 8
CURRENT_NO = 9
BID_RANGE = 'No. 9!A2:K'
WON_RANGE = 'No. 9!K3:M'
GOAL = 7500

END_DATE = 'July 1st'
#END_DATE = 'October 31st'

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

    [u'item', u'bidder_url', u'bidder_name', u'quantity', u'max_bid', u'bid_order', u'cancelled', u'lost', u'pending', u'pseudonym']

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
        if row == []:
            print "Skipping blank row."
            continue
        else:
            #print "Working on row:\n%s" % ( row )
            pass
        bid = {}
        for i in range( len( headers ) ):
            if headers[i] in types:
                conversion = types[headers[i]]
            else:
                conversion = str

            try:
                bid[headers[i]] = conversion( row[i] )
            except Exception as e:
                print "ERROR handling row:\n%s\nIndex: %s\nHeader: %s\nLed to:\n%s" % ( row, i, headers[i], e )
                raise


        bids.append( bid )

    return bids


def compute_winners( bids ):
    '''Given a list of bids from process_bids, compute winners.

    The quantity available is taken to be the quantity field of bidder
    pseudonym RESERVE.

    '''

    def winner_sort( b ):
        # Sort by descending max_bid, then increasing bid_order
        return ( -b['max_bid'], b['bid_order'] )

    # Prepare for analysis:
    for b in bids:
        b['won_quantity'] = 0

    quantities = { b['item'] : b['quantity'] for b in bids if b['pseudonym'] == 'RESERVE' }

    winners = {}

    running_total = 0

    for item in sorted( quantities.keys() ):
        item_bids = sorted( [ b for b in bids if b['item'] == item ], key=winner_sort )

        # Get both the winning bids, and the price paid.
        available = quantities[item]
        price = None
        for ib in item_bids:
            if ib['cancelled'] != '':
                # Skip cancelled bids.
                continue

            desired = ib['quantity']

            if available <= 0:
                ib['won_quantity'] = 0

                if price is None:
                    price = ib['max_bid']
            elif desired > available:
                ib['won_quantity'] = available
                available = 0
                price = ib['max_bid']
            else:
                ib['won_quantity'] = desired
                available -= desired

        for ib in item_bids:
            ib['current_price'] = price

        running_total += quantities[item] * price

        winners[item] = [ ib for ib in item_bids if ib['won_quantity'] > 0 ]

    return winners, running_total

WIN_FRONT = '''
Welcome to my Discount Lightning $8k Order Auction No. %d.

[b]Announcing your pseudonym at end of auction now gets you a $3 discount on shipping.[/b]

By now you know the drill: Faster, Cheaper, and high Transparency.  Details in Post #2.

[size=5][b]How bidding works:[/b][/size]

[ul]
  [li]A bid on a set of identical items is recorded as: (Bidder, Max Bid, Quantity, Time of Bid)[/li]
  [li]A sorted list of all bids for an item is maintained sorted by descending Max Bid, ascending Time[/li]
  [li]Winning bidders are allocated items by going down the sorted list and allocating items up to desired Quantity until all items have been allocated[/li]
  [li]The winning bid amount is the Max Bid amount of the first bid to not have their whole requested quantity
  [ul][li]All winning bidders pay the same amount[/li][/ul][/li]
[/ul]

I'll update the currently leading bidders in this post, quantities, and the current bid price by pseudonym as new bids come in.

If you'd like to learn more, see the example below in the second post on this thread.

[b]Threshold and Timelines:[/b]
[ul]
  [li]The auction will end as soon as I process a bid that puts the total value at or over $%s, including the value of all items still at their reserve price.[/li]
  [li]PyP choices are due at the time of payment.[/li]
  [li]If the $%s threshold is not met by %s the auction is not funded and will be cancelled.[/li]
  [li]Items will be mailed to you within 1 week after my receipt of them from True Adventures.  Be forewarned that some items (like adventure modules) are sent out by True Adventures much later than others.[/li]
[/ul]

[b]Shipping:[/b] Within the US: $8 per customer.  Outside the US: Actual shipping cost.  You may choose to announce your pseudonym after the auction in exchange for a $3 discount on shipping. If your total purchases in this auction exceeds $2000 I will cover your shipping for free.


[b]Payment:[/b] Is via PayPal and due within 3 days of me notifying you of your won bid.  The amount that reaches my PayPal account must be the sum of your won bids and shipping.

[size=6][b]Current Bids:[/b][/size]

''' % ( CURRENT_NO, GOAL, GOAL, END_DATE )

#NOTE: You may redeem 2 PyP selections for a complete C/UC/R ONYX Set.  You may redeem 18 PyP selections for a complete C/UC/R/UR ONYX Set.  Limit one such substitution per auction, priority will be given to winning bidders in order of highest bidder first, breaking ties on earliest bid.



WIN_BACK = '''
[b]The Fine Print:[/b]
See the next post for details on bidding and other corner case rules.
'''

def print_winners( winners, running_total ):
    '''Print a display of the winners.'''

    print WIN_FRONT

    print "$%0.02f of $%0.0f goal - %0.02f%% Funded\n" % ( running_total, GOAL, 100*running_total / GOAL )

    first = True
    for item in sorted( winners.keys() ):
        item_first = True
        row_message = "[u][b]%s[/b][/u]" % ( item )
        for wb in winners[item]:
            if item.startswith( '2020 ONYX UR' ):
                if first:
                    first = False
                    print "[u][b]2020 ONYX URs:[/b][/u]"
                row_message = "ONXY" + item[12:]
                print "%s : %s - $%0.02f" % ( row_message, wb['pseudonym'], wb['current_price'] )
            else:
                if item_first:
                    item_first = False
                    print row_message
                print "Qty. %d : %s - $%0.02f" % ( wb['won_quantity'], wb['pseudonym'], wb['current_price'] )
            if item.startswith( '2020 ONYX UR Tabor' ):
                print ""
        if not item.startswith( '2020 ONYX UR' ):
            print ""

    print WIN_BACK

def update_sheet( service, sheet, winners ):
    '''Only update the won_quantity, current_price, and old_won_quantity
    columns, assuming the sheet has stayed the same since we began the
    operation of the script.'''

    # Build up the values in the won_quantity, current_price, and old_won_quantity columns.
    result = []
    for row in sheet[1:]:
        # item and bid_order form a unique key.
        item = row[0]
        bid_order = int( row[5] )

        result_won = 0
        result_price = ''
        result_old_won = int( row[10] )

        for w in winners[item]:
            if w['bid_order'] == bid_order:
                result_won = w['won_quantity']
                result_price = w['current_price']

        result.append( [ result_won, result_price, result_old_won ] )

    request = service.spreadsheets().values().update(
        spreadsheetId = AUCTION_SHEET_ID,
        range = WON_RANGE,
        valueInputOption='RAW',
        body={ 'values' : result } )
    result = request.execute()

def main():
    # Get the auction sheet and current bids.

    service = auth()

    sheet = get_sheet( service, AUCTION_SHEET_ID, BID_RANGE )

    bids = process_bids( sheet )

    winners, running_total = compute_winners( bids )

    print_winners( winners, running_total )

    update_sheet( service, sheet, winners )

    print "Running total: %0.02f - %0.02f" % ( running_total, 100*running_total / GOAL )

if __name__ == '__main__':
    main()
