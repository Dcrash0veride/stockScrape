from urllib.request import Request, urlopen
# noinspection PyUnresolvedReferences
from bs4 import BeautifulSoup as Soup, Comment
from auth import drive_service
from auth import spreadsheet_service
from googleapiclient.http import MediaFileUpload
import re
import argparse
import csv


""" This script makes a request to finviz with screener settings included into
the URL and returns a list of tickers Matching the criteria.
Make a request to the screener, must stipulate user agent or mod_security
gives us the ol' 403"""


# Global Variables
root_url = "https://finviz.com/"

# Initialize empty ticker list
ticker_list = []

# next page list outside of loop
next_pages = []

# Scanner Settings for 52wkH 52wkL and near high and low

""" New High screen settings for average volume, share price,
and of course new high"""

new_high = ("screener.ashx?v=111&f=geo_usa,sh_avgvol_o100,sh_price_o5,"
            "sh_relvol_o1,ta_highlow52w_nh&ft=4&o=-change&ar=180")

""" New Low settings for average volume,share price, below 200sma,
and of course new low"""

new_low = ("screener.ashx?v=111&f=sh_avgvol_o100,sh_price_o5,ta_highlow52w_nl"
           ",ta_sma200_pb&ft=4&o=-change")

""" Near High settings to filter large amounts of results above all sma, price
over 5, avg vol over 100K rvol over 1.5 USA origin
"""

near_high = ("screener.ashx?v=111&f=geo_usa,"
             "sh_avgvol_o100,sh_price_o5,sh_relvol_o1.5,ta_highlow52w_b0to3h,"
             "ta_sma20_pa,ta_sma200_pa,ta_sma50_pa&ft=4&o=-change")

# Near low settings to filter only the highest quality shitty stocks to trade
near_low = ("screener.ashx?v=111&f=geo_usa,sh_avgvol_o100,sh_price_o5,"
            "sh_relvol_o1.5,ta_highlow52w_a0to5h,ta_sma20_pb,"
            "ta_sma200_pb,ta_sma50_pb&ft=4&o=-change")

# Ace scan for mid cap
ace_scan = ("screener.ashx?v=111&f=cap_mid,sh_avgvol_o100,sh_price_o5,"
            "sh_relvol_o1.5,ta_sma20_pa,ta_sma200_pa,ta_sma50_pa&ft=4")

# Header Spoof to get around the 403 from mod_security
header_spoof = {'User-Agent': 'Mozilla'}

# Parser set up for to accept arguments
"""TODO:ADD ALL Option, will probably be used the most"""
parser = argparse.ArgumentParser(description="Scan Stocks Matching Criteria")
group = parser.add_mutually_exclusive_group()
group.add_argument('-nh', help="Scan for stocks with new 52wk High", action="store_true")
group.add_argument('-nl', help="Scan for stock with new 52wk low", action='store_true')
group.add_argument('-nrh', help="Scan for stocks approaching 52wk High", action='store_true')
group.add_argument('-nrl', help="Scan for stocks approaching 52wk Low", action='store_true')
group.add_argument('-ace', help="Scan for stocks that might meet ACE criteria", action='store_true')
args = parser.parse_args()

"""Create a function to make a request to the root_url with specified settings
and create a soup object
"""


def make_request(s_url):
    req = Request(s_url, headers=header_spoof)
    uclient = urlopen(req)
    page_html = uclient.read()
    uclient.close()
    page_soup = Soup(page_html, "html.parser")
    soup_pre_processing(page_soup)


def soup_pre_processing(s_page):
    # Find comments in page
    comment = s_page.findAll(text=lambda text: isinstance(text, Comment))
    # Find number of pages of results
    getnext = s_page.findAll('a', {'class': 'screener-pages'})
    # If more than one page append the links to next_pages
    if not ticker_list:
        i = 0
        while i < len(getnext):
            next_pages.append(getnext[i]['href'])
            i += 1
    make_ticker_list(comment)


def make_ticker_list(s_comment):
    comment_soup = Soup(str(s_comment), 'lxml')
    c_pattern = r"TS\\n(.*?)\\nTE"
    c_search = re.search(c_pattern, str(comment_soup))
    tl_unformatted = c_search.group(1)
    t_pattern = r"[A-Z]{3,4}"
    t_result = re.findall(t_pattern, tl_unformatted)
    for ticker in t_result:
        if ticker not in ticker_list:
            ticker_list.append(ticker)
            if next_pages:
                next_url = root_url + next_pages[0]
                next_pages.pop(0)
                make_request(next_url)
#    create_json(ticker_list)
#    write_csv(ticker_list)

def create_json(ticker_l):
    range_ = '"range": "Sheet1!A1:J1000",'
    dimension = '"majorDimension": "ROWS",'
    s_header = ['ticker', 'MA150D', 'MA250D' 'CpriceAbove150200dma',
                '150dmaOver200dma', '200dmaTrUP1mnth',
                '50dmaAbove150200dma', 'CpriceAbove50dma',
                'Cprice30pctAbv52wkl', 'Cpricewthn25pct52wkh',
                'RsratingOver70']
    value_request_body = [

        [

        ]
    ]
    value_request_body.append(s_header)
    row_counter = 2
    for ticker in ticker_l:
        ticker_column = "A"+str(row_counter)
        ma150d_column = "B" + str(row_counter)
        ma250d_column = "C" + str(row_counter)
        data = [ticker, '=average(query(GOOGLEFINANCE('+ ticker_column +
                            ',"price",WORKDAY(TODAY(),-150),TODAY()),'
                            '"Select Col2 order by Col1 DESC"))',
                            '=average(query(GOOGLEFINANCE('+ ticker_column +
                            ',"price",WORKDAY(TODAY(),-200),TODAY()),'
                            '"Select Col2 order by Col1 DESC"))',
                            '=if(GOOGLEFINANCE('+ ticker_column +
                            ',"price")>' + ma150d_column + ',"TRUE")',
                            '=if(' + ma150d_column + '>' + ma250d_column +
                            ',"TRUE")', '=if(average(query(GOOGLEFINANCE('
                            + ticker_column + ',"price",WORKDAY(TODAY(),-200),'
                                              'TODAY()-30),'
                            '"Select Col2 order by Col1 DESC"))<'
                            + ma250d_column + ',"TRUE")', '=if(average(query'
                                                         '(GOOGLEFINANCE('
                            + ticker_column + ',"price",WORKDAY(TODAY(),-50),'
                                              'TODAY()),"Select Col2 order by '
                                              'Col1 DESC"))>' + ma150d_column +
                            ',"TRUE")','=if(average(query(GOOGLEFINANCE('
                            + ticker_column + ',"price",WORKDAY(TODAY(),-50),'
                                              'TODAY()),"Select Col2 order by Col1'
                                              ' DESC"))<GOOGLEFINANCE('
                            + ticker_column + ',"price"),"TRUE")']
        value_request_body.append(data)
        row_counter += 1
    tempvar = {
        "majorDimension": "ROWS",
        "values":value_request_body
    }
    sheet = spreadsheet_service.spreadsheets()
    sheet.values().append(
        spreadsheetId='1_hqi1Bbz0de8DmrRdwT4wXgynJ7sDAUs1KP7FgG7Cmw',
        range="Sheet1!A1:J150",
        valueInputOption='USER_ENTERED',
        body=tempvar).execute()


def write_csv(ticker_l):
    print("executing")
    with open('C:\\Users\\dcrash0veride\\Documents\\output_\\new.csv',
              'w', newline='') as f:
        f_header = ['Ticker', 'Cprice', 'MA150D', 'MA200D', 'CpriceAbove150200dma',
                    '150dmaOver200dma', '200dmaTrUP1mnth',
                    '50dmaAbove150200dma', 'CpriceAbove50dma',
                    'Cprice30pctAbv52wkl', 'Cpricewthn25pct52wkh',
                    'RsratingOver70', '52wkL', '52wkH']

        writer = csv.writer(f)
        writer.writerow(f_header)
        row_counter = 2
        for ticker in ticker_l:
            ticker_column = "A"+str(row_counter)
            ma150d_column = "C" + str(row_counter)
            ma200d_column = "D" + str(row_counter)
            weeklow_column = "M" + str(row_counter)
            weekhigh_column = "N" + str(row_counter)
            data = [ticker, '=GOOGLEFINANCE(' + ticker_column + ',"price")',
                    '=average(query(GOOGLEFINANCE('
                    + ticker_column + ',"price",WORKDAY(TODAY(),-150)'
                    ',TODAY()),"Select Col2 order by Col1 DESC"))',
                    '=average(query(GOOGLEFINANCE(' + ticker_column +
                    ',"price",WORKDAY(TODAY(),-200),TODAY()),'
                    '"Select Col2 order by Col1 DESC"))',
                    '=if(GOOGLEFINANCE(' + ticker_column +
                    ',"price")>' + ma150d_column + ',"TRUE")',
                    '=if(' + ma150d_column + '>' + ma200d_column +
                    ',"TRUE")', '=if(average(query(GOOGLEFINANCE('
                    + ticker_column + ',"price",WORKDAY(TODAY(),-200),'
                                      'TODAY()-30),'
                                      '"Select Col2 order by Col1 DESC"))<'
                    + ma200d_column + ',"TRUE")', '=if(average(query'
                                                  '(GOOGLEFINANCE('
                    + ticker_column + ',"price",WORKDAY(TODAY(),-50),'
                    'TODAY()),"Select Col2 order by Col1 DESC"))>'
                    + ma150d_column + ',"TRUE")', '=if(average(query'
                    '(GOOGLEFINANCE(' + ticker_column + ',"price"'
                    ',WORKDAY(TODAY(),-50),TODAY()),"Select Col2 order '
                    'by Col1 DESC"))<GOOGLEFINANCE('
                    + ticker_column + ',"price"),"TRUE")',
                    '=IF(GOOGLEFINANCE(' + ticker_column + ',"price")'
                    '>(' + weeklow_column + ' *0.3),"TRUE")',
                    '=IF(GOOGLEFINANCE(' + ticker_column + ',"price")'
                    '>' + weekhigh_column + '-(' + weekhigh_column + ''
                    ' * 0.25),"TRUE")', 'NULL', '=GOOGLEFINANCE('
                    + ticker_column + ', "low52")', '=GOOGLEFINANCE('
                    + ticker_column + ', "high52")']
            writer.writerow(data)
            row_counter += 1


def export_csv_file(csv_file):
    folder_id = '12-BjIS-Dn5JC0OAUISI0w3u9-AyRQtnL'
    file_meta = {
        'name': 'Trading Ticker List',
        'mimeType': 'application/vnd.google-apps.spreadsheet',
        'parents': [folder_id]
        }
    media = MediaFileUpload(csv_file, mimetype='text/csv')
    file = drive_service.files().create(body=file_meta,
                                        media_body=media,
                                        fields='id').execute()
    print('File ID: %s' % file.get('id'))

def read_sheet_values(spreadsheet_id):
    sheet = spreadsheet_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range='Trading Ticker List').execute()
    real_range = result.get('range', [])
    real_result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                     range=real_range).execute()

    values = real_result.get('values', [])
    print(values)

make_request(root_url+new_low)
create_json(ticker_list)


