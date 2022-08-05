from urllib.request import Request, urlopen
# noinspection PyUnresolvedReferences
from bs4 import BeautifulSoup as Soup, Comment
import drive_driver
from drive_driver import *
import re


""" This script makes a request to finviz with screener settings included into
the URL and returns a list of tickers Matching the criteria.
Make a request to the screener, must stipulate user agent or mod_security
gives us the ol' 403"""

"""This is stock scrape V2, enough changes are about to be made to qualify
 as a major revision, v3 is in planning stages"""


def main():

    root_url = "https://finviz.com/"

    scanner_settings = {'new_high': ["screener.ashx?v=111&f=geo_usa,"
                                    "sh_avgvol_o100,sh_price_o5,"
                                    "sh_relvol_o1,"
                                    "ta_highlow52w_nh&ft=4&o="
                                    "-change&ar=180"],
                        'new_low': ["screener.ashx?v=111&f=sh_avgvol_o100,"
                                   "sh_price_o5,ta_highlow52w_nl"
                                   ",ta_sma200_pb&ft=4&o=-change"],
                        'near_high': ["screener.ashx?v=111&f=geo_usa,"
                                     "sh_avgvol_o100,"
                                     "sh_price_o5,"
                                     "sh_relvol_o1.5,"
                                     "ta_highlow52w_b0to3h,"
                                     "ta_sma20_pa,"
                                     "ta_sma200_pa,"
                                     "ta_sma50_pa&ft=4&o=-change"],
                        'near_low': ["screener.ashx?v=111&f=geo_usa,"
                                    "sh_avgvol_o100,"
                                    "sh_price_o5,"
                                    "sh_relvol_o1.5,"
                                    "ta_highlow52w_a0to5h,"
                                    "ta_sma20_pb,"
                                    "ta_sma200_pb,"
                                    "ta_sma50_pb&ft=4&o=-change"],
                        'ace_scan': ["screener.ashx?v=111&f=cap_mid,"
                                    "sh_avgvol_o100,"
                                    "sh_price_o5,"
                                    "sh_relvol_o1.5,"
                                    "ta_sma20_pa,"
                                    "ta_sma200_pa,"
                                    "ta_sma50_pa&ft=4"]}


    header_spoof = {'User-Agent': 'Mozilla'}


    def scrape(url):
        req = Request(url, headers=header_spoof)
        uclient = urlopen(req)
        page_html = uclient.read()
        uclient.close()
        page_soup = Soup(page_html, "html.parser")
        return page_soup

    def soup_comment(p_soup):
        comment = p_soup.findAll(text=lambda text: isinstance(text, Comment))
        return comment

    def soup_pages(p_soup):
        next_pages = p_soup.findAll('a', {'class': 'screener-pages'})
        return next_pages

    def comment_cleaner(dirty_c):
        comment_soup = Soup(str(dirty_c), 'lxml')
        c_pattern = r"TS\\n(.*?)\\nTE"
        c_search = re.search(c_pattern, str(comment_soup))
        tl_unformatted = c_search.group(1)
        t_pattern = r"[A-Z]{3,4}"
        t_result = re.findall(t_pattern, tl_unformatted)
        return t_result

    """At this point I think I have broken the functions down into
    pretty small chunks and just need to write a function to drive this 
    appropriately"""

    def url_builder(scanner_settings):
        for k,v in scanner_settings.items():
            last_supper = scrape(root_url + str(v[0]))
            if soup_pages(last_supper):
                for nxt_page in soup_pages(last_supper):
                    new_rl = nxt_page['href']
                    scanner_settings[k].append(new_rl)
        return scanner_settings



    daily_scan = url_builder(scanner_settings)
    t_dict = {}

    def build_ticker_dict():
        for k, v in daily_scan.items():
            index = 0
            t_dict[k] = []
            while index < len(v):
                ticker_list = comment_cleaner(soup_comment(scrape(root_url + v[index])))
                t_dict[k].append(ticker_list)
                index += 1

    """Okay so I ended up just building more functions and not really driver code the way
    I had imagined, but we now have a series of functions that return a dictionary of ticker symbols
    along with the scan the generated them. The next step is to create a spreadshet and add tabs
    for each of the scans and populate the sheet. SO this is the FINAL PART WOOT!"""

    sheet_obj = drive_driver.Sheet()

""" Note to self, you removed your email when you started tracking this in GIT
if you forgot and stuff isn't working that is probably why it should be in
shared_with= field"""
    sh_id = sheet_obj.create('Trading Scan Results', shared_with='')

    build_ticker_dict()


    for k, v in t_dict.items():
        sheet_obj.create_new_tab(sh_id, k)
        temp_val = []
        outer_list_length = len(v)
        for i in range(0, outer_list_length):
            inner_list_length = len(v[i])
            for s in range(0, inner_list_length):
                temp_val.append(v[i][s])
        s_header = ['ticker', 'Cprice', 'MA150D', 'MA200D',
                    'CpriceAbove150200dma',
                    '150dmaOver200dma', '200dmaTrUP1mnth',
                    '50dmaAbove150200dma', 'CpriceAbove50dma',
                    'Cprice30pctAbv52wkl', 'Cpricewthn25pct52wkh',
                    '52wkL', '52wkH','RsratingOver70']

        value_request_body = [

            [

            ]
        ]
        value_request_body.append(s_header)
        row_counter = 3
        for easy in range(0 , len(temp_val)):
            ticker_column = "A" + str(row_counter)
            ma150d_column = "C" + str(row_counter)
            ma250d_column = "D" + str(row_counter)
            l_counter = "L" + str(row_counter)
            m_counter = "M" + str(row_counter)
            data = [temp_val[easy], '=GOOGLEFINANCE(' + ticker_column +
                    ',"price")', '=average(query(GOOGLEFINANCE('
                    + ticker_column + ',"price",WORKDAY(TODAY(),-150),TODAY()),'
                             '"Select Col2 order by Col1 DESC"))''',
                    '=average(query(GOOGLEFINANCE(' + ticker_column +
                    ',"price",WORKDAY(TODAY(),-200),TODAY()),'
                    '"Select Col2 order by Col1 DESC"))',
                    '=if(GOOGLEFINANCE(' + ticker_column + ', "price")>'
                    + ma150d_column + ', "TRUE")',
                    '=if( ' + ma150d_column + '> ' + ma250d_column +
                    ',"TRUE")', '=if(average(query(GOOGLEFINANCE('
                            + ticker_column + ',"price",WORKDAY(TODAY(),-200),'
                            'TODAY()-30),"Select Col2 order by Col1 DESC"))< '
                            + ma250d_column + ',"TRUE")', '=if(average(query'
                            '(GOOGLEFINANCE(' + ticker_column + ',"price",'
                            'WORKDAY(TODAY(),-50),TODAY()),'
                            '"Select Col2 order by Col1 DESC"))>'
                            + ma150d_column +',"TRUE")',
                    '=if(average(query(GOOGLEFINANCE(' + ticker_column +
                    ',"price",WORKDAY(TODAY(),-50),TODAY()),'
                    '"Select Col2 order by Col1 DESC"))<GOOGLEFINANCE('
                    + ticker_column + ',"price"),"TRUE")',
                    '=IF(GOOGLEFINANCE(' + ticker_column + ',"price")>('
                    + l_counter + ' *0.3),"TRUE")', '=IF(GOOGLEFINANCE('
                    + ticker_column + ',"price")>' + m_counter + '-(' +
                    m_counter + ' * 0.25),"TRUE")', '=GOOGLEFINANCE(' +
                    ticker_column + ',"low52")', '=GOOGLEFINANCE(' +
                    ticker_column + ',"high52")']
            value_request_body.append(data)
            row_counter += 1
        range_notation = k
        tempvar = {
            "majorDimension": "ROWS",
            "values": value_request_body
        }
        sheet = spreadsheet_service.spreadsheets()
        sheet.values().append(
            spreadsheetId=sh_id,
            range=range_notation,
            valueInputOption='USER_ENTERED',
            body=tempvar).execute()



main()
