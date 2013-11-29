#By: Eran Ben Horin, eran at valuation.co.il

from pandas import Series, DataFrame, ols
import pandas.io.data as web
import pandas as pd
import numpy as np
import datetime
from dateutil.relativedelta import relativedelta
import statsmodels.api as sm
import csv
import mechanize
from lxml import etree
from lxml.html.clean import clean_html
import lxml
import cookielib

#Required functions for parsing, etc.
#http://ariffwambeck.co.uk/2010/11/20/html-table-parser-in-python/
def parse(filec, missingCell='NaN'):
    """
    Parses all HTML tables found in the String. Missing data or those
    without text content will be replaced with the missingCell string.

    Returns a list of lists of strings, corresponding to rows within all
    found tables.
    """
    #utf_8_parser = etree.HTMLParser(encoding="utf-8")
    doc = lxml.html.fragment_fromstring(clean_html(filec))
    tableList = doc.xpath("//table")
    dataList = []
    for table in tableList:
        dataList.append(parseTable(table, missingCell))
    return dataList


def parseTable(table, missingCell):
    """
    Parses the individual HTML table, returning a list of its rows.
    """
    rowList = []
    for row in table.xpath('.//tr'):
        colList = []
        cells = row.xpath('.//th') + row.xpath('.//td')
        for cell in cells:
            # The individual cell's content
            content = cell.text_content().encode('utf-8')
            if content == "":
                content = missingCell
            colList.append(content)
        rowList.append(colList)
    return rowList


def convert_list_to_dataframe(dflist):
    """
    A Dataframe is created from Table, which is given as list
    """
    #Ensure there is Header and atleast one row
    if len(dflist) > 1:
        #First row is table header, so we define it with columns
        #Table data is from 2nd row onwards
        return DataFrame(dflist[1:], columns=dflist[0])
        
#Fetch index
def fetch_index_data_as_html(start_date='01/08/2013', end_date='01/10/2013', frequency='Weekly', index_id=137, lang='eng'):
    """
    Fetch indices details between the given periode
    Format of args must be kept proper

    """
    FREQ_DICT = {
        'daily': 'rbFrequency1',
        'Weekly': 'rbFrequency2',
        'Monthly': 'rbFrequency3'
    }

    LANG_DICT = {
        'eng': 'g_b2f63986_2b4a_438d_b1b1_fb08c9e1c862',
        'heb': 'g_54223d45_af2f_49cf_88ed_9e3db1499c51'
    }

    sdate = start_date.split('/')
    edate = end_date.split('/')

    br = mechanize.Browser()
    cj = cookielib.LWPCookieJar()
    br.set_cookiejar(cj)
    # Browser options
    br.set_handle_equiv(True)
    br.set_handle_redirect(True)
    br.set_handle_referer(True)
    br.set_handle_robots(False)

    # Follows refresh 0 but not hangs on refresh > 0
    br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)
    br.addheaders = [('User-agent', """Mozilla/5.0 (Windows NT 6.1; rv:24.0) Gecko/20100101 Firefox/24.0""")]
    #Combine proper link
    link ="http://www.tase.co.il/%s/MarketData/Indices/MarketCap/Pages/IndexHistoryData.aspx?Action=1&addTab=&IndexId=%s" % (lang, index_id)

    try:
        #Opening Indices Data Page
        br.open(link)
    except:
        print "Failed to Open Indices Page"
        return

    br.select_form(nr=0)
    br.set_all_readonly(False)
    try:
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbPeriodOTC'] = ['rbPeriodOTC8']
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbPeriod'] = ['rbPeriod8']
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyFromCalendar$TaseCalendar$dateInput_TextBox'] = '%s/%s/%s' % (sdate[0], sdate[1], sdate[2])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyToCalendar$TaseCalendar$dateInput_TextBox'] = '%s/%s/%s' % (edate[0], edate[1], edate[2])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbFrequency'] = [FREQ_DICT[frequency]]
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyFromCalendar_TaseCalendar'] = '%s-%s-%s' % (sdate[2], sdate[1], sdate[0])
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyFromCalendar_TaseCalendar_calendar_SD'] = '[[%s,%s,%s]]' % (int(sdate[2]), int(sdate[1]), int(sdate[0]))
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyToCalendar_TaseCalendar'] = '%s-%s-%s' % (edate[2], edate[1], edate[0])
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyToCalendar_TaseCalendar_calendar_SD'] = '[[%s,%s,%s]]' % (int(edate[2]), int(edate[1]), int(edate[0]))
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_extraDateAfterCalendar_TaseCalendar'] = '%s-%s-%s' % (edate[2], edate[1], edate[0])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyToCalendar$TaseCalendar$dateInput'] = '%s-%s-%s 0:0:0' % (int(edate[2]), int(edate[1]), int(edate[0]))
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyFromCalendar$TaseCalendar$dateInput'] = '%s-%s-%s 0:0:0' % (int(sdate[2]), int(sdate[1]), int(sdate[0]))
    except:
        print "Failed to Assign components, Check parameters passed with Webpage"
        return
    #First Submission should have Event Taget, for 'Weekly' input click simulation
    br["__EVENTTARGET"] = r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$'+FREQ_DICT[frequency]
    br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$hiddenID'] = "0"

    #Emulate POST the Weekly Button click
    try:
        br.submit()
    except:
        print "Error Submitting First Weekly Button Click"
        return

    br.select_form(nr=0)
    br.set_all_readonly(False)
    br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$hiddenID'] = "0"

    try:
        #Actual Post Method to get Indices Data
        br.submit()
    except:
        print "Error Submitting POST Request"
        return
    html = br.response().read()
    #Test:Save Downloaded page
    # fileo =open('page_%s.html' %(lang), 'wb')
    # fileo.write(html)
    # fileo.close()

    #Parsing Data
    utf_8_parser = etree.HTMLParser(encoding="utf-8")
    html_tree = etree.HTML(html, parser=utf_8_parser)
    table_etree = html_tree.xpath('//*[@id="ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_gridHistoryData_DataGrid1"]')
    #print table_etree
    try:
        table_html = etree.tostring(table_etree[0])
        #Test:Save Extracted Table
        # fileo =open('table_%s.html' %(lang), 'wb')
        # fileo.write(table_html)
        # fileo.close()
    except:
        table_html = "<table><tr><td>Blank</td></tr></table>"
        print "Failed to fetch the tabledata"
    return table_html

#Fetch stocks data
def fetch_company_data_as_html(start_date='01/09/2013', end_date='01/10/2013', frequency='Weekly', company_id='000230', share_id='00230011', lang='eng'):
    """
    Fetch share value details for a company between the given periode
    Format of args must be kept proper

    """
    FREQ_DICT = {
        'daily': 'rbFrequency1',
        'Weekly': 'rbFrequency2',
        'Monthly': 'rbFrequency3'
    }

    LANG_DICT = {
        'eng': 'g_301c6a3d_c058_41d6_8169_6d26c5d97050',
        'heb': 'g_c001c0d9_0cb8_4b0f_b75a_7cc3b6f7d790'
    }

    company_id = str(company_id)
    share_id = str(share_id)

    sdate = start_date.split('/')
    edate = end_date.split('/')

    br = mechanize.Browser()
    cj = cookielib.LWPCookieJar()
    br.set_cookiejar(cj)
    # Browser options
    br.set_handle_equiv(True)
    br.set_handle_redirect(True)
    br.set_handle_referer(True)
    br.set_handle_robots(False)

    # Follows refresh 0 but not hangs on refresh > 0
    br.set_handle_refresh(mechanize._http.HTTPRefreshProcessor(), max_time=1)
    br.addheaders = [('User-agent', """Mozilla/5.0 (Windows NT 6.1; rv:24.0) Gecko/20100101 Firefox/24.0""")]
    #Combine proper link
    link = "http://www.tase.co.il/"+lang+"/general/company/Pages/companyHistoryData.aspx?companyID="+company_id+"&subDataType=0&shareID="+share_id

    try:
        #Opening Company Data Page
        br.open(link)
    except:
        print "Failed to Open Company Page"
        return

    br.select_form(nr=0)
    br.set_all_readonly(False)
    try:
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbPeriodOTC'] = ['rbPeriodOTC8']
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbPeriod'] = ['rbPeriod8']
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyFromCalendar$TaseCalendar$dateInput_TextBox'] = '%s/%s/%s' % (sdate[0], sdate[1], sdate[2])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyToCalendar$TaseCalendar$dateInput_TextBox'] = '%s/%s/%s' % (edate[0], edate[1], edate[2])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$rbFrequency'] = [FREQ_DICT[frequency]]
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyFromCalendar_TaseCalendar'] = '%s-%s-%s' % (sdate[2], sdate[1], sdate[0])
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyFromCalendar_TaseCalendar_calendar_SD'] = '[[%s,%s,%s]]' % (int(sdate[2]), int(sdate[1]), int(sdate[0]))
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyToCalendar_TaseCalendar'] = '%s-%s-%s' % (edate[2], edate[1], edate[0])
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_dailyToCalendar_TaseCalendar_calendar_SD'] = '[[%s,%s,%s]]' % (int(edate[2]), int(edate[1]), int(edate[0]))
        br[r'ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_extraDateAfterCalendar_TaseCalendar'] = '%s-%s-%s' % (edate[2], edate[1], edate[0])
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyToCalendar$TaseCalendar$dateInput'] = '%s-%s-%s 0:0:0' % (int(edate[2]), int(edate[1]), int(edate[0]))
        br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$dailyFromCalendar$TaseCalendar$dateInput'] = '%s-%s-%s 0:0:0' % (int(sdate[2]), int(sdate[1]), int(sdate[0]))
    except:
        print "Failed to Assign components, Check parameters passed with Webpage"
        return
    #First Submission should have Event Taget, for 'Weekly' input click simulation
    br["__EVENTTARGET"] = r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$'+FREQ_DICT[frequency]
    br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$hiddenID'] = "0"

    #Emulate POST the Weekly Button click
    try:
        br.submit()
    except:
        print "Error Submitting First Weekly Button Click"
        return

    br.select_form(nr=0)
    br.set_all_readonly(False)
    br[r'ctl00$SPWebPartManager1$'+LANG_DICT[lang]+'$ctl00$HistoryData1$hiddenID'] = "0"

    try:
        #Actual Post Method to get Company Share Data
        br.submit()
    except:
        print "Error Submitting POST Request"
        return
    html = br.response().read()
    #Test:Save Downloaded page
    # fileo =open('page_%s.html' %(lang), 'wb')
    # fileo.write(html)
    # fileo.close()

    #Parsing Data
    utf_8_parser = etree.HTMLParser(encoding="utf-8")
    html_tree = etree.HTML(html, parser=utf_8_parser)
    table_etree = html_tree.xpath('//*[@id="ctl00_SPWebPartManager1_'+LANG_DICT[lang]+'_ctl00_HistoryData1_gridHistoryData_DataGrid1"]')
    #print table_etree
    try:
        table_html = etree.tostring(table_etree[0])
        #Test:Save Extracted Table
        # fileo =open('table_%s.html' %(lang), 'wb')
        # fileo.write(table_html)
        # fileo.close()
    except:
        table_html = "<table><tr><td>Blank</td></tr></table>"
        print "Failed to fetch the tabledata"
    return table_html

#Fetch an index prices and return a DataFrame
def get_index_price(index_id, start_date, end_date, frequency):
    
    M_DICT = {
        'TA100': '137',
        'TA25': '142',
        'SP500': '%5EGSPC'
    }
    
    index_data = fetch_index_data_as_html(index_id=M_DICT[index_id], start_date=start_date, end_date=end_date, frequency=frequency)
    
    df = None
    #Check company data is blank or not
    if index_data:
        #the html code from website is dirty, so we are cleaning it with parse function, convert it as list of tables
        #parse function returns array of multiple tables,
        tables = parse(index_data)
        #Since we will get only one table from the website, we only use table[0]
        for table in tables:
            df = convert_list_to_dataframe(table)
        
        #Set Date as index and sort the DataFrame in ascending order
        df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%y')
        df = df.set_index('Date')
        df = df.sort()
        for index, row in df.iterrows():
            row['Closing Index Value'] = float(row['Closing Index Value'].replace(',',''))

    df[index_id] = df['Closing Index Value']
    return DataFrame(df[index_id])


def get_stocks(tickers, market, start_date, end_date, frequency):

    #Set Frequency for resampling
    FREQ_DICT = {
        'Weekly': 'W-FRI',
        'Monthly': 'M',
    }
    
    start_yahoo = datetime.datetime.strptime(start_date, '%d/%m/%Y')
    end_yahoo = datetime.datetime.strptime(end_date, '%d/%m/%Y')
    
    #Set market portfolio
    if (market != 'TA100') and (market != 'TA25'):
        if (market == 'SP500'):
            prices = DataFrame(web.get_data_yahoo('VFINX', start_yahoo, end_yahoo)['Adj Close'].resample(FREQ_DICT[frequency], how='last', fill_method='ffill'), columns=['SP500'])
        else:
            prices = DataFrame(web.get_data_yahoo(market, start_yahoo, end_yahoo)['Adj Close'].resample(FREQ_DICT[frequency], how='last', fill_method='ffill'), columns=[market])
    else:
        prices = get_index_price(index_id = market, start_date = start_date, end_date = end_date, frequency = frequency).resample(FREQ_DICT[frequency], how = 'last')
    
    #Set Stocks Prices
    i = 0
    while (i < len(tickers)):
        get_df_ticker = DataFrame(web.get_data_yahoo(tickers[i], start_yahoo, end_yahoo)['Adj Close'].resample(FREQ_DICT[frequency], how = 'last'), columns=[tickers[i]])
        prices = pd.concat([prices, get_df_ticker], join='outer', axis = 1)
      
        i += 1
    
    changes = prices.pct_change()
        
    return prices, changes[1:]


def get_betas_table(list, m, f, start_date, end_date):

    prices, returns = get_stocks(list, m, start_date, end_date, f)
    cols = prices.columns.tolist()
    cols.append(cols.pop(0))
    prices, returns = prices[cols], returns[cols]
    
    betas_columns = ['Beta', 'Adj Beta', 'Std Err', 't-stat', 'p-value', 'Adj R^2', 'N']
    betas = DataFrame(columns = betas_columns, index = list)
    betas.index.name = 'Ticker'
    
    #Params are based on dir(model)
    for i in returns.columns[:-1] :
        mask = pd.notnull(returns[i])
        betas.ix[i]['Beta'] = pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).beta['x']
        betas.ix[i]['Adj Beta'] = format(0.67 * betas.ix[i]['Beta'] + 0.33, '.3f')
        betas.ix[i]['Beta'] = format(betas.ix[i]['Beta'], '.3f')
        betas.ix[i]['Std Err'] = format(pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).std_err['x'], '.3f')
        betas.ix[i]['t-stat'] = format(pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).t_stat['x'], '.3f')
        betas.ix[i]['p-value'] = format(pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).p_value['x'], '.3f')
        betas.ix[i]['Adj R^2'] = format(pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).r2_adj, '.3f')
        betas.ix[i]['N'] = pd.ols(y=returns[i][mask], x=returns[m][mask], intercept=True).nobs
    
    #Create CSV
    prices.to_csv('C:\Users\eran\Desktop\My Dropbox\Valuation.co.il\Python Tools\Beta Tool\\prices.csv')
    returns.to_csv('C:\Users\eran\Desktop\My Dropbox\Valuation.co.il\Python Tools\Beta Tool\\returns.csv')
    
    #Return Betas table
    return betas


#Get input...
list = ['F', 'AAPL', 'GOOG']
m = 'TA100'
start_date = '1/8/2011'
end_date = '31/10/2013'
f = 'Monthly'

#Get the table
betas = get_betas_table(list, m, f, start_date, end_date)

#Print the table and save it to CSV
print betas
betas.to_csv('C:\Users\eran\Desktop\My Dropbox\Valuation.co.il\Python Tools\Beta Tool\\betas.csv')