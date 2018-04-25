# 
# convert Interactive Broker Active statement CSV file to Power Portfolio transactions format
#
# Royer Wang
# Mar. 2018

import os
import sys
import re
import csv
import decimal

import logging

from datetime import datetime

from commfunc import detect_file_encoding
from commvar import TRANS_HEADER




TRANS_ROOTDIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_FILENAME = os.path.basename(os.path.abspath(__file__))

# The log file path
LOGFILE = os.path.join(TRANS_ROOTDIR,"log", re.sub(r"\.py",".log", SCRIPT_FILENAME, flags=re.I))


# The Out Put Trans file
#OUTFILE_BASENAME = "ib_trans.csv"

# The output file full path
#OUTFILE_PATH = os.path.join(TRANS_ROOTDIR, "outtrans",OUTFILE_BASENAME)
OUTFILE_PATH = None




# The default symbol map file of IB
SYMBOL_MAP_FILE = os.path.join(TRANS_ROOTDIR,"symbolmap_ib.py")

# load symbol map
from symbolmap_ib import SYMBOLMAP



def Statement_line(row, statement):
    if row[1] != "Data":
        return
    statement[row[2]] = row[3]
    if row[2] == "Period":
        p = [d.strip() for d in row[3].split('-')]
        statement['Period_begin'] = datetime.strptime(p[0], "%B %d, %Y")
        statement['Period_end'] = datetime.strptime(p[1], "%B %d, %Y")


def AccountInfo_line(row, statement):
    if not row[0] in statement:
        statement[row[0]] = {}
    if row[1] == "Data" and row[2]=="Account": # only need Account ID infomation
        statement[row[0]][row[2]] = row[3]
    if row[1] == "Data" and row[2] == "Base Currency": # we need Base Currency infomation
        statement[row[0]][row[2]] = row[3]


def Trades_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns':{}, "data":[]}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]]=col
            if row[col] == "Notional Value" and 'Proceeds' not in statement[row[0]]['columns'] :
                statement[row[0]]['columns']['Proceeds'] = col
            col += 1
        statement[row[0]]['columns']['srcline'] = col
    if row[1] == "Data":
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Quantity']]
        t[statement[row[0]]['columns']['Quantity']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['T. Price']]
        t[statement[row[0]]['columns']['T. Price']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Proceeds']]
        t[statement[row[0]]['columns']['Proceeds']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Comm/Fee']]
        t[statement[row[0]]['columns']['Comm/Fee']] = locale.atof(s) if len(s) > 0 else 0.0
        
        s = t[statement[row[0]]['columns']['Basis']]
        t[statement[row[0]]['columns']['Basis']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)

def TransactionFee_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns':{},'data':[] }
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
        statement[row[0]]['columns']['srcline'] = col
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Quantity']]
        t[statement[row[0]]['columns']['Quantity']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Trade Price']]
        t[statement[row[0]]['columns']['Trade Price']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)


def Deposit_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data':[]}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)


def Fees_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)


def Dividends_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0
  
        statement[row[0]]['data'].append(t)


def WithholdingTax_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)

def Interest_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Amount']]
        t[statement[row[0]]['columns']['Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)

def DividendAccruals_line(row, statement, linenum):
    if not row[0] in statement:
        statement[row[0]] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement[row[0]]['columns'][row[col]] = col
            col += 1
        statement[row[0]]['columns']['srcline'] = col
    if row[1] == "Data" and row[statement[row[0]]['columns']['Code']] == "Re":
        t = [x for x in row]
        t.append(linenum)
        t.append(linenum)
        s = t[statement[row[0]]['columns']['Quantity']]
        t[statement[row[0]]['columns']['Quantity']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Tax']]
        t[statement[row[0]]['columns']['Tax']] = locale.atof(s) if len(s) > 0 else 0.0

        s = t[statement[row[0]]['columns']['Gross Amount']]
        t[statement[row[0]]['columns']['Gross Amount']] = locale.atof(s) if len(s) > 0 else 0.0

        statement[row[0]]['data'].append(t)    

def LentInterest_line(row, statement, linenum):
    if not 'Lent Interest' in statement:
        statement['Lent Interest'] = {'columns': {}, 'data': []}
    if row[1] == "Header":
        col = 0
        while col < len(row):
            statement['Lent Interest']['columns'][row[col]] = col
            col += 1
    if row[1] == "Data" and not row[2].startswith("Total"):
        t = [x for x in row]
        t.append(linenum)
        s = t[statement['Lent Interest']['columns']['Interest Paid to Customer']]
        t[statement['Lent Interest']['columns']['Interest Paid to Customer']] = locale.atof(s) if len(s) > 0 else 0.0

        statement['Lent Interest']['data'].append(t)

def Symbolinfo_line(row, statement):
    if not 'SymbolInfo' in statement:
        statement['SymbolInfo'] = {'columns': {
            'Asset Category':0,
            'Symbol': 1,
            'Description': 2,
            'Conid': 3,
            'Multiplier':4,
            'Expiry':5,
            'Delivery Month':6,
            'Type':7,
            'Strike':8,
        }, 'data': [], 'map': {}}
    if row[1] == "Data":
        t = [''] * 9
        t[0] = row[2]   # Asset Category
        t[1] = row[3]   # Symbol
        t[2] = row[4]   # Description
        t[3] = row[5]   # Conid
        if row[2] == "Stocks":
            t[4] = row[7]   # Multipile for stocks
        elif row[2] == "Equity and Index Options":
            t[4] = row[6]   # Multiplier for options
            t[5] = row[7]   # Expiry for options;
            t[6] = row[8]   # Delivery Month for options
            t[7] = row[9]   # Type for options
            t[8] = row[10]  # strike price for options
        elif row[2] == 'Futures':
            t[4] = row[6]   # Multiplier for Futures
            t[5] = row[7]   # Expiry for Futures
            t[6] = row[8]   # Delivery Month for Futures
        else:
            logging.warning("Unkown Asset Category: {0} in {1}".format(row[2],row[0]))
        statement['SymbolInfo']['data'].append(t)
        statement['SymbolInfo']['map'][t[1]] = t



import argparse

argParser = argparse.ArgumentParser(description="Convert IB .csv format statement file to PowerPortfolio transactions file")
argParser.add_argument("srcfile", help="The IB statement csv format file")
argParser.add_argument("desfile", help="The output transaction file path", nargs='?')
argParser.add_argument("-A", "--account", dest="accountname", help="The account name in PowerPortfolio", default="IB")
argParser.add_argument("-B", "--brokeraccid", dest="acid_in_broker", 
    help="The accountid in Broker. for check is correct transaction file.")
args = argParser.parse_args()

ACCOUNTNAME = args.accountname
ACCOUNTID_IN_BROKER = args.acid_in_broker



import locale
if os.name == 'nt':
    locale.setlocale(locale.LC_ALL,locale="USA")
else:
    locale.setlocale(locale.LC_ALL, 'en-US.UTF-8')

if args.desfile:
    OUTFILE_PATH = args.desfile
# make sure log file directory is exist
if not os.path.exists(os.path.dirname(LOGFILE)):
    os.makedirs(os.path.dirname(LOGFILE))

logging.basicConfig(filename=LOGFILE, level=logging.DEBUG, filemode='w',\
    format='%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S' )
consolehandler = logging.StreamHandler()
consolehandler.setLevel(logging.WARNING)
consolehandler.setFormatter(fmt=logging.Formatter("%(levelname)s %(message)s"))
logging.getLogger().addHandler(consolehandler)


# make sure output file directory is exist
if OUTFILE_PATH:
    try:
        if not os.path.exists(os.path.dirname(OUTFILE_PATH)):
            os.makedirs(os.path.dirname(OUTFILE_PATH))
            logging.info("created output directory: " + os.path.dirname(OUTFILE_PATH))
    except OSError as err:
        str = "Creating output directory Error: " + err.strerror
        logging.error(str)
        exit(1)


logging.info("Start convert {0} to {1}".format(args.srcfile, OUTFILE_PATH))




LAST_TRADE_DATE = None
writeheader = True
outfile_encoding = 'utf-8'
if OUTFILE_PATH and os.path.exists(OUTFILE_PATH):
    outfile_encoding = detect_file_encoding(OUTFILE_PATH)
    with open(OUTFILE_PATH, mode='r', newline='', encoding=outfile_encoding) as transfile:
        csvout = csv.DictReader(transfile, delimiter=',')
        writeheader = False
        for row in csvout:
            if not LAST_TRADE_DATE or (LAST_TRADE_DATE and row['Date'] > LAST_TRADE_DATE):
                if row[TRANS_HEADER[0]] == ACCOUNTNAME:
                    LAST_TRADE_DATE = row['Date']
    logging.info("The last trade date in {0} is {1}".format(os.path.basename(OUTFILE_PATH), LAST_TRADE_DATE))

if args.srcfile:
    src_encoding = detect_file_encoding(args.srcfile)
else:
    src_encoding = 'utf-8'

outdata = []
srclines = 0
Statement = {}
with open(args.srcfile, mode='r', newline='', encoding=src_encoding) as srcfile:
    csvout = csv.reader(srcfile)
    segment = None
    for row in csvout:
        srclines += 1
        try:
            if row[0] != segment:
                segment = row[0]
            if segment == "Statement" :
                Statement_line(row, Statement)
            elif segment == "Account Information":
                AccountInfo_line(row, Statement)
            elif segment == "Trades":
                Trades_line(row, Statement, srclines)
            elif segment == "Transaction Fees":
                TransactionFee_line(row, Statement, srclines)
            elif segment == "Deposits & Withdrawals":
                Deposit_line(row, Statement, srclines)
            elif segment == "Fees":
                Fees_line(row, Statement, srclines)
            elif segment == "Dividends":
                Dividends_line(row, Statement, srclines)
            elif segment == "Withholding Tax":
                WithholdingTax_line(row, Statement, srclines)
            elif segment == "Interest":
                Interest_line(row, Statement, srclines)
            elif segment == "Change in Dividend Accruals":
                DividendAccruals_line(row, Statement, srclines)
            elif 'IB Managed Securities Lent Interest Details' in segment:
                LentInterest_line(row, Statement, srclines)
            elif segment == "Financial Instrument Information":
                Symbolinfo_line(row, Statement)
        except Exception as e:
            strerr = "file: {file} line: {line} has error: {error}".format(file=os.path.basename(args.srcfile), line=srclines, error=str(e))
            logging.error(strerr)
            exit(1)


# if has LAST_TRADE_DATE, then check total statement period
if LAST_TRADE_DATE and LAST_TRADE_DATE > Statement['Period_end'].strftime('%Y-%m-%d'):
    logging.warning('This statement period: {0} - {1} is already in {2}'.format(Statement['Period_begin'], Statement['Period_end'], OUTFILE_PATH))
    exit(1)

# start combine Transactions Fees to match trade
if 'Transaction Fees' in Statement:
    tfcols = Statement['Transaction Fees']['columns']
    tdcols = Statement['Trades']['columns']
    cbcombined = 0
    for tf in Statement['Transaction Fees']['data']:
        tf.append(0)
        tf_date = tf[tfcols['Date/Time']].split(',')[0].strip()
        tf_time = tf[tfcols['Date/Time']].split(',')[1].strip()
        for trade in Statement['Trades']['data']:
            td_date = trade[tdcols['Date/Time']].split(',')[0].strip()
            td_time = trade[tdcols['Date/Time']].split(',')[1].strip()
            if tf[tfcols['Currency']] == trade[tdcols['Currency']] and \
                tf[tfcols['Symbol']] == trade[tdcols['Symbol']] and \
                (tf[tfcols['Date/Time']] == trade[tdcols['Date/Time']] or \
                (tf_date == td_date and tf_time >= td_time and \
                tf[tfcols['Trade Price']]== trade[tdcols['T. Price']] and \
                'P' in trade[tdcols['Code']].split(';'))):
                tf[-1] = 1
                cbcombined += 1
                trade[tdcols['Comm/Fee']] += tf[tfcols['Amount']]
                break
        
        if tf[-1] == 0:
            logging.warning('Transaction Fees(line: {0}) symbol: {1} can not match to Trade record.'.format(tf[-2], tf[tfcols['Symbol']]))

    total_tf = len(Statement['Transaction Fees']['data'])
    left_tf = total_tf - cbcombined
    if left_tf == 0:
        logging.info('Transaction Fees total: {0}. successed matched to trade record: {1}'.format(total_tf, cbcombined))
    else:
        logging.warning("Not all Trasaction Fees record match to trade record. total: {0}, left: {1}".format(total_tf, left_tf))


if ACCOUNTID_IN_BROKER and ACCOUNTID_IN_BROKER != Statement['Account Information']['Account']:
    strwarning = "IB Account {0} is not you asked account: {1}".format(ACCOUNTID_IN_BROKER, ACCOUNTNAME)
    logging.warning(strwarning)

pp_trans = []

convert_date = datetime.now().strftime("%Y-%m-%d")
cbconverted = 0

symbolnotmap = []
# Process Trades Segment
if 'Trades' in Statement:
    for ib_trade in Statement['Trades']['data']:
        tdate = ib_trade[Statement['Trades']['columns']['Date/Time']][0:10]
        if LAST_TRADE_DATE and LAST_TRADE_DATE >= tdate:
            logging.warning('line: {0}: trade date already in {1}. skip this line.'.format(ib_trade[-1], OUTFILE_PATH))
            continue
        trans = []
        trans.append(ACCOUNTNAME)   #AccountName
        trans.append(tdate)   #Date

        AssetCategory = ib_trade[Statement['Trades']['columns']['Asset Category']]
        Currency = ib_trade[Statement['Trades']['columns']['Currency']]
        Symbol = ib_trade[Statement['Trades']['columns']['Symbol']]
        Quantity = ib_trade[Statement['Trades']['columns']['Quantity']]
        Price = ib_trade[Statement['Trades']['columns']['T. Price']]
        Fee = ib_trade[Statement['Trades']['columns']['Comm/Fee']]
        Proceeds = ib_trade[Statement['Trades']['columns']['Proceeds']]
        Code = ib_trade[Statement['Trades']['columns']['Code']]
        srcline = ib_trade[Statement['Trades']['columns']['srcline']]

        # process stocks, options and Futures trade record
        if AssetCategory in ('Stocks', 'Equity and Index Options','Futures'):

            #TransType 
            if Code[0] == 'O':
                if Quantity > 0:
                    trans.append('Buy')
                elif Quantity < 0:
                    trans.append('SelltoOpen')
                else:
                    strerr = "line: {srcline} Open Position trade but Quantity is 0. CAN NOT analyze. Exit!".format(srcline=srclines)
                    logging.error(strerr)
                    exit(1)
            elif Code[0] == 'C':
                if Quantity > 0:
                    trans.append('BuytoClose')
                elif Quantity < 0:
                    trans.append('Sell')
                else:
                    strerr = "line: {srcline} Close Position trade but Quantity is 0. CAN NOT analyze. Exit!".format(srcline=srclines)
                    logging.error(strerr)
                    exit(1)
            else:
                strerr = "line: {srcline} Code: '{code}' is unsupport. exit!".format(srcline=srcline, code = Code)
                logging.error(strerr)
                exit(1)
            
            # SubTransType
            scode = Code.split(';')
            subtype = ''
            if 'Ep' in scode:
                subtype = 'Expired'
            trans.append(subtype)

            # Symbol
            ppsymbol = Symbol
            if Symbol in SYMBOLMAP:
                ppsymbol = SYMBOLMAP[Symbol]
            else:
                if not Symbol in symbolnotmap:
                    logging.warning("Symbol: '{0}' is not in symbolmap.py, use IB Symbol as PowerPortfolio Symbol.".format(Symbol))
                    symbolnotmap.append(Symbol)
            trans.append(ppsymbol)

            # Qty 
            #TODO: split DO NOT abs(Quantity). I don't know IB how to record split info in statement report.
            trans.append(abs(Quantity))

            # Price
            trans.append(Price)

            # Fee
            trans.append(abs(Fee))

            # AccruedInterest
            trans.append('')

            # Ammount
            if trans[2] in ('Buy', 'BuytoClose'):
                # Qty * Price + Fee
                trans.append(abs(Proceeds) + trans[7])
            else:
                # Qty * Price - Fee
                trans.append(abs(Proceeds) - trans[7])

            # Comment
            trans.append('')

            # OrderID
            trans.append('')

            #ConvertDate
            trans.append(convert_date)

            pp_trans.append(trans)
            cbconverted += 1

        elif AssetCategory == 'Forex':
            trans_pair = []
            trans_pair.append(ACCOUNTNAME)
            trans_pair.append(ib_trade[Statement['Trades']['columns']['Date/Time']][0:10])
            # TransType
            if Proceeds > 0:
                trans.append("BuyCurrency")
                trans_pair.append("BuyCurrency_Pair")
            else:
                trans.append("SellCurrency")
                trans_pair.append("SellCurrency_Pair")
            
            # SubTransType
            trans.append('')
            trans_pair.append('')

            # Symbol
            ppsymbol = Currency
            if ppsymbol == "CNH":
                ppsymbol = "CNY"    # CNY in IB symbol is CNH
            ppsymbol = '*' + ppsymbol
            trans.append(ppsymbol)
            ss = Symbol.split('.')
            pairsymbol = ''
            if ss[0] == Currency:
                pairsymbol = ss[1]
            else:
                pairsymbol = ss[0]
            if pairsymbol == "CNH":
                pairsymbol = "CNY"
            pairsymbol = '*' + pairsymbol
            trans_pair.append(pairsymbol)

            # Qty
            trans.append(abs(Quantity))
            trans_pair.append(abs(Quantity))

            # Price
            trans.append(abs(Price))
            trans_pair.append(1)

            # Fee
            trans.append('')
            trans_pair.append(abs(Fee))

            # AccruedInterest
            trans.append('')
            trans_pair.append('')

            # Amount
            trans.append(abs(Proceeds))
            trans_pair.append(abs(Quantity) + abs(Fee))

            # Comment
            trans.append('')
            strcomm = " Use to " + trans[2] + " " + Currency + " Ammount: " + str(abs(Proceeds))
            trans_pair.append(strcomm)

            # OrderID
            trans.append('')
            trans_pair.append('')

            # ConvertDate
            trans.append(convert_date)
            trans_pair.append(convert_date)

            pp_trans.append(trans)
            pp_trans.append(trans_pair)
            cbconverted += 1
        else:
            logging.warning('line: {srcline} Asset Category: {ac} is unsupport. skipped.'.format(srcline=srcline, ac=AssetCategory))

    logging.info("{oldnumber} Trades. {newcount} was successful convert to {transcount} transactions. ".format(oldnumber=len(Statement['Trades']['data']), newcount=cbconverted, transcount = len(pp_trans)))
else:
    logging.info("no trades record in file.")
cbTrades = len(pp_trans)


# Process Deposit & Whithdrawal
cbconverted = 0
if 'Deposits & Withdrawals' in Statement:
    for dw in Statement['Deposits & Withdrawals']['data']:
        Currency = dw[Statement['Deposits & Withdrawals']['columns']['Currency']]
        tradedate = dw[Statement['Deposits & Withdrawals']['columns']['Settle Date']]
        Description = dw[Statement['Deposits & Withdrawals']['columns']['Description']]
        amount = dw[Statement['Deposits & Withdrawals']['columns']['Amount']]

        if LAST_TRADE_DATE and LAST_TRADE_DATE >= tradedate:
            logging.warning('line: {0} trade date is already in {1}. skip this line.'.format(dw[-1], OUTFILE_PATH))
            continue

        trans = []

        # AccountName
        trans.append(ACCOUNTNAME)
        
        # Date
        trans.append(tradedate)

        # TransType
        if amount > 0:
            trans.append('Deposit')
        else:
            trans.append('Withdraw')
        
        # SubTransType
        trans.append('')

        # Symbol
        ppsymbol = Currency
        if Currency == "CNH":
            ppsymbol = "CNY"
        ppsymbol = '*' + ppsymbol
        trans.append(ppsymbol)

        # Qty
        trans.append('')

        # Price
        trans.append('')

        # Fee
        trans.append('')

        # AccruedInterest
        trans.append('')

        # Amount
        trans.append(abs(amount))

        # Comment
        trans.append(Description)

        # OrderID
        trans.append('')

        # ConvertDate
        trans.append(convert_date)

        pp_trans.append(trans)

        cbconverted += 1

    logging.info('convet {0} Deposit & Whithdrawal successful'.format(cbconverted))


# Fees

cbconverted = 0
if 'Fees' in Statement:
    for fee in Statement['Fees']['data']:
        trans = []
        Currency = fee[Statement['Fees']['columns']['Currency']]
        tradedate = fee[Statement['Fees']['columns']['Date']]
        Description = fee[Statement['Fees']['columns']['Description']]
        amount = fee[Statement['Fees']['columns']['Amount']]
        if LAST_TRADE_DATE and LAST_TRADE_DATE >= tradedate:
            logging.warning('line: {0} trade date already in {1}. skip this line.'.format(fee[-1], OUTFILE_PATH))
            continue
        # AccountName
        trans.append(ACCOUNTNAME)

        # Date
        trans.append(tradedate)

        # TransType
        trans.append('OtherFee')

        # SubTransType
        trans.append('')

        # Symbol
        ppsymbol = Currency if Currency != "CNH" else "CNY"
        ppsymbol = '*' + ppsymbol
        trans.append(ppsymbol)

        # Qty
        trans.append('')

        # Price
        trans.append('')

        # Fee
        trans.append('')

        # AccruedInterest
        trans.append('')

        # Amount
        trans.append( -1 * amount)

        # Comment
        trans.append(Description)

        # OrderID
        trans.append('')

        # ConvertDate
        trans.append(convert_date)

        cbconverted += 1
        pp_trans.append(trans)
    logging.info("{cb} Fees convert to Transacton".format(cb=cbconverted))


# Dividends 
cbconverted = 0
if 'Dividends' in Statement:
    for divd in Statement['Dividends']['data']:

        trans = []

        Currency = divd[Statement['Dividends']['columns']['Currency']]
        paydate = divd[Statement['Dividends']['columns']['Date']]
        Description = divd[Statement['Dividends']['columns']['Description']]
        Symbol = Description[0:Description.find('(')].strip()
        amount = divd[Statement['Dividends']['columns']['Amount']]
        
        if LAST_TRADE_DATE and LAST_TRADE_DATE >= paydate:
            logging.warning('line: {0} trade date already in {1}. skip this line.'.format(divd[-1], OUTFILE_PATH))
            continue

        # try get quantity from 'Change in Dividend Accruals' segment
        Quantity = ''
        if 'Change in Dividend Accruals' in Statement:
            for cdac in Statement['Change in Dividend Accruals']['data']:
                c = cdac[Statement['Change in Dividend Accruals']['columns']['Currency']]
                pd = cdac[Statement['Change in Dividend Accruals']['columns']['Pay Date']]
                s = cdac[Statement['Change in Dividend Accruals']['columns']['Symbol']]
                a = -1 * cdac[Statement['Change in Dividend Accruals']['columns']['Gross Amount']]
                if a == amount and c == Currency and pd == paydate and s == Symbol:
                    Quantity = cdac[Statement['Change in Dividend Accruals']['columns']['Quantity']]
                    break
        
        # AccountName
        trans.append(ACCOUNTNAME)

        # Date
        trans.append(paydate)

        # TransType
        trans.append('Dividend')

        # SubTransType
        trans.append('')

        # Symbol
        ppsymbol = Symbol
        if Symbol in SYMBOLMAP:
            ppsymbol = SYMBOLMAP[Symbol]
        else:
            if not Symbol in symbolnotmap:
                symbolnotmap.append(Symbol)
                logging.warning("Symbol: '{0}' is not in symbolmap.py, use IB Symbol as PowerPortfolio Symbol.".format(Symbol))
        trans.append(ppsymbol)

        # Qty
        trans.append(Quantity)

        # Price
        trans.append('')

        # Fee
        trans.append('')

        # AccruedInterest
        trans.append('')

        # Amount
        trans.append(amount)

        # Comment
        trans.append(Description)

        # OrderID
        trans.append('')

        # ConvertDate
        trans.append(convert_date)

        cbconverted += 1
        pp_trans.append(trans)
    
    logging.info("Convert {0} dividend records".format(cbconverted))


# Withholding Tax
cbconverted = 0
if 'Withholding Tax' in Statement:
    for wht in Statement['Withholding Tax']['data']:

        trans = []

        Currency = wht[Statement['Withholding Tax']['columns']['Currency']]
        paydate = wht[Statement['Withholding Tax']['columns']['Date']]
        Description = wht[Statement['Withholding Tax']['columns']['Description']]
        Symbol = Description[0:Description.find('(')].strip()
        amount = wht[Statement['Withholding Tax']['columns']['Amount']]

        if LAST_TRADE_DATE and LAST_TRADE_DATE >= paydate:
            logging.warning('line: {0} trade date already in {1}. skip this line.'.format(wht[-1], OUTFILE_PATH))
            continue

        # AccountName
        trans.append(ACCOUNTNAME)

        # Date
        trans.append(paydate)

        # TransType
        trans.append('WithHoldTax')

        # SubTransType
        trans.append('')

        # Symbol
        ppsymbol = Symbol
        if Symbol in SYMBOLMAP:
            ppsymbol = SYMBOLMAP[Symbol]
        else:
            if not Symbol in symbolnotmap:
                symbolnotmap.append(Symbol)
                logging.warning("Symbol: '{0}' is not in symbolmap.py, use IB Symbol as PowerPortfolio Symbol.".format(Symbol))
        trans.append(ppsymbol)

        # Qty
        trans.append('')

        # Price
        trans.append('')

        # Fee
        trans.append('')

        # AccruedInterest
        trans.append('')

        # Amount
        trans.append(-1 * amount)

        # Comment
        trans.append(Description)

        # OrderID
        trans.append('')

        # ConvertDate
        trans.append(convert_date)

        cbconverted += 1
        pp_trans.append(trans)
    
    logging.info("Convert {0} Withholding Tax records".format(cbconverted))



# Interest
cbconverted = 0

if 'Interest' in Statement:
    for interest in Statement['Interest']['data']:

        trans = []

        Currency = interest[Statement['Interest']['columns']['Currency']]
        paydate = interest[Statement['Interest']['columns']['Date']]
        Description = interest[Statement['Interest']['columns']['Description']]
        amount = interest[Statement['Interest']['columns']['Amount']]
        if LAST_TRADE_DATE and LAST_TRADE_DATE >= paydate:
            logging.warning('line: {0} trade date already in {1}. skip this line.'.format(interest[-1], OUTFILE_PATH))
            continue
        # AccountName
        trans.append(ACCOUNTNAME)

        # Date
        trans.append(paydate)

        # TransType
        if amount > 0:
            trans.append('Interest')
        else:
            trans.append('IntPaid')
        
        # SubTransType
        trans.append('')

        # Symbol
        ppsymbol = Currency if Currency != "CNH" else "CNY"
        ppsymbol = "*" + ppsymbol
        trans.append(ppsymbol)

        # Qty
        trans.append('')

        # Price
        trans.append('')

        # Fee
        trans.append('')

        # AccruedInterest
        trans.append('')

        # Amount
        trans.append(abs(amount))

        # Comment
        trans.append(Description)

        # OrderID
        trans.append('')

        # ConvertDate
        trans.append(convert_date)

        cbconverted += 1
        pp_trans.append(trans)
    
    logging.info(" Convert {0} interest record.".format(cbconverted))


# sort all transaction by Date
from operator import itemgetter
pp_trans_sorted = sorted(pp_trans, key=itemgetter(1))

import io

if OUTFILE_PATH:
    # first backup
    desfile = OUTFILE_PATH + ".bak"
    from shutil import copyfile
    copyfile(OUTFILE_PATH, desfile)
    f = open(OUTFILE_PATH, mode='a', newline='', encoding=outfile_encoding)
else:
    f = io.StringIO(newline='')


csvwriter = csv.writer(f)
if writeheader:
    header = []
    header.append(TRANS_HEADER)
    csvwriter.writerows(header)

csvwriter.writerows(pp_trans_sorted)

if not OUTFILE_PATH:
    print(f.getvalue())
    f.close()

f.close()

logging.info('Convert finished. total {0} records.'.format(len(pp_trans_sorted)))


