'''
Convert TD Ameritrade Thinkorswim csv statement file to Power Portfolio transaction
'''

import os
import sys
import re
import csv

import logging

from datetime import datetime
from decimal import Decimal
import decimal

decimal.getcontext().prec = 20
from commvar import TRANS_HEADER
from commfunc import CheckPowerPortfolioFile, detect_file_encoding

TRANS_ROOTDIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_FILENAME = os.path.basename(os.path.abspath(__file__))

# The log file path
LOGFILE = os.path.join(TRANS_ROOTDIR,"log", re.sub(r"\.py",".log", SCRIPT_FILENAME, flags=re.I))


# The all PowerPortfolio transactions file
ALLTRANS_BASENAME = "TDUS.csv"

# The output file full path
ALLTRANS_PATH = os.path.join(TRANS_ROOTDIR, "outtrans",ALLTRANS_BASENAME)

BASE_CURRENCY = 'USD'

import argparse

argParser = argparse.ArgumentParser(description="Convert TD Canada thinkorswim .csv format statement file to PowerPortfolio transactions file")
argParser.add_argument("srcfile", help="The TD thinkorswim statement csv format file")
argParser.add_argument("desfile", help="The output transaction file path", nargs='?')
argParser.add_argument("-A", "--account", dest="accountname", help="The account name in PowerPortfolio", default="TDUSD")
argParser.add_argument("-B", "--brokeraccid", dest="acid_in_broker", 
    help="The accountid in Broker. for check is correct transaction file.")



args = argParser.parse_args()

ACCOUNTNAME = args.accountname
ACCOUNTID_IN_BROKER = args.acid_in_broker

# make sure log file directory is exist
if not os.path.exists(os.path.dirname(LOGFILE)):
    os.makedirs(os.path.dirname(LOGFILE))


# logging.basicConfig(filename=LOGFILE, level=logging.DEBUG, filemode='w',\
#    format='%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S' )

#logging.basicConfig(level=logging.DEBUG)
filehandler = logging.FileHandler(filename=LOGFILE, mode='w', encoding='utf-8')
filehandler.setFormatter(fmt=logging.Formatter('%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S'))
filehandler.setLevel(logging.DEBUG)
logging.getLogger().addHandler(filehandler)

consolehandler = logging.StreamHandler()
consolehandler.setLevel(logging.WARNING)
consolehandler.setFormatter(fmt=logging.Formatter("%(levelname)s %(message)s"))

logging.getLogger().addHandler(consolehandler)
logging.getLogger().setLevel(logging.DEBUG)

logging.info('start convert "{0}" to "{1}"'.format(args.srcfile, args.desfile if args.desfile else "STDOUT"))




LAST_TRADE_DATE = None
old_trans = []

if args.desfile:
    if not os.path.exists(os.path.dirname(os.path.abspath(args.desfile))):
        os.makedirs(os.path.dirname(os.path.abspath(args.desfile)))
        logging.info("Create directory: {0}".format(os.path.dirname(os.path.abspath(args.desfile))))
    else:
        # check desfile is powerPortfolio file and get all transactions record
        if os.path.exists(os.path.abspath(args.desfile)):
            try:
                lines = 0
                with open(args.desfile, mode='r', encoding='utf-8', newline='') as f:
                    csvout = csv.reader(f)
                    for row in csvout:
                        lines += 1
                        if lines == 1:
                            # it should header line
                            if len(TRANS_HEADER) == len(row):
                                for i in range(len(TRANS_HEADER)):
                                    if TRANS_HEADER[i].lower() != row[i].lower():
                                        logging.error('"{0}" header line is incorrect column {1} is not {2}'.format(args.desfile, i, TRANS_HEADER[i]))
                                        exit(1)
                            else:
                                logging.error('{0} header line is incorrect.'.format(args.desfile))
                                exit(1)
                        else:
                            old_trans.append(row)
                            if LAST_TRADE_DATE == None or (LAST_TRADE_DATE < row[1]):
                                LAST_TRADE_DATE = row[1]
                    logging.info('Last Trade Date in "{0}" is {1}'.format(args.desfile, LAST_TRADE_DATE))
            except Exception as e:
                logging.error('Read "{0}" error: {1}. exit convert.'.format(args.desfile, str(e)))
                exit(1)

if os.path.exists(args.srcfile):
    src_encoding = detect_file_encoding(args.srcfile)
    logging.info('"{0}" is encoding by {1}'.format(args.srcfile, src_encoding))
else:
    logging.error('"{0}" does not exist.'.format(args.srcfile))
    exit(1)

# convert TD Ameritrade Date format Month/Day/Year to YYYY-MM-DD format, and write back
with open(args.srcfile, mode='r+', encoding=src_encoding) as f:
    content = f.read()
    part = r'\d{1,2}\/\d{1,2}\/\d{2}'
    m = re.compile(part)
    def convdate(mo):
        d = datetime.strptime(mo.group(0), "%m/%d/%y")
        return d.strftime('%Y-%m-%d')
    if m.search(content) != None:
        content = re.sub(part, convdate, content)
        f.seek(0,0)
        f.write(content)
        f.truncate()
        logging.info('"{0}" convert date format to YYYY-MM-DD successful.'.format(args.srcfile))
    else:
        logging.info('there is no date format convert in "{0}".'.format(args.srcfile))

totalTrans_src = 0
totalTrans_pp = 0

srclines = 0

src_tradhis = []
new_trans = []

strtoday = datetime.now().strftime('%Y-%m-%d')

with open(args.srcfile, mode='r', encoding=src_encoding) as srcfile:
    csvout = csv.reader(srcfile)
    segment = None
    skipsegment = False
    needheader = False
    for row in csvout:
        srclines += 1
        if srclines == 1:
            if len(row) == 1:
                r = re.match(r"^Account Statement for\s(\w+)\s.+since\s(\d{4}-\d{1,2}-\d{1,2}) through (\d{4}-\d{1,2}-\d{1,2})", row[0])
                if r :
                    accountid_in_broker = r.group(1)
                    sdate = r.group(2)
                    edate = r.group(3)
                    if (ACCOUNTID_IN_BROKER and ACCOUNTID_IN_BROKER.lower() != accountid_in_broker.lower()):
                        logging.error('Account ID in broker: "{0}" is not match you provided: "{1}". convert exit.'.format(accountid_in_broker, ACCOUNTID_IN_BROKER))
                        exit(1)
                else:
                    logging.error('"{0}" first line is not match "Account Statment for <accountid> since 2011-1-1 through 2012-1-1". convert exit.'.format(args.srcfile))
                    exit(1)
            else:
                logging.error('"{0}" first line is not match "Account Statment for <accountid> since 2011-1-1 through 2012-1-1". convert exit.'.format(args.srcfile))
                exit(1)
        else:
            if len(row) == 0:
                continue
            elif len(row) == 1:
                segment = row[0]
                if segment in ('Cash Balance', 'Account Trade History'):
                    skipsegment = False
                    needheader = True
                else:
                    skipsegment = True
                    needheader = False
            else:
                if skipsegment:
                    continue
                if needheader:
                    needheader = False
                    continue
                if segment == "Cash Balance":
                    if LAST_TRADE_DATE and LAST_TRADE_DATE >= row[0]:   #DATE field
                        continue
                    if row[2] == 'BAL': # TYPE = BAL, skip
                        continue
                    elif row[2] == 'SFEE': # TYPE = SFEE, convert to otherfee
                        t = [''] * len(TRANS_HEADER)
                        t[0] = ACCOUNTNAME
                        t[1] = row[0]            # Date
                        t[2] = 'OtherFee'        # TransType
                        t[4] = '*' + BASE_CURRENCY  #Symbol
                        t[9] = abs(Decimal(row[5].replace(',','')))            # Amount
                        t[10] = row[4]          # Comment
                        t[11] = row[3]          # REF as OrderID
                        t[12] = strtoday        # Convert Date
                        new_trans.append(t)
                    elif row[2] == 'FND':
                        # normal is deposit/withdraw, but sametimes is transfer symbol
                        if row[4].find('CASH IN') is not -1:
                            # Deposit
                            t = [''] * len(TRANS_HEADER)
                            t[0] = ACCOUNTNAME
                            t[1] = row[0]       # Date
                            t[2] = 'Deposit'    # TransType
                            t[4] = '*' + BASE_CURRENCY  # Symbol
                            t[9] = Decimal(row[7].replace(',',''))  # Amount
                            t[10] = row[4]      # Comment
                            t[11] = row[3]      # OrderID
                            t[12] = strtoday
                            new_trans.append(t)
                        elif re.search(r'^(\d+\.?\d*)\s', row[4]):
                            # fund but description start with number, it may be transfer symbol in
                            r = re.search(r'^(\d+\.?\d*)\s', row[4])
                            t = [''] * len(TRANS_HEADER)
                            t[0] = ACCOUNTNAME
                            t[1] = row[0]       # Date
                            t[2] = 'TransferInLongPos'
                            t[5] = r.group(1)   # Qty
                            t[10] = row[4]      # Comment
                            t[11] = row[3]      # OrderID
                            t[12] = strtoday
                            new_trans.append(t)
                            logging.warning('line: {0} TYPE: FND, maybe is symbol transfer in. Description: "{1}"'.format(srclines, row[4]))
                        elif row[4].startswith('DRIP'):
                            # FND - comment indicate this is Dividend reinvstment(DRIP), change it to Buy Trans
                            # But TD Ameritrade miss Qty infomation
                            r = re.search(r'^(\S+).+\s(\S+)$', row[4])
                            t = [''] * len(TRANS_HEADER)
                            t[0] = ACCOUNTNAME
                            t[1] = row[0]               # Date
                            t[2] = "Buy"                # TransType
                            t[3] = "DRIP"               # SubTransType
                            if r:
                                t[4] = r.group(2)       # Symbol
                            t[9] = abs(Decimal(row[7].replace(',',''))) # Amount
                            t[10] = row[4]
                            t[11] = row[3]
                            t[12] = strtoday
                            new_trans.append(t)
                            logging.warning('line: {0} miss Quantity information for DRIP.'.format(srclines))
                        else:
                            amount = Decimal(row[7].replace(',',''))
                            t = [''] * len(TRANS_HEADER)
                            t[0] = ACCOUNTNAME
                            t[1] = row[0]               # Date
                            if amount > 0:              # TransType
                                t[2] = 'Deposit'
                            else:
                                t[2] = 'Withdraw'
                            t[4] = '*' + BASE_CURRENCY  # Symbol
                            t[9] = abs(amount)
                            t[10] = row[4]              # Comment
                            t[11] = row[3]              # OrderID
                            t[12] = strtoday
                            new_trans.append(t)
                    elif row[2] == 'TRD':
                        if re.search(r'^(\d+\.?\d*)\s', row[4]):
                            # symbol transfer in
                            r = re.search(r'^(\d+\.?\d*)\sShares(.*)\s(.+)$', row[4])
                            if r != None:
                                t = [''] * len(TRANS_HEADER)
                                t[0] = ACCOUNTNAME
                                t[1] = row[0]               # Date
                                t[2] = 'TransferInLongPos'  # TransType
                                t[4] = r.group(3)           # Symbol
                                t[5] = r.group(1)           # Qty
                                t[10] = row[4]              # Comment
                                t[11] = row[3]              # OrderID
                                t[12] = strtoday
                                new_trans.append(t)
                            else:
                                logging.warning('line {0} cannot find symbol in description: "{1}"'.format(srclines, row[4]))
                                r = re.search(r'^(\d+\.?\d*)\s', row[4])
                                t = [''] * len(TRANS_HEADER)
                                t[0] = ACCOUNTNAME
                                t[1] = row[0]               # Date
                                t[2] = 'TransferInLongPos'  # TransType
                                t[5] = r.group(1)           # Qty
                                t[10] = row[4]              # Comment
                                t[11] = row[3]              # OrderID
                                t[12] = strtoday
                                new_trans.append(t)
                        else:
                            r = re.search(r'(\S+)\s([-+]?[\d,\.]+)\s(.+)\s@([\d,\.]+)', row[4])
                            t = [''] * len(TRANS_HEADER)
                            t[0] = ACCOUNTNAME
                            t[1] = row[0]               # Date
                            if r != None:
                                if r.group(1) == 'SOLD':
                                    t[2] = 'Sell'
                                elif r.group(1) == 'BOT':
                                    t[2] = "Buy"
                                else:
                                    logging.warning('line: {0} unknow trade type in Cash Balance. description: "{1}"'.format(srclines, row[4]))
                                t[4] = r.group(3)
                                t[5] = abs(Decimal(r.group(2).replace(',','')))   # Qty
                                t[6] = Decimal(r.group(4).replace(',',''))        # Price
                            else:
                                logging.warning('line: {0} can not recognize desctiption: "{1}"'.format(srclines, row[4]))
                            if len(row[6]) > 0:
                                commission = Decimal(row[6].replace(',',''))
                            else:
                                commission = Decimal('0.0')
                            t[7] = abs(commission)               # Commission & Fee
                            if len(row[7]) > 0:
                                amount = Decimal(row[7].replace(',',''))
                            else:
                                amount = Decimal('0.0')
                                logging.warning('line: {0} TRD in Cash Balance Amount is 0'.format(srclines))
                            amount += commission
                            t[9] = abs(amount)              # Amount
                            t[10] = row[4]                  # Comment
                            t[11] = row[3]                  # OrderID
                            t[12] = strtoday
                            new_trans.append(t)
                    elif row[2] == 'DOI':
                        r = re.search(r'^(\S+).+\s(\S+)$', row[4])
                        t = [''] * len(TRANS_HEADER)
                        t[0] = ACCOUNTNAME
                        t[1] = row[0]           # Date
                        if r:
                            if r.group(1).startswith('DIV'):
                                t[2] = 'Dividend'       # TransType
                            elif r.group(1).startswith('WHTX'):
                                t[2] = 'WithHoldTax'
                            elif r.group(1).startswith('DRIP'):
                                t[2] = 'Buy'
                                logging.warning('line: {0} miss Quantity information for DRIP'.format(srclines))
                                t[3] = "DRIP"
                            else:
                                logging.warning('line: {0} unknow DOI TransType from Description: "{1}"'.format(srclines,row[4]))
                            t[4] = r.group(2)   # Symbol
                        else:
                            logging.warning('line: {0} can not get DiV symbol. description: "{1}"'.format(srclines, row[4]))
                        t[9] = abs(Decimal(row[7].replace(',',''))) # Amount
                        t[10] = row[4]                              # Comment
                        t[11] = row[3]                              # OrderID
                        t[12] = strtoday
                        new_trans.append(t)
                    elif row[2] == '':
                        continue
                    else:
                        logging.warning('line {0} unknow TYPE in Cash Balance Segment.'.format(srclines))
                elif segment == "Account Trade History":
                    src_tradhis.append(row)
                    

# now group same orderID trader in new_trans
targe_trans = []
import copy

for t in new_trans:
    if t[2] != 'Buy' and t[2] != 'Sell':
        targe_trans.append(copy.deepcopy(t))
    else:
        if len(t[11]) == 0:
            targe_trans.append(copy.deepcopy(t))
        else:
            bfind = False
            for tt in targe_trans:
                if tt[11] == t[11]:
                    if tt[2] != t[2]:
                        logging.error(' Group same OrderID trans Eror. Convert Stop!')
                        exit(2)
                    tt[5] += t[5]       # Qty
                    tt[7] += t[7]       # Commission & Fee
                    tt[9] += t[9]       # Amount
                    bfind = True
                    break
            
            if not bfind:
                targe_trans.append(copy.deepcopy(t))

# fix price for buy or sell after group multiple transactions
for t in targe_trans:
    if t[2] in ('Buy', 'Sell'):
        if t[5] and t[9] :  # Qty & Amount must have value
            if t[2] == 'Buy':
                t[6] = (t[9] - t[7])/t[5]
            else:   # Sell
                t[6] = (t[9] + t[7])/t[5]

# fix short sell through Account Trade History and check OrderId is matched
for ath in src_tradhis:
    bfind = False
    transtype = ''
    if ath[3] == 'BUY' and ath[5] == 'TO OPEN':
        transtype = 'Buy'
    elif ath[3] == 'BUY' and ath[5] == 'TO CLOSE':
        transtype = 'BuytoClose'
    elif ath[3] == 'SELL' and ath[5] == 'TO CLOSE':
        transtype = 'Sell'
    elif ath[3] == 'SELL' and ath[5] == 'TO OPEN':
        transtype = 'SelltoOpen'
    else:
        logging.warning('unknow trade operator in Account Trade History. orderID={0}'.format(ath[13]))
    
    # make symbol
    symbol = ath[6]
    if ath[2] != 'STOCK':
        # it is option
        # option symbol format: GOOGL 12FEB18 1256.0 P
        strike = '{0:.1f}'.format(Decimal(ath[8].replace(',','')))
        symbol = symbol + ' ' + ath[7].replace(' ', '') + ' ' + strike + ' ' + ath[9][0]
    for t in targe_trans:
        if t[11] == ath[13]:
            bfind = True
            if transtype and t[2] != transtype:
                logging.info('change Transtype from {0} to {1} at orderID={2}'.format(t[2], transtype, t[11]))
                t[2] = transtype
            if t[4] != symbol:
                logging.info('Replace symbol: {0} with {1} . OrdierID = {2}'.format(t[4], symbol, t[11]))
                t[4] = symbol
            if ath[2] != 'STOCK':
                #refix price for options
                if t[2] == 'Buy' or t[2] == 'BuytoClose':
                    t[6] = (t[9]-t[7])/(t[5]*100)
                elif t[2] == 'Sell' or t[2] =='SelltoOpen':
                    t[6] = (t[9]+t[7])/(t[5]*100)
            break
    if bfind == False:
        logging.warning('there has a  Account Trade history not find in Cash Balance. OrderID={0}'.format(ath[13]))
        pass


if len(targe_trans):
    if args.desfile:
        from shutil import copyfile
        cdesfile = args.desfile + ".bak"
        copyfile(args.desfile, cdesfile)
        desh = open(args.desfile, mode='w', encoding="utf-8", newline='')
    else:
        desh = sys.stdout

    csvwriter = csv.writer(desh)
    header = []
    header.append(TRANS_HEADER)
    csvwriter.writerows(header)
    csvwriter.writerows(old_trans)
    csvwriter.writerows(targe_trans)

    if args.desfile:
        desh.close()

logging.info("Convert successful. total {0} new trasaction appended.".format(len(targe_trans)))