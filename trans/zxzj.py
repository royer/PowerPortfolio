'''
Convert 中信证券资金流水记录到csv格式，并且从gb2312转换为utf-8

中信证券字段定义
序号   字段名称      长度
1      发生日期      16
2      成交时间      16
3      业务名称      24
4      证券代码      16
5      证券名称      16
6      成交价格      16
7      成交数量      17
8      成交金额      17
9      股份余额      16
10     手续费        14
11     印花税        14
12     过户费        14
13     附加费        14
14     交易所清算费   20
15     发生金额      18
16     资金本次余额  20
17     委托编号      16
18     股东代码      18
19     资金帐号      16
20     币种          14
21     备注          191
'''

FIELDS = [
     '发生日期', 
     '成交时间',
     '业务名称',
     '证券代码',
    '证券名称',
     '成交价格', 
     '成交数量', 
     '成交金额', 
     '股份余额',
     '手续费',
     '印花税', 
     '过户费',
     '附加费', 
     '交易所清算费', 
     '发生金额', 
     '资金本次余额', 
     '委托编号',
     '股东代码', 
     '资金帐号', 
     '币种', 
     '备注' 
]


import argparse
import sys
import os
import logging
import re
import decimal

import csv

from datetime import datetime

from commvar import TRANS_HEADER

TRANS_ROOTDIR = os.path.dirname(os.path.abspath(__file__))
SCRIPT_FILENAME = os.path.basename(os.path.abspath(__file__))

# The log file path
LOGFILE = os.path.join(TRANS_ROOTDIR,"log", re.sub(r"\.py",".log", SCRIPT_FILENAME, flags=re.I))


# load symbol map
from symbolmap_zxzj import SYMBOLMAP


argParser = argparse.ArgumentParser()
argParser.add_argument("srcfile", help="中信证券资金流水 TXT格式文件(GBK 编码)", nargs='?')
argParser.add_argument('desfile',help="PowerPortfolio Transacton csv file(utf-8)",nargs='?')
argParser.add_argument("-A", "--account", dest="accountname", help="The account name in PowerPortfolio", default="中信证券")
argParser.add_argument("-B", "--brokeraccid", dest="acid_in_broker", 
    help="The accountid in Broker. for check is correct transaction file.")
argParser.add_argument("-s","--skiplines", help="Skip first n lines. default is 2", type=int, default=2)
args = argParser.parse_args()

ACCOUNTNAME = args.accountname
ACCOUNTID_IN_BROKER = args.acid_in_broker
SKIPLINES = args.skiplines


import locale
if os.name == 'nt':
    locale.setlocale(locale.LC_ALL,locale="USA")
else:
    locale.setlocale(locale.LC_ALL, 'en-US.UTF-8')

DESFILE = args.desfile
# make sure log file directory is exist
if not os.path.exists(os.path.dirname(os.path.abspath(LOGFILE))):
    os.makedirs(os.path.dirname(os.path.abspath(LOGFILE)))

logging.basicConfig(filename=LOGFILE, level=logging.DEBUG, filemode='w',\
    format='%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S' )
consolehandler = logging.StreamHandler()
consolehandler.setLevel(logging.WARNING)
consolehandler.setFormatter(fmt=logging.Formatter("%(levelname)s %(message)s"))
logging.getLogger().addHandler(consolehandler)

logging.info("Start convert {srcfile} to {outfile}".format(srcfile=args.srcfile if args.srcfile else 'STDIN', outfile=DESFILE if DESFILE else 'STDOUT'))


bwriteheader = True
bbackup = False
# make sure dest file directory is exist
if DESFILE:
    try:
        if not os.path.exists(os.path.dirname(os.path.abspath(DESFILE))):
            os.makedirs(os.path.dirname(DESFILE))
            logging.info("created dest file({0}) directory successful.".format(os.path.abspath(DESFILE)))
    except OSError as err:
        str = "Creating dest file: {0} directory Error: ".format(os.path.abspath(DESFILE)) + err.strerror
        logging.error(str)
        exit(1)


# Get Last Trade Date from Append File if this file exist
LAST_TRADE_DATE = None

from commfunc import CheckPowerPortfolioFile

if DESFILE and os.path.exists(os.path.abspath(DESFILE)):
    bok, LAST_TRADE_DATE = CheckPowerPortfolioFile(DESFILE)
    if bok:
        bwriteheader = False
        bbackup = True





srcfile = open(args.srcfile, encoding="gbk", mode='r') if args.srcfile else sys.stdin

cblines = 0

fields = []
newfilelines = []

for line in srcfile:
    cblines += 1
    if cblines <= SKIPLINES:
        continue
    gbline = line.encode('gbk')

    newline = []

    if cblines == 3:
        # Header line
        cbfields = 0
        fieldstart = 0
        inname = True
        i = 0
        while i < len(gbline):
            if inname:
                if gbline[i] == ord(' '):
                    inname = False
                    n = gbline[fieldstart:i]
                    n = n.decode('gbk')
                    field = {'name': n, 'length': 0}
                    fields.append(field)
            else:
                if gbline[i] != ord(' '):
                    inname = True
                    inspace = False
                    fields[-1]['length'] = i - fieldstart
                    fieldstart = i
            i += 1

        # Check Header is correct
        if len(fields) != 21:
            logging.error("File:{0} 字段个数不是21个。好像不是中信证券的资金流水文件。 程序退出".format(srcfile.name))
            exit(1)
        for f in fields:
            if not f['name'] in FIELDS:
                logging.error("'{0}' 不认识的字段。好像不是中信证券的资金流水文件。 程序退出".format(f['name']))
                exit(1)
        
        # make header line
        newline = list(x['name'] for x in fields)
    else:
        start = 0
        i = 0
        for field in fields:
            if i < len(fields)-1:
                d = gbline[start:start+field['length']].decode('GBK')
            else:
                d = gbline[start:].decode('GBK')
            start += field['length']
            newline.append(d.strip())
            i += 1

    newfilelines.append(newline)

srcfile.close()

convert_date = datetime.now().strftime("%Y-%m-%d")

# 针对每天一笔的港股通费用进行合并成一条
specialfee = []
specialfeerows = 0
trans = []

firstdate = None
lastdate = None
# iteral source transaction record skip head line
fileline = 4
for srcline in newfilelines[1:]:
    tradedate = srcline[0][0:4] + '-' + srcline[0][4:6] + '-' + srcline[0][6:8]
    if ACCOUNTID_IN_BROKER and ACCOUNTID_IN_BROKER != srcline[18]:
        logging.warning("special AccountID({0}) in Broker not match accountid in file ({1}). skip this line.".format(ACCOUNTID_IN_BROKER, srcline[18]))
        fileline += 1
        continue
    if LAST_TRADE_DATE and tradedate <= LAST_TRADE_DATE:
        logging.warning("line: {0} date is early than last record date: {1}".format(fileline, LAST_TRADE_DATE))
        fileline += 1
        continue
    if not firstdate or (firstdate and srcline[0] < firstdate):
        firstdate = srcline[0]
    if not lastdate or (lastdate and srcline[0] > lastdate):
        lastdate = srcline[0]
    
    if srcline[2] == "港股通组合费收取": #业务名称是港股通组合费收取 进行特殊处理
        if len(specialfee) == 0:
            emptytrans = [''] * len(TRANS_HEADER)
            emptytrans[0] = ACCOUNTNAME
            emptytrans[2] = 'OtherFee'      # TransType
            emptytrans[3] = ''              # SubTransType
            emptytrans[4] = '*CNY'          # Symbol
            emptytrans[9] = decimal.Decimal('0')               # Amount
            emptytrans[12] = convert_date

            specialfee.append(emptytrans)
        specialfee[0][1] = tradedate # Date
        specialfee[0][11] = locale.atoi(srcline[16])  #OrderID
        specialfee[0][9] += (-1 * decimal.Decimal(srcline[14])) # 从发生金额字段得到 Amount
        specialfeerows += 1
    elif srcline[2] == "股息入帐":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME # AccountName
        t[1] = tradedate     # Date
        t[2] = 'Dividend'    # TransType
        t[4] = srcline[3] + ".CN" #Symbol
        t[9] = decimal.Decimal(srcline[14]) # Ammount
        t[10] = srcline[20]   # Comment
        t[11] = locale.atoi(srcline[16]) #OrderID
        t[12] = convert_date
        trans.append(t)
    elif srcline[2] == "股息红利税补缴":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME
        t[1] = tradedate
        t[2] = 'WithHoldTax'
        ss = r"证券代码:'(\d{6})"
        rs = re.search(ss, srcline[20])
        if rs:
            t[4] = rs.group(1) + ".CN" #Symbol
        else:
            logging.warning("Can not find symbol in note for 股息红利税补缴 record. use *CNY as symbol. line: {0}".format(fileline))
            t[4] = "*CNY"
        t[9] = -1 * decimal.Decimal(srcline[14])  # Ammount
        t[10] = srcline[20]     #Comment
        t[11] = locale.atoi(srcline[16])  # OrderID
        t[12] = convert_date

        trans.append(t)
    elif srcline[2] == "银行转取":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME      # AccountName
        t[1] = tradedate        # Date
        t[2] = "Withdraw"       # TransType
        t[4] = "*CNY"           # Symbol
        t[9] = -1 * decimal.Decimal(srcline[14]) # Ammount
        #t[10] = srcline[20]     # Comment
        t[11] = locale.atoi(srcline[16])    #Order ID
        t[12] = convert_date
        trans.append(t)
    elif srcline[2] == "证券买入":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME
        t[1] = tradedate
        t[2] = "Buy"
        t[5] = abs(decimal.Decimal(srcline[6])) # Qty
        t[6] = abs(decimal.Decimal(srcline[5])) # Price
        #t[10] = srcline[20] # Comment
        t[11] = locale.atoi(srcline[16]) # Order ID
        t[12] = convert_date    # Convert Date
        fee_cny = abs(decimal.Decimal(srcline[9]))      # 手续费
        fee_cny += abs(decimal.Decimal(srcline[10]))    # 印花税
        fee_cny += abs(decimal.Decimal(srcline[11]))    # 过户费
        fee_cny += abs(decimal.Decimal(srcline[12]))    # 附加费
        fee_cny += abs(decimal.Decimal(srcline[13]))    # 附加费
        amount_cny = abs(decimal.Decimal(srcline[14]))  # Ammount CNY
        if srcline[20].startswith('港股通'):
            t[4] = srcline[3][1:]+".HK"             # Symbol
            t[10] = srcline[20]                     # Comment
            # get fx rate from note
            ss = r"汇率:(\d+\.\d+)"
            r = re.search(ss, srcline[20])
            if not r:
                logging.error("Can not find fxrate from Note. line: {0} stop convert.".format(fileline))
                exit(1)
            fxrate = decimal.Decimal(r.group(1))
            t[7] = fee_cny / fxrate
            t[9] = amount_cny / fxrate
            amount_hkd = t[9]
            t[6] = (t[9] - t[7]) / t[5]             # recorrect price
            trans.append(t)                        # insert trade hkd transaction

            # insert Currency trans
            # Buy HKD
            t = [''] * len(TRANS_HEADER)
            t[0] = ACCOUNTNAME
            t[1] = tradedate
            t[2] = "BuyCurrency"
            t[4] = "*HKD"
            t[5] = amount_cny       # Qty
            t[6] = fxrate
            t[9] = amount_hkd
            t[10] = "Insert Fake Currency for Buy {0} ".format(srcline[3][1:]+".HK")
            t[11] = locale.atoi(srcline[16])
            t[12] = convert_date
            trans.append(t)
            # Buy Currency Pair for CNY
            t = [''] * len(TRANS_HEADER)
            t[0] = ACCOUNTNAME
            t[1] = tradedate
            t[2] = "BuyCurrency_Pair"
            t[4] = "*CNY"
            t[5] = amount_cny       # Qty
            t[6] = 1                # Price
            t[9] = amount_cny       # amount
            t[11] = locale.atoi(srcline[16])
            t[12] = convert_date
            trans.append(t)
        else:
            t[4] = srcline[3] + ".CN"
            t[7] = fee_cny                          # Fee
            t[9] = amount_cny                       # Amount
            # recorrect price
            t[6] = (t[9] - t[7])/t[5]
            trans.append(t)
    elif srcline[2] == "证券卖出":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME
        t[1] = tradedate
        t[2] = "Sell"
        t[5] = abs(decimal.Decimal(srcline[6])) # Qty
        t[6] = abs(decimal.Decimal(srcline[5])) # Price
        #t[10] = srcline[20] # Comment
        t[11] = locale.atoi(srcline[16]) # Order ID
        t[12] = convert_date    # Convert Date
        fee_cny = abs(decimal.Decimal(srcline[9]))      # 手续费
        fee_cny += abs(decimal.Decimal(srcline[10]))    # 印花税
        fee_cny += abs(decimal.Decimal(srcline[11]))    # 过户费
        fee_cny += abs(decimal.Decimal(srcline[12]))    # 附加费
        fee_cny += abs(decimal.Decimal(srcline[13]))    # 附加费
        amount_cny = abs(decimal.Decimal(srcline[14]))  # Ammount CNY
        if srcline[20].startswith('港股通'):
            t[4] = srcline[3][1:]+".HK"             # Symbol
            t[10] = srcline[20]                     # Comment
            # get fx rate from note
            ss = r"汇率:(\d+\.\d+)"
            r = re.search(ss, srcline[20])
            if not r:
                logging.error("Can not find fxrate from Note. line: {0} stop convert.".format(fileline))
                exit(1)
            fxrate = decimal.Decimal(r.group(1))
            t[7] = fee_cny / fxrate
            t[9] = amount_cny / fxrate
            amount_hkd = t[9]
            t[6] = (t[9] + t[7]) / t[5]             # recorrect price
            trans.append(t)                        # insert trade hkd transaction

            # insert Currency trans
            # Sell HKD
            t = [''] * len(TRANS_HEADER)
            t[0] = ACCOUNTNAME
            t[1] = tradedate
            t[2] = "SellCurrency"
            t[4] = "*HKD"
            t[5] = amount_cny       # Qty
            t[6] = fxrate
            t[9] = amount_hkd
            t[10] = "Insert Fake Currency for Sell {0} ".format(srcline[3][1:]+".HK")
            t[11] = locale.atoi(srcline[16])
            t[12] = convert_date
            trans.append(t)
            # Cell Currency Pair for CNY
            t = [''] * len(TRANS_HEADER)
            t[0] = ACCOUNTNAME
            t[1] = tradedate
            t[2] = "SellCurrency_Pair"
            t[4] = "*CNY"
            t[5] = amount_cny       # Qty
            t[6] = 1                # Price
            t[9] = amount_cny       # amount
            t[11] = locale.atoi(srcline[16])
            t[12] = convert_date
            trans.append(t)
        else:
            t[4] = srcline[3] + ".CN"
            t[7] = fee_cny                          # Fee
            t[9] = amount_cny                       # Amount
            # recorrect price
            t[6] = (t[9] + t[7])/t[5]
            trans.append(t)
    elif srcline[2] == "利息归本":
        t = [''] * len(TRANS_HEADER)
        t[0] = ACCOUNTNAME      # AccountName
        t[1] = tradedate        # Date
        t[2] = "Interest"       # TransType
        t[4] = "*CNY"           # Symbol
        t[9] = decimal.Decimal(srcline[14]) # Ammount
        #t[10] = srcline[20]     # Comment
        t[11] = locale.atoi(srcline[16])    #Order ID
        t[12] = convert_date
        trans.append(t)
    else:
        logging.error("Unknow 业务名称: {0} line: {1}. convert stop.".format(srcline[2], fileline))
        exit(1)

    fileline += 1
    

if len(specialfee) > 0:
    trans.append(specialfee[0])
    specialfee[0][10] = "港股通组合费 " + firstdate + " - " + lastdate
    logging.info("合并了 {0} 条港股通组合费".format(specialfeerows))


# sort all transaction by Date
from operator import itemgetter
pp_trans_sorted = sorted(trans, key=itemgetter(1))


header = []
header.append(TRANS_HEADER)

if len(pp_trans_sorted):
    if DESFILE and bbackup:
        from shutil import copyfile
        ctagfile = DESFILE + ".bak"
        copyfile(DESFILE, ctagfile)
    # write to dest file
    desfile = open(DESFILE, mode='a', newline='', encoding='utf-8') if DESFILE else sys.stdout

    csv_desfile_writer = csv.writer(desfile)
    if bwriteheader:
        csv_desfile_writer.writerows(header)
    csv_desfile_writer.writerows(pp_trans_sorted)

    if DESFILE:
        desfile.close()
    logging.info("write {0} transactions to {1}".format(len(pp_trans_sorted), DESFILE if DESFILE else "STDOUT"))
else:
    print("There is no transactons converted.")

# Append to APPENDFILE
# writeheader = False if os.path.exists(APPENDFIlE) else True

# appendfile = open(APPENDFIlE, mode='a', newline='', encoding='utf-8')

# csv_af_writer = csv.writer(appendfile)
# if writeheader:
#     csv_af_writer.writerows(header)
# csv_af_writer.writerows(pp_trans_sorted)
# appendfile.close()
# logging.info("Append {0} transactions to {1}".format(len(pp_trans_sorted), APPENDFIlE))

logging.info("Convert Successful. total convert {0} transactions".format(len(newfilelines)-1))





