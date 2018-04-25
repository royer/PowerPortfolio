'''
Convert TD Canada - TD Direct Investing Activity csv file to PowerPortfolio csv or Banktivity transaction file.
'''

import os
import sys
import datetime
import re
import csv
import decimal
from decimal import Decimal

PPF_TRANS_HEADER = (
    'Account',          # 0
    'Date',             # 1
    'TransType',        # 2
    'SubTransType',     # 3
    'Symbol',           # 4
    'Qty',              # 5
    'Price',            # 6
    'Fee',              # 7
    'AccruedInterest',  # 8
    'Amount',           # 9
    'Comment',          # 10
    'OrderID',          # 11
    'ConvertDate'       # 12
)

MONTH_2LTO3L = {
    # stupid TD use  month 2 letter abbreviation in option name, we need map it to 3 letter
    'JA': 'JAN',
    'FB': 'FEB',
    'MR': "MAR",
    'AP': 'APL',
    'MA': 'MAY',
    'JN': 'JUN',
    'JL': 'JUL',
    'AU': "AUG",
    'SE': 'SEP',
    'OC': 'OCT',
    'NO': 'NOV',
    'DE': 'DEC'
}

from symbolmap_tdc import SYMBOLMAP

ugly_oname_map = {}

def translate_symbol(description, target):
    ms = re.compile(r"^(?P<otype>\D{3,4})\s+(?P<multp>-?\d+)\s+(?P<symbol>[a-zA-Z]+)'(?P<year>\d{2})\s+(?P<smonth>\D{2})@(?P<strike>\d+\.?\d*)")
    ml = re.compile(r'^(?P<op>[A-Z]+)\s.+(?P<expire>[A-Z]{3} \d{1,2},\d{4})$')
    symbol = '' 
    opcode = ''
    if description.startswith('PUT') or description.startswith('CALL'):
        rs = ms.search(description)
        if rs:
            if rs.group('smonth') not in MONTH_2LTO3L:
                print('2 letter month abbreviation: {0} not exist in my map! convert abort'.format(rs.group('smonth')), file=sys.stderr)
                exit(1)
            strike = '{0:.1f}'.format(Decimal(rs.group('strike')))
            uglysymbol = rs.group('symbol') + ' ' + MONTH_2LTO3L[rs.group('smonth')] + rs.group('year') + ' ' + strike + ' ' + rs.group('otype')[0]
            if uglysymbol in ugly_oname_map:
                symbol = ugly_oname_map[uglysymbol]
            else:
                symbol = uglysymbol
            sl = ms.sub('', description).strip()
            if sl:
                rl = ml.search(sl)
                if rl:
                    exp = datetime.datetime.strptime(rl.group('expire'),'%b %d,%Y').strftime('%d%b%y').upper()
                    nicesymbol = rs.group('symbol') + ' ' + exp + ' ' + strike + ' ' + rs.group('otype')[0]
                    ugly_oname_map[uglysymbol] = nicesymbol
                    symbol = nicesymbol

        else:
            print('can not recognized option symbol: {0}! convert abort'.format(description), file=sys.stderr)
            exit(1)
    else:
        symbol = SYMBOLMAP[description][target]
    return symbol, opcode

import argparse

argParser = argparse.ArgumentParser(description='Convert TD Canada Direct Inversting Activity csv file to PowerPortfolio transaction file or Banktivity transaction file')
argParser.add_argument("srcfile", help="TD Direct Investing Activity csv file.")
argParser.add_argument("desfile", help="out put file", nargs='?')
argParser.add_argument("-A", "--account", help="PowerPortfolio Account name. default is TDUSD", default="TDUSD")
argParser.add_argument('-C', '--currency', help="the base currency of this account. default: USD", default='USD')
argParser.add_argument('-t','--target', help="convert to PowerPortfolio or Banktivity. PPF: PowerPortfolio; BKT: Banktivity. default is PPF", default="PPF")

args = argParser.parse_args()

ACCOUNTNAME = args.account
DEFCURRENCY = args.currency

TARGET = args.target.upper()
if TARGET not in ('PPF', 'BKT'):
    print('target only support PPF or BKT. unkonwn target: {0}'.format(TARGET), file=sys.stderr)
    exit(1)

import chardet

if not os.path.exists(args.srcfile):
    print("Source file: '{0}' doesn't exist, Converter exit.".format(args.srcfile))
    exit(1)

if args.desfile and os.path.exists(args.desfile):
    c = input('File: {0} already exist, Overwrite?(y/n)'.format(args.desfile))
    if c.lower() != 'y':
        exit(1)

with open(args.srcfile, mode='rb') as f:
    bob = f.read()
    src_encoding = chardet.detect(bob)['encoding']

ppf_trans = []      # the power portfolio transaction
bkt_trans = []      # the banktivkty transaction

stoday = datetime.datetime.now().strftime('%Y-%m-%d')

with open(args.srcfile, mode="r", newline='', encoding=src_encoding) as fs:
    csvr = csv.reader(fs)
    line = 0
    for row in csvr:
        line += 1
        if line < 5:
            continue        # skip first 4 line
        tdate = (datetime.datetime.strptime(row[0],'%d %b %Y')).strftime('%Y-%m-%d')
        namount = Decimal(row[7])
        
        commission = Decimal(row[6]) if row[6] else Decimal('0.0')
        shares = Decimal(row[4]) if row[4] else Decimal('0.0')

        ppft = [''] * len(PPF_TRANS_HEADER)
        ppft[0] = ACCOUNTNAME
        ppft[1] = tdate
        ppft[12] = stoday
        if row[3] == 'DIV':
            ppft[2] = 'Dividend'
            symbol,op = translate_symbol(row[2], TARGET )
            ppft[4] = symbol
            ppft[9] = abs(namount)
        elif row[3] == 'WHTX02':
            ppft[2] = 'WithHoldTax'
            symbol,op = translate_symbol(row[2], TARGET )
            ppft[4] = symbol
            ppft[9] = abs(namount)
        elif row[3] == 'DRIP':
            ppft[2] = 'Buy'
            ppft[3] = 'DRIP'
            symbol,op = translate_symbol(row[2], TARGET )
            ppft[4] = symbol
            ppft[5] = shares
            ppft[9] = abs(namount)
        elif row[3] == 'BUY':
            ppft[2] = "Buy"
            symbol,op = translate_symbol(row[2], TARGET )
            ppft[4] = symbol
            if op == 'CLOSING':
                ppft[2] = 'BuytoClose'
            ppft[5] = abs(shares)
            if symbol.endswith(' P') or symbol.endswith(' C'):
                multp = 100
            else:
                multp = 1
            price = (abs(namount) - abs(commission))/(abs(shares)*multp)
            ppft[6] = price
            ppft[7] = abs(commission)
            ppft[9] =abs(namount)
        elif row[3] == 'SELL':
            ppft[2] = 'Sell'
            symbol,op = translate_symbol(row[2], TARGET )
            ppft[4] = symbol
            if op == 'OPENING':
                ppft[2] = 'SelltoOpen'
            ppft[5] = abs(shares)
            if symbol.endswith(' P') or symbol.endswith(' C'):
                multp = 100
            else:
                multp = 1
            price = (abs(namount) + abs(commission))/(abs(shares)*multp)
            ppft[6] = price
            ppft[7] = abs(commission)
            ppft[9] =abs(namount)            
        elif row[3] == 'CONV':
            if namount < 0:
                ppft[2] = "Withdraw"
            else:
                ppft[2] = 'Deposit'
            ppft[3] = 'CONV'
            ppft[4] = '*' + DEFCURRENCY
            ppft[9] = abs(namount)
            if commission != 0:
                ppft[7] = commission
            ppft[10] = row[2]   # comment
        else:
            print("Unknow Action: {0}. convert stop!".format(row[3]), file=sys.stderr)
            exit(1)
        
        ppf_trans.append(ppft)


from operator import itemgetter
ppf_trans_sorted = sorted(ppf_trans, key=itemgetter(1))

if args.desfile:
    f = open(args.desfile, mode='w', newline='', encoding='utf-8')
else:
    f = sys.stdout

csvw = csv.writer(f)

if TARGET == 'PPF':
    header = []
    header.append(PPF_TRANS_HEADER)
    csvw.writerows(header)
    csvw.writerows(ppf_trans_sorted)

if args.desfile:
    f.close()
