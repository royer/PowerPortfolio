# detect encoding of a file
def detect_file_encoding(filename):
    import chardet

    with open(filename, "rb") as f :
        bob = f.read()
        r = chardet.detect(bob)
    
    return r['encoding']



def CheckPowerPortfolioFile(filename):
    # check file is PowerPortfolio transaction csv file
    #
    # Return:
    #   Tuple(boolean, date<string>)
    #   boolean: True is PowerPortfolio transaction csv file or False
    #   date: the last trade date in file

    from commvar import TRANS_HEADER

    ret = [False, None]
    try:
        with open(filename, encoding="utf-8", newline='') as f:
            import csv
            line = 0
            csvout = csv.reader(f)
            for row in csvout:
                if line == 0:
                    # compare Header
                    if len(TRANS_HEADER) == len(row):
                        for i in range(len(TRANS_HEADER)):
                            if TRANS_HEADER[i].lower() != row[i].lower():
                                ret[0] = False
                    else:
                        ret[0] = False
                        break
                else:
                    if ret[1] == None or (ret[1] != None and ret[1] < row[1]):
                        ret[1] = row[1]
                line+=1
            if line > 0:
                ret[0] = True
               

    except:
        ret[0] = False
    
    return tuple(ret)