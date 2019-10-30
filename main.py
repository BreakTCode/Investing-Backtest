import pandas as pd
from datetime import date as dt
import xlsxwriter as xlw
import sys
#Objects
class ATIS(object):
    def __init__(self, startdate, sheets):
        if startdate.day + 365 > sheets['Date'][sheets.shape[0]].day:
            sys.exit()
        self.ONEHK   = 100000
        self.SHARES = 0
        self.INVD1  = ""
        self.COMPD  = ""
        self.FVAL   = ""
        self.COST   = [0]
        self.STARTD = startdate
        YEARENDDATE = ""
        count = 0

        while YEARENDDATE == "":
            if sheet[sheets.loc(startdate + 365)]:
                YEARENDDATE = startdate + 365
            else:
                startdate += 1
        self.YND    = YEARENDDATE


#Variables
count = 0
Meth1A = [0]
Meth2A = [0]
Curweek = 0

sheets = pd.read_excel('^GSPC (1990-2010).xlsx', header=0)
writer = xlw.Workbook('ATIDown ' + str(dt.today()) + '.xlsx')
worksheet1 = writer.add_worksheet('Dates and values')

while count < sheets.shape[0]:
    if count == 0:
        Meth1A[0] = ATIS(sheets['Date'][0], sheets)
        Meth2A[0] = ATIS(sheets['Date'][0], sheets)
        Meth2A[0].INVD1 = sheets['Date'][1]
    else:
        Meth1A.append(ATIS(sheets['Date'][count], sheets))
        Meth2A.append(ATIS(sheets['Date'][count], sheets))
        Meth2A[count].INVD1 = sheets['Date'][count + 1]
    #Method 1
    while sheets['Date'][count].day < Meth1A[count].YEARENDDATE:
        if sheets['Close'][j] < sheets['Close'][j - 1]:
            if Meth1A[count].ONEHK == 0:
                None
            else:
                if Meth1A[count].ONEHK == 100000:
                    Meth1A[count].INVD1 = sheets['Date'][count]
                if Meth1A[count].COST[0] == 0:
                    Meth1A[count].COST[0] = sheets['Close'][j]
                    Meth1A[count].SHARES = 20000 / sheets['Close'][j]
                    Meth1A[count].ONEHK = Meth1A[count].ONEHK - 20000
                else:
                    Meth1A[count].COST.append(sheets['Close'][j])
                    Meth1A[count].SHARES += 20000 / sheets['Close'][j]
                    Meth1A[count].ONEHK = Meth1A[count].ONEHK - 20000
                    if Meth1A[count].ONEHK == 0:
                        Meth1A[count].COMPD = sheets['Date'][count]
    Meth1A[count].FVAL = sheets['Close'][j] *  Meth1A[count].SHARES / 100000 - 1

    for i in Meth1A:
        print(Meth1A[i].FVAL)
    #Method 2



writer.close()


