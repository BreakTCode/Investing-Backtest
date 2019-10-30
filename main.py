import pandas as pd
from datetime import date as dt
import xlsxwriter as xlw
import sys

#Classes

'''

ATIS is a object that has attributes to track the different 
elements of a particular investment and calculate their total
return. 

It expects the following to initialize:

startdate: The date that we're starting this investment test.
sheets: The pandas object that reads our excel data file.

'''
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

        # Test if a year from now is in the list of dates
        # otherwise, add a day to the start day until you
        # can locate approx. year from start.
        # 
        # This resolves issues with weekends and holidays
        # as the stock market isn't open on weekends.
         
        while YEARENDDATE == "":
            if sheet[sheets.loc(startdate + 365)]:
                YEARENDDATE = startdate + 365
            else:
                startdate += 1
        self.YND    = YEARENDDATE


#Variables
Row = 0
Meth1A = [0]
Meth2A = [0]
LoopRow = 0

# Initializes the pandas object to read our excel file and our xlsxwriter.
sheets = pd.read_excel('^GSPC (1990-2010).xlsx', header=0)
writer = xlw.Workbook('ATIDown ' + str(dt.today()) + '.xlsx')
worksheet1 = writer.add_worksheet('Dates and values')

# While our current row in the excel file is less than the overall shape:
    # If it's our first row
        # Set the first startdate and create the objects
    # Else 
        # Append it to our list of objects.
while Row < sheets.shape[0]:
    if Row == 0:
        Meth1A[0] = ATIS(sheets['Date'][0], sheets)
        Meth2A[0] = ATIS(sheets['Date'][0], sheets)
        Meth2A[0].INVD1 = sheets['Date'][1]
    else:
        Meth1A.append(ATIS(sheets['Date'][Row], sheets))
        Meth2A.append(ATIS(sheets['Date'][Row], sheets))
        Meth2A[Row].INVD1 = sheets['Date'][Row + 1]
    # Method 1
    # While our current row within the year timeframe is less than the end date
        # If our current close value is less than the prior days close value
            # If we're out of the 100k to invest
                # Do nothing
            # Else
                # If it's our first investment
                    # Update investment day 1.
                # If it's our first cost value
                    # Set initial values of Cost and shares.
                    # Subtract 20k from our total amount we have left to invest. 
                # Else
                    # Append the to list of values
                    # If we've spent the last of the 100k
                        # Set our completion date to current row date.
    while sheets['Date'][LoopRow].day < Meth1A[Row].YEARENDDATE:
        if sheets['Close'][LoopRow] < sheets['Close'][LoopRow - 1]:
            if Meth1A[LoopRow].ONEHK == 0:
                None
            else:
                if Meth1A[Row].ONEHK == 100000:
                    Meth1A[Row].INVD1 = sheets['Date'][LoopRow]
                if Meth1A[Row].COST[0] == 0:
                    Meth1A[Row].COST[0] = sheets['Close'][LoopRow]
                    Meth1A[Row].SHARES = 20000 / sheets['Close'][LoopRow]
                    Meth1A[Row].ONEHK = Meth1A[Row].ONEHK - 20000
                else:
                    Meth1A[Row].COST.append(sheets['Close'][LoopRow])
                    Meth1A[Row].SHARES += 20000 / sheets['Close'][LoopRow]
                    Meth1A[Row].ONEHK = Meth1A[Row].ONEHK - 20000
                    if Meth1A[Row].ONEHK == 0:
                        Meth1A[Row].COMPD = sheets['Date'][LoopRow]
    Meth1A[Row].FVAL = ((sheets['Close'][LoopRow] *  Meth1A[Row].SHARES) / 100000) - 1

    for i in Meth1A:
        print(Meth1A[i].FVAL)
    #Method 2



writer.close()


