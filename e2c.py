#   (A,B,C =>Columns on TH.xlsx)
#   1C  Merge ('A'+'B') into 1 date 
#   2C  Convert type ('C') check with ('F') 
#       Credit-Deposit(USD)=>DEPOSIT    if Deposit(BTC)=>TRANSFER
#       Debit-Withdrawal(USD)=>WITHDRAW if Withdraw(BTC)=>TRANSFER
#   3C  if 2C = BUY/SELL,3C = GEMINI
#   csv

'''
Done
4C  if 5C = USD return ('H')
    elif 5C = BTC return ('K')
    elif 5C = ETH return ('N')
5C  if ('D') ="USD" return USD
    elif ('D')="BTCUSD" return BTC
    elif ('D')="ETHUSD" return ETH
6C  if 5C =(BTC or ETH) return ('H')
8C  =('I')
9C  if 8C not empty return USD
7C  if 6C not empty return USD
'''

from openpyxl import *

fileIN = load_workbook(filename = 'th.xlsx')

ws1 = fileIN['Account History'] 
w2 = fileIN.create_sheet(title="2")
ws2 = fileIN['2']

# 5C    Sorts (ws1[Col'4'] into ws2[Col'5'])
for i in range(2,ws1.max_row):
    if ws1.cell(row=i,column=4).value == "ETH":
        ws2.cell(row=i,column=5).value = ws1.cell(row=i,column=4).value
        
    elif ws1.cell(row=i,column=4).value == "ETHUSD":
        ws2.cell(row=i,column=5).value = "ETH"
        
    elif ws1.cell(row=i,column=4).value == "BTC":
        ws2.cell(row=i,column=5).value = ws1.cell(row=i,column=4).value
        
    elif ws1.cell(row=i,column=4).value == "BTCUSD":
        ws2.cell(row=i,column=5).value = "BTC"
        
    else:
        ws2.cell(row=i,column=5).value = "USD"

# 6C    Sorts (ws1[Col'8'] into ws2[Col'6'])
for i in range(2,ws1.max_row):
    if ws2.cell(row=i,column=5).value == "ETH":
        x = ws1.cell(row=i,column=8).value
        if x != None:   # skips none-type before passing abs
            ws2.cell(row=i,column=6).value = abs(x)
            
    elif ws2.cell(row=i,column=5).value == "BTC":
        x = ws1.cell(row=i,column=8).value
        if x != None:   # skips none-type before passing abs
            ws2.cell(row=i,column=6).value = abs(x)

# 4C
for i in range(2,ws1.max_row):
    if ws2.cell(row=i,column=5).value == "USD":
        x = ws1.cell(row=i,column=8).value
        if x != None:   # skips none-type before passing abs
             ws2.cell(row=i,column=4).value = abs(x)

    elif ws2.cell(row=i,column=5).value == "ETH":
        ws2.cell(row=i,column=4).value = ws1.cell(row=i,column=14).value
        
    elif ws2.cell(row=i,column=5).value == "BTC":
        ws2.cell(row=i,column=4).value = ws1.cell(row=i,column=11).value
        
             
# 8C
for i in range(2,ws1.max_row):
    if ws1.cell(row=i,column=9).value != None: # skips none-type before passing abs
        ws2.cell(row=i,column=8).value = abs(ws1.cell(row=i,column=9).value)
    
# 7C
for i in range(2,ws1.max_row):
    if  ws2.cell(row=i,column=6).value != None:
        ws2.cell(row=i,column=7).value = "USD"

# 9C
for i in range(2,ws1.max_row):
    if  ws2.cell(row=i,column=8).value != None:
        ws2.cell(row=i,column=9).value = "USD"


# saves results into file
fileIN.save(filename = 'test.xlsx')
