from __future__ import print_function
from datetime import datetime

from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet


# this file is Nested_sum.xlsx, but needs to be opened in Excel
# re-calc-ed, so that all cells have a starting value, and saved
filename = "Nested_sum_koala2.xlsx"

### Graph Generation ###
beginning = datetime.now()
print("loading file")
c = ExcelCompiler(filename)
print("excel compiler made", datetime.now() - beginning)
sp = c.gen_graph()
print("graph generated", datetime.now() - beginning)

columns = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
# columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U']
# columns = ['V', 'W', 'X', 'Y', 'Z']
# columns = ['H']

max_rows = 3
# max_rows = 3
addresses = ['Sheet1!{}{}'.format(column, row) for row in range(2, max_rows) for column in columns]

print("addresses made", datetime.now() - beginning)
print("len(addresses)", len(addresses))
second_beginning = datetime.now()
for address in addresses:
    evaluated_value = sp.evaluate(address)
    sp.set_value(address, evaluated_value)
    print("EVALUATED VALUE", address, evaluated_value)

print("Evaluation done", datetime.now() - second_beginning)
print("all done", datetime.now() - beginning)
