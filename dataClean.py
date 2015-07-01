__author__ = 'Rowbot'
import xlrd
from collections import OrderedDict
import simplejson as json


def remove_vintage(n):
    return n[0:-3] + n[-1]

#placeholder
def lookup_vareital(n):
    return n

# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('Input/Dirty Data.xlsx')
sh = wb.sheet_by_index(0)

# List to hold dictionaries
data_list = []
print (sh.nrows)
# Iterate through each row in worksheet and fetch values into dict
# for rownum in range(2, sh.nrows):
for rownum in range(2, 100):
    data_item = OrderedDict()
    row_values = sh.row_values(rownum)
    data_item['Item Code'] = remove_vintage(row_values[3])
    data_item['Brand Code'] = row_values[16]
    data_item['Brand'] = row_values[4]
    data_item['Varietal Code'] = row_values[17]
    data_item['Vareital'] = lookup_vareital(row_values[17])
    data_item['Distributor'] = row_values[5]
    data_item['State'] = row_values[6]
    data_item['Sales Rep'] = row_values[7]
    data_item['Item ID'] = row_values[8]
    data_item['Item'] = row_values[9]
    data_item['SKU Tag'] = row_values[10]
    data_item['Item Pre'] = row_values[11]


    data_list.append(data_item)

print ("the size of data_list: " + str(len(data_list)))
# Serialize the list of dicts to JSON
j = json.dumps(data_list, sort_keys=True, indent=4 * ' ')
print(j)

# Write to file
with open('data.json', 'w') as f:
    f.write(j)

