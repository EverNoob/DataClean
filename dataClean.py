__author__ = 'Rowbot'
import xlrd
from collections import OrderedDict
import simplejson as json

#objects related to lookup sheet functions
lookup_sheet_names = ['NAV Region', 'Region Rep', 'Brand Code','Varietals Code', 'Brand Associations', 'Brand Name Changes' ]
varietal_dict = {}
distributor_dict = {}



def remove_vintage(n):
    return n[0:-3] + n[-1]

#placeholder
def lookup_varietal(n):
    return varietal_dict[n]

#placeholder
def lookup_distributor(n):
    return n

def build_lookup_tables():
    look_wb = xlrd.open_workbook('Lookup Tables/Lookup Tables.xlsx')
    varietal_dict = build_lookup_dict(look_wb,'Varietals Code',1)
    #distributor_dict = build_lookup_dict(look_wb,'')
    return

def build_lookup_dict(lookup_wb, sheet_name,value_column):
    temp = {}
    lookup_ws = lookup_wb.sheet_by_name(sheet_name)
    for i in range(1,lookup_ws.nrows):
        temp[lookup_ws.cell(i,0).value] = lookup_ws.cell(i,1).value
    return temp



# Open the workbook and select the first worksheet
wb = xlrd.open_workbook('Input/Dirty Data.xlsx')
sh = wb.sheet_by_index(0)

# List to hold dictionaries
data_list = []


build_lookup_tables()
print(varietal_dict)
# Iterate through each row in worksheet and fetch values into dict
# for rownum in range(2, sh.nrows):
for rownum in range(2, 100):
    data_item = OrderedDict()
    row_values = sh.row_values(rownum)
    data_item['Item Code'] = remove_vintage(row_values[3])
    data_item['Brand Code'] = row_values[16]
    data_item['Brand'] = row_values[4]
    data_item['Varietal Code'] = row_values[17]
    data_item['Varietal'] = lookup_varietal(row_values[17])
    data_item['Distributor'] = lookup_distributor(row_values[5])
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

