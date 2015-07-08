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

#open lookup workbook and grab worksheets
wb_lookup = xlrd.open_workbook('Lookup Tables/Lookup Tables.xlsx')
ws_navregion = wb_lookup.sheet_by_name('NAV Region')
ws_regionrep = wb_lookup.sheet_by_name('Region Rep')
ws_brandcode = wb_lookup.sheet_by_name('Brand Code')
ws_varietalscode = wb_lookup.sheet_by_name('Varietals Code')
ws_brandassoc = wb_lookup.sheet_by_name('Brand Associations')
ws_brandnamechanges = wb_lookup.sheet_by_name('Brand Name Changes')


# Dictionary to hold dictionaries
raw_data_list = []

def build_json_from_ws(ws):
    temp_list = []
    column_titles = list(ws.row_values(0))
    for rownum in range(1, ws.nrows):
        row_values = ws.row_values(rownum)
        temp_dict = {}
        for titlenum in range(0,len(column_titles)):
            temp_dict[column_titles[titlenum]] = row_values[titlenum]
        temp_list.append(temp_dict)
    return temp_list

def build_navregion_lookup_dict(ws_navregion):
    temp_dict = {}
    for rownum in range(1,ws_navregion.nrows):
        row_values = ws_navregion.row_values(rownum)
        temp_dict[row_values[0]] = row_values[1]
    return temp_dict

def build_regionrep_lookup_dict(ws_regionrep):
    temp_dict = {}
    for rownum in range(1,ws_regionrep.nrows):
        row_values = ws_regionrep.row_values(rownum)
        temp_dict[row_values[0]] = row_values[1]
    return temp_dict

def build_twoitem_lookup_dict(ws,reverse=False):
    temp_dict = {}
    for rownum in range(1,ws.nrows):
        row_values = ws.row_values(rownum)
        if(reverse):
            temp_dict[row_values[1]] = row_values[0]
        else:
            temp_dict[row_values[0]] = row_values[1]
    return temp_dict


def prettyprint(d):
    print(json.dumps(d, sort_keys=True, indent=4 * ' '))

prettyprint(build_navregion_lookup_dict(ws_navregion))
prettyprint(build_regionrep_lookup_dict(ws_regionrep))
prettyprint(build_twoitem_lookup_dict(ws_brandcode,reverse=True))

def build_raw_list():
    global i, temp_list, rownum, row_values, data_item
    # Iterate through each row in worksheet and fetch values into dict
    # for rownum in range(2, sh.nrows):
    i = 0
    temp_list = []
    print("the number of rows is: " + str(sh.nrows))
    for rownum in range(2, sh.nrows):
        row_values = sh.row_values(rownum)
        if row_values[2] != 'PAC' and "CONSUMER" not in row_values[0] and "Kirkland" not in row_values[
            4] and "zBarter" not in row_values[0] and row_values[2] != "":
            data_item = {}
            i += 1
            print(i)
            row_values = sh.row_values(rownum)
            data_item['Customer Name'] = row_values[0]
            data_item['Ship-to State'] = row_values[1]
            data_item['Salesperson Code'] = row_values[2]
            data_item['Item No.'] = row_values[3]
            data_item['Description'] = row_values[4]
            data_item['Product Group Code'] = row_values[5]
            data_item['Posting Month'] = lookup_distributor(row_values[6])
            data_item['Posting Date'] = str(row_values[7])
            data_item['Document Type'] = row_values[8]
            data_item['Location Code'] = row_values[9]
            data_item['Document No.'] = row_values[10]
            data_item['Quantity'] = row_values[11]
            data_item['Sales Amount (Actual)'] = row_values[12]
            data_item['Quantity (positive)'] = lookup_distributor(row_values[13])
            data_item['Vintage'] = row_values[14]
            data_item['Customer No.'] = row_values[15]
            data_item['Brand Code'] = row_values[16]
            data_item['Varietal Code'] = row_values[17]

            temp_list.append(data_item)
    return temp_list

def generate_clean_data_list(rdl):
    scrubbed_data_list = []
    clean_objdict = {}
    for raw_dict in rdl:
        clean_objdict['Item code w/o vintage'] = remove_vintage(raw_dict['Item No.'])
        clean_objdict['Brand Code'] = raw_dict['Brand Code']
        clean_objdict['']



# def generate_brand_lookup():
#     xlrd.

raw_data_list = build_raw_list()

    # for item in raw_data_list:
    #     if item[]

print ("the size of data_list: " + str(len(raw_data_list)))
# Serialize the list of dicts to JSON
j = json.dumps(raw_data_list, sort_keys=True, indent=4 * ' ')
#print(j)

# Write to file
with open('data.json', 'w') as f:
    f.write(j)

# Dictionary for scrubbed data

clean_data= {}

