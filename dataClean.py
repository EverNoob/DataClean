#! python 3
__author__ = 'Rowbot'
import xlrd
from collections import OrderedDict
import simplejson as json

#objects related to lookup sheet functions
lookup_sheet_names = ['NAV Region', 'Region Rep', 'Brand Code',
                      'Varietals Code', 'Brand Associations', 'Brand Name Changes']
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

def build_brandassoc_lookup(ws_brandassoc):
    baj = build_json_from_ws(ws_brandassoc)
    temp_dict = {}
    for item in baj:
        temp_dict[item['Brand Code']] = item
    return temp_dict

def prettyprint(d):
    print(json.dumps(d, sort_keys=True, indent=4 * ' '))


def build_lookup_dict():
    temp_dict = {}
    #open lookup workbook and grab worksheets
    wb_lookup = xlrd.open_workbook('Lookup Tables/Lookup Tables.xlsx')
    temp_dict['NAV Region'] = build_twoitem_lookup_dict(wb_lookup.sheet_by_name('NAV Region'))
    temp_dict['Region Rep'] = build_twoitem_lookup_dict(wb_lookup.sheet_by_name('Region Rep'))
    temp_dict['Brand Code'] = build_twoitem_lookup_dict(wb_lookup.sheet_by_name('Brand Code'), reverse=True)
    temp_dict['Varietals Code'] = build_twoitem_lookup_dict(wb_lookup.sheet_by_name('Varietals Code'))
    temp_dict['Brand Associations'] = build_brandassoc_lookup(wb_lookup.sheet_by_name('Brand Associations'))
    temp_dict['Brand Name Changes'] = build_twoitem_lookup_dict(wb_lookup.sheet_by_name('Brand Name Changes'))
    return temp_dict

def build_lookup_json():
    return json.dumps([build_lookup_dict()], sort_keys=True, indent=4 * ' ')

def build_raw_list(sh):
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
    print(i)
    return temp_list

def generate_clean_data_list(rdl, lud):
    scrubbed_data_list = []
    for raw_dict in rdl:
        clean_objdict = {}
        clean_objdict['Item code w/o vintage'] = remove_vintage(raw_dict['Item No.'])
        clean_objdict['Brand Code'] = raw_dict['Brand Code']
        clean_objdict['Brand'] = lud["Brand Associations"][raw_dict['Brand Code']]
        scrubbed_data_list.append(clean_objdict)
    return scrubbed_data_list

def main():

    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook('Input/Dirty Data.xlsx')
    sh = wb.sheet_by_index(0)

    raw_data_list = build_raw_list(sh)

    print ("the size of data_list: " + str(len(raw_data_list)))
    # Serialize the list of dicts to JSON
    j = json.dumps(raw_data_list, sort_keys=True, indent=4 * ' ')
    L = build_lookup_json()
    lud = build_lookup_dict()
    #print(j)

    # Write to file
    with open('data.json', 'w') as f:
        f.write(j)
        f.close()

    with open('lookup.json','w') as f:
        f.write(L)
        f.close()

    prettyprint(generate_clean_data_list(raw_data_list, lud))

    # Dictionary for scrubbed data

    clean_data= {}

if __name__ == '__main__':
    main()