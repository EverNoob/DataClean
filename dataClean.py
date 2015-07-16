#! python 3
__author__ = 'Rowbot'
import xlrd
import openpyxl
import simplejson as json
import datetime as dt

__libchange__ = False

size_dict = {'B': '750ml', 'E': '375ml', 'A': '750ml'}


def remove_vintage(n):
    return n[0:-3] + n[-1]

def excel_date(date1):
    temp = dt.datetime(1899, 12, 30)
    delta = date1 - temp
    return float(delta.days)


# def build_lookup_tables():
#     look_wb = xlrd.open_workbook('Lookup Tables/Lookup Tables.xlsx')
#     varietal_dict = build_lookup_dict(look_wb,'Varietals Code',1)
#     #distributor_dict = build_lookup_dict(look_wb,'')
#     return


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
        rt = ws_navregion.rt(rownum)
        temp_dict[rt[0]] = rt[1]
    return temp_dict


def build_regionrep_lookup_dict(ws_regionrep):
    temp_dict = {}
    for rownum in range(1, ws_regionrep.nrows):
        rt = ws_regionrep.rt(rownum)
        temp_dict[rt[0]] = rt[1]
    return temp_dict


def build_twoitem_lookup_dict(ws,reverse=False):
    temp_dict = {}
    for num in range(1, ws.nrows):
        row_values = ws.row_values(num)
        if reverse:
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

#reads in raw excel data and output a list of dictionary objects for each row
def build_raw_list(sh):
    global i, temp_list, rownum, rt, data_item
    # Iterate through each row in worksheet and fetch values into dict
    # for rownum in range(2, sh.nrows):
    i = 0
    temp_list = []
    row_tuples = sh.rows
    for rt in row_tuples[2:]:
        #TODO create a third-party vendor list for items like "Kirkland"
        if rt[2].value != 'PAC' and "CONSUMER" not in rt[0].value and "Kirkland" not in rt[4].value and "zBarter" not in rt[0].value and rt[2].value != "":
            data_item = {}
            i += 1
            data_item['Customer Name'] = rt[0].value
            data_item['Ship-to State'] = rt[1].value
            data_item['Salesperson Code'] = rt[2].value
            data_item['Item No.'] = rt[3].value
            data_item['Description'] = rt[4].value
            data_item['Product Group Code'] = rt[5].value
            data_item['Posting Month'] = rt[6].value
            data_item['Posting Date'] = excel_date(rt[7].value)
            data_item['Year'] = rt[7].value.year
            data_item['Document Type'] = rt[8].value
            data_item['Location Code'] = rt[9].value
            data_item['Document No.'] = rt[10].value
            data_item['Quantity'] = rt[11].value
            data_item['Sales Amount (Actual)'] = rt[12].value
            data_item['Quantity (positive)'] = rt[13].value
            data_item['Vintage'] = rt[14].value
            data_item['Customer No.'] = rt[15].value
            data_item['Brand Code'] = rt[16].value
            data_item['Varietal Code'] = rt[17].value
            temp_list.append(data_item)
    print('The number of raw items of interest is: ' + str(i))
    return temp_list


def build_brand_lookup():

    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    brand_dict = {}
    state_dict = {}

    for rownum in range(2, ws.get_highest_row()+1):
        brand_dict[ws.cell(row=rownum, column=1).value] = {'Brand Code': ws.cell(row=rownum, column=2).value,
                                                           'Brand': ws.cell(row=rownum, column=3).value,
                                                           'Varietal Code': ws.cell(row=rownum, column=4).value,
                                                           'Varietal': ws.cell(row=rownum, column=5).value,
                                                           'SKU TAG': ws.cell(row=rownum, column=11),
                                                           'Portfolio': ws.cell(row=rownum, column=24).value,
                                                           'Category': ws.cell(row=rownum, column=25).value,
                                                           'Sales/Key Acct Rep': ws.cell(row=rownum, column=26).value,
                                                           'ISM': ws.cell(row=rownum, column=27).value,
                                                           'IBM': ws.cell(row=rownum, column=28).value,
                                                           'SKU Cost': ws.cell(row=rownum, column=31).value
                                                           }
        state_dict[ws.cell(row=rownum, column=6).value] = ws.cell(row=rownum, column=7).value

    return brand_dict, state_dict





def change_canada_region(obj_dict, state_dict):
    sr = obj_dict['Sales Rep'].lower()
    pf = obj_dict['Portfolio'].lower()
    if 'canada' in sr:
        obj_dict['State'] = 'Canada'
        if sr[0] is 'e':
            obj_dict['Sales Rep'] = 'East Canada'
        else:
            obj_dict['Sales Rep'] = 'West Canada'
    if 'in' in sr:
        obj_dict['Sales Rep'] = 'International'
        #cover the instance where there is a new distributor
        if obj_dict['Distributor'] in state_dict:
            obj_dict['State'] = state_dict[obj_dict['Distributor']]
    if 'sales shipment' in obj_dict['Document Type'].lower() and ('core' in pf or 'v&e' in pf):
        obj_dict['Sales Rep'] = 'Precept House'

    return obj_dict


def generate_clean_data_list(rdl):
    scrubbed_data_list = []
    brand_dict, state_dict = build_brand_lookup()

    for raw_dict in rdl:
        clean_objdict = {}
        clean_objdict['Item code w/o vintage'] = itemcode = remove_vintage(raw_dict['Item No.'])
        clean_objdict['Brand Code'] = raw_dict['Brand Code']
        clean_objdict['Brand'] = brand_dict[itemcode]['Brand']
        clean_objdict['Varietal Code'] = brand_dict[itemcode]['Varietal Code']
        clean_objdict['Varietal'] = brand_dict[itemcode]['Varietal']
        clean_objdict['Distributor'] = raw_dict['Customer Name']
        clean_objdict['State'] = raw_dict['Ship-to State']
        clean_objdict['Sales Rep'] = raw_dict['Salesperson Code']
        clean_objdict['Item ID'] = raw_dict['Item No.']
        clean_objdict['Item'] = raw_dict['Description']
        clean_objdict['SKU Tag'] = brand_dict[itemcode]['SKU Tag']
        clean_objdict['Item Pre'] = raw_dict['Item No.'][:8]
        clean_objdict['Size'] = itemcode[-1]
        clean_objdict['Month'] = raw_dict['Posting Month']
        clean_objdict['Year'] = raw_dict['Year']
        clean_objdict['Date'] = raw_dict['Posting Date']
        clean_objdict['Document Type'] = raw_dict['Document Type']
        clean_objdict['Warehouse'] = raw_dict['Location Code']
        clean_objdict['Document #'] = raw_dict['Document No.']
        clean_objdict['Opposite #'] = raw_dict['Quantity']
        clean_objdict['Sales $'] = raw_dict['Sales Amount (Actual)']
        clean_objdict['Total Cases'] = raw_dict['Quantity (positive)']
        clean_objdict['Vintage'] = raw_dict['Vintage']
        clean_objdict['Portfolio'] = brand_dict[itemcode]['Portfolio']
        clean_objdict['Category'] = brand_dict[itemcode]['Category']
        clean_objdict['Sales/Key Acct Rep'] = brand_dict['Sales/Key Acct Rep']
        clean_objdict['ISM'] = brand_dict['ISM']
        clean_objdict['IBM'] = brand_dict['IBM']
        clean_objdict['Customer ID'] = raw_dict['Customer No.']
        clean_objdict['Sales FOB'] = raw_dict['Sales Amount (Actual)'] / raw_dict['Quantity (positive)']
        clean_objdict['SKU Cost'] = 'PlaceHolder'


        scrubbed_data_list.append(clean_objdict)
    return scrubbed_data_list


def main():

    # Open the workbook and select the first worksheet
    wb = xlrd.open_workbook('Input/Dirty Data.xlsx')
    sh = wb.sheet_by_index(0)

    wbtest = openpyxl.load_workbook('Input/Dirty Data.xlsx')
    shtest = wbtest.get_active_sheet()

    raw_data_list = build_raw_list(shtest)

    print ("the size of data_list: " + str(len(raw_data_list)))
    # Serialize the list of dicts to JSON
    j = json.dumps(raw_data_list, sort_keys=True, indent=4 * ' ')
    L = build_lookup_json()


    # Write to file
    with open('rawdata.json', 'w') as f:
        f.write(j)
        f.close()

    with open('lookup.json', 'w') as f:
        f.write(L)
        f.close()

    prettyprint(generate_clean_data_list(raw_data_list))



if __name__ == '__main__':
    main()