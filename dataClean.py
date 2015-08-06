#! python 3
__author__ = 'Rowbot'

import openpyxl
import simplejson as json
import datetime as dt
import re

__libchange__ = False

size_dict = {'B': '750ml', 'E': '375ml', 'A': '750ml'}
region_spelling_dict = {'MIDATLANTI': 'Mid Atlantic',
                        'NEW ENGLAN': 'New England',
                        'TRI-STATES': 'Tri-States',
                        'CENTRAL': 'Central',
                        'OH VALLEY': 'Ohio Valley',
                        'TEXAS': 'Texas Region',
                        'CAROLINAS': 'Carolinas',
                        'FLORIDA': 'Florida',
                        'GULF STATE': 'Gulf States',
                        'MOUNTAIN': 'Mountain',
                        'SOUTHWEST': 'Southwest',
                        'IN': 'International'
                        }
innovation_east = ['Mid Atlantic', 'New England', 'Tri-States', 'Carolinas', 'Florida', 'Gulf States']

uselesslist = ['Item code w/o vintage',	'Brand Code',	'Brand',	'Varietal Code',	'Varietal', 	'Distributor',	'State',	'Sales Rep',	'Item ID',	'Item',	'SKU Tag',	'Item Pre',	'Size',	'Month',	'Year',	'Date',	'Document Type',	'Warehouse',	'Document #',	'Opposite #',	'Sales $',	'Total Cases',	'Vintage',	'Portfolio',	'Category',	'Sales/Key Acct Rep',	'ISM',	'IBM',	'Customer ID',	'Sales FOB',	'SKU Cost',	'SKU DA',	'Total DA$',	'GP$/CASE',	'Total GP$']

new_items = []


def remove_vintage(n):
    return n[0:-3] + n[-1]

def excel_date(date1):
    temp = dt.datetime(1899, 12, 30)
    delta = date1 - temp
    return float(delta.days)

def build_regionrep_lookup_dict():
    wb = openpyxl.load_workbook('Lookup Tables/Lookup Tables.xlsx')
    ws = wb.get_sheet_by_name('Region Rep')
    temp_dict = {}
    for rownum in range(2, ws.get_highest_row()+1):
        #TODO find out if the above line is including or overshooting the last row
        temp_dict[ws.cell(row=rownum, column=1).value] = ws.cell(row=rownum, column=2).value
    return temp_dict


def prettyprint(d):
    print(json.dumps(d, sort_keys=True, indent=4 * ' '))


#reads in raw excel data and output a list of dictionary objects for each row
def build_raw_list(sh):
    # Iterate through each row in worksheet and fetch values into dict
    i = 0
    temp_list = []
    row_tuples = sh.rows
    for rt in row_tuples[2:]:
        #TODO create a third-party vendor list for items like "Kirkland"
        if rt[2].value != 'PAC' and "CONSUMER" not in rt[0].value and "Kirkland" not in rt[4].value and "zBarter" not in rt[0].value and rt[2].value != "" and 'wine country connect' not in rt[0].value.lower():
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

def fix_sales_rep(obj_dict, salesrep_by_region, naluai_brand_list,
                  gangle_brand_list, swift_brand_list, jackson_brand_list,
                  owens_brand_list, nelson_brand_list):
    region = obj_dict['Sales Rep']
    brand = obj_dict['Brand']
    category = obj_dict['Category']
    if region in salesrep_by_region:
        obj_dict['Sales/Key Acct Rep'] = salesrep_by_region[region]
    if brand in jackson_brand_list:
        print('Jackson')
        obj_dict['Sales/Key Acct Rep'] = 'Jackson'
    if 'innovation' in region.lower() and brand in owens_brand_list:
        obj_dict['Sales/Key Acct Rep'] = 'Owens'
    if brand in gangle_brand_list:
        print('Gangel')
        obj_dict['Sales/Key Acct Rep'] = 'Gangel'
    if brand in naluai_brand_list:
        print('Naluai')
        obj_dict['Sales/Key Acct Rep'] = 'Naluai'
    if brand in swift_brand_list:
        obj_dict['Sales/Key Acct Rep'] = 'Swift'
    if brand in nelson_brand_list:
        obj_dict['Sales/Key Acct Rep'] = 'Nelson'
    if category == 'Total Wine' or 'total wine' in obj_dict['Distributor'].lower():
        obj_dict['Sales/Key Acct Rep'] = 'Bukoskey'
    return obj_dict



def generate_clean_data_list(rdl):
    scrubbed_data_list = []
    brand_dict = json.load(open('JSON Files/brandlookup.json', 'r'))
    state_by_distributor_dict = json.load(open('JSON Files/state_by_distributor.json', 'r'))
    region_by_state_dict = json.load(open('JSON Files/regionbystate.json'))
    sales_rep_by_region = json.load(open('JSON Files/regionrepbyregion.json'))
    varietal_by_code = json.load(open('JSON Files/varietalbycode.json'))
    naluai_brand_list = json.load(open('JSON Files/naluai_brand_list.json'))
    gangel_brand_list = json.load(open('JSON Files/gangel_brand_list.json'))
    swift_brand_list = json.load(open('JSON Files/swift_brand_list.json'))
    jackson_brand_list = json.load(open('JSON Files/jackson_brand_list.json'))
    owens_brand_list = json.load(open('JSON Files/owens_brand_list.json'))
    nelson_brand_list = json.load(open('Json Files/nelson_brand_list.json'))

    for raw_dict in rdl:
        clean_objdict = {}
        #TODO add if statements to control what happens when itemcode is not in the brand_dict
        clean_objdict['Item code w/o vintage'] = remove_vintage(raw_dict['Item No.'])
        clean_objdict['Brand Code'] = raw_dict['Item No.'][:3]
        clean_objdict['Distributor'] = raw_dict['Customer Name']
        clean_objdict['State'] = raw_dict['Ship-to State']
        clean_objdict['Varietal Code'] = var = raw_dict['Item No.'][3:6]
        if var in varietal_by_code:
            clean_objdict['Varietal'] = varietal_by_code[clean_objdict['Varietal Code']]
        else:
            clean_objdict['Varietal'] = raw_dict['Varietal Code']
        # TODO this may need to be drawn from the region by state dictionary, testing req'd
        if raw_dict['Ship-to State'] != '':
            clean_objdict['Sales Rep'] = region_by_state_dict[raw_dict['Ship-to State']]
        else:
            clean_objdict['Sales Rep'] = raw_dict['Salesperson Code']
        clean_objdict['Item ID'] = itemid = raw_dict['Item No.']
        clean_objdict['Item'] = raw_dict['Description']
        #TODO this SKU Tag assignment doesn't follow the formula, need to find out what that is.
        if 'f8' not in clean_objdict['Item ID'].lower():
            clean_objdict['Item Pre'] = raw_dict['Item No.'][:9]
        else:
            clean_objdict['Item Pre'] = raw_dict['Item No.'][:8]
        clean_objdict['Size'] = itemid[-1]
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
        clean_objdict['Customer ID'] = distributor = raw_dict['Customer No.']
        clean_objdict['Sales FOB'] = raw_dict['Sales Amount (Actual)'] / raw_dict['Quantity (positive)']
        clean_objdict['SKU DA'] = ''
        clean_objdict['Total DA$'] = ''
        clean_objdict['GP$/CASE'] = ''
        clean_objdict['Total GP$'] = ''

        # brand dicionary dependent items
        if itemid in brand_dict:
            if distributor in brand_dict[itemid]:
                #if the brand item and distributor relationship are the same
                #copy all the stuff over because its the same
                clean_objdict['Brand'] = brand_dict[itemid][distributor]['Brand']
                clean_objdict['SKU Tag'] = brand_dict[itemid][distributor]['SKU Tag']
                clean_objdict['Portfolio'] = brand_dict[itemid][distributor]['Portfolio']
                clean_objdict['Category'] = brand_dict[itemid][distributor]['Category']
                clean_objdict['Sales/Key Acct Rep'] = brand_dict[itemid][distributor]['Sales/Key Acct Rep']
                clean_objdict['ISM'] = brand_dict[itemid][distributor]['ISM']
                clean_objdict['IBM'] = brand_dict[itemid][distributor]['IBM']
                clean_objdict['SKU Cost'] = brand_dict[itemid][distributor]['SKU Cost']
                clean_objdict['Sales Rep'] = brand_dict[itemid][distributor]['Sales Rep']
                clean_objdict['SKU DA'] = brand_dict[itemid][distributor]['SKU DA']

            else:
                # pick the first item out of the dictionary and use it's values to take
                # advantage of the fact that they are of the same brand and vintage
                # region, category, salesperson take figuring out
                #NOTE: I may be able to create a lookup table for distributors name: {region: ####, salesrep: #####, etc...}
                reference_item = brand_dict[itemid][list(brand_dict[itemid])[0]]
                clean_objdict['Brand'] = reference_item['Brand']
                clean_objdict['SKU Tag'] = reference_item['SKU Tag']
                clean_objdict['Portfolio'] = reference_item['Portfolio']
                clean_objdict['Category'] = reference_item['Category']
                # if the region is in the rep by region lookup, associate the lineitem with lookup value
                # if clean_objdict['Sales Rep'] in sales_rep_by_region:
                #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[clean_objdict['Sales Rep']]
                # else:
                #     clean_objdict['Sales/Key Acct Rep'] = "#N/A"
                clean_objdict = fix_sales_rep(clean_objdict, sales_rep_by_region, naluai_brand_list, gangel_brand_list, swift_brand_list, jackson_brand_list, owens_brand_list, nelson_brand_list)
                #TODO figure out a more elegent solution
                # the code below appears to be duplicate, but is in fact applying spelling corrections to the region stuff UGLY CODE :(
                # if raw_dict['Salesperson Code'] in region_spelling_dict:
                #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_spelling_dict[raw_dict['Salesperson Code']]]
                # else:
                #     state = raw_dict['Ship-to State']
                #     if state is not "":
                #         clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_by_state_dict[raw_dict['Ship-to State']]]
                #     else:
                #         clean_objdict['Sales/Key Acct Rep'] = "#N/A"
                clean_objdict['ISM'] = reference_item['ISM']
                clean_objdict['IBM'] = reference_item['IBM']
                clean_objdict['SKU Cost'] = reference_item['SKU Cost']
        #otherwise the itemcode is not yet in the database
        else:
            new_items.append(itemid)
            print(itemid, ": new code")
            check_fail = True
            standby_brand = "#N/A"
            continue_searching = True
            for key in brand_dict.keys():
                if itemid[:3] == key[:3] and continue_searching:
                    standby_brand = brand_dict[key][list(brand_dict[key])[0]]['Brand']
                    continue_searching = False
                if itemid[:8] in key:
                    if distributor in brand_dict[key]:
                        clean_objdict['Brand'] = brand_dict[key][distributor]['Brand']
                        clean_objdict['SKU Tag'] = brand_dict[key][distributor]['SKU Tag']
                        clean_objdict['Portfolio'] = brand_dict[key][distributor]['Portfolio']
                        clean_objdict['Category'] = brand_dict[key][distributor]['Category']
                        clean_objdict['Sales/Key Acct Rep'] = brand_dict[key][distributor]['Sales/Key Acct Rep']
                        clean_objdict['ISM'] = brand_dict[key][distributor]['ISM']
                        clean_objdict['IBM'] = brand_dict[key][distributor]['IBM']
                        clean_objdict['SKU Cost'] = brand_dict[key][distributor]['SKU Cost']
                        check_fail = False
                        break
                    else:
                        reference_item = brand_dict[key][list(brand_dict[key])[0]]
                        clean_objdict['Brand'] = reference_item['Brand']
                        clean_objdict['SKU Tag'] = '#N/A'
                        clean_objdict['Portfolio'] = '#N/A'
                        clean_objdict['Category'] = '#N/A'
                        # if raw_dict['Salesperson Code'] in region_spelling_dict:
                        #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_spelling_dict[raw_dict['Salesperson Code']]]
                        # else:
                        #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_by_state_dict[raw_dict['Ship-to State']]]
                        clean_objdict = fix_sales_rep(clean_objdict, sales_rep_by_region, naluai_brand_list,
                                      gangel_brand_list, swift_brand_list, jackson_brand_list,
                                      owens_brand_list, nelson_brand_list)
                        clean_objdict['ISM'] = '#N/A'
                        clean_objdict['IBM'] = '#N/A'
                        clean_objdict['SKU Cost'] = '#N/A'
                        check_fail = False
                        break
            if check_fail:
                        clean_objdict['Brand'] = standby_brand
                        #TODO create a size lookup library
                        clean_objdict['SKU Tag'] = standby_brand + clean_objdict['Varietal'] + clean_objdict['Size']
                        clean_objdict['Portfolio'] = '#N/A'
                        clean_objdict['Category'] = '#N/A'
                        clean_objdict = fix_sales_rep(clean_objdict, sales_rep_by_region, naluai_brand_list,
                                      gangel_brand_list, swift_brand_list, jackson_brand_list,
                                      owens_brand_list, nelson_brand_list)
                        # if raw_dict['Salesperson Code'] in region_spelling_dict:
                        #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_spelling_dict[raw_dict['Salesperson Code']]]
                        # else:
                        #     clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_by_state_dict[raw_dict['Ship-to State']]]
                        clean_objdict['ISM'] = '#N/A'
                        clean_objdict['IBM'] = '#N/A'
                        clean_objdict['SKU Cost'] = '#N/A'

        # if 'Sales/Key Acct Rep' not in clean_objdict:
        #     if raw_dict['Salesperson Code'] in region_spelling_dict:
        #             clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_spelling_dict[raw_dict['Salesperson Code']]]
        #     else:
        #         state = raw_dict['Ship-to State']
        #         if state is not "":
        #             clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_by_state_dict[raw_dict['Ship-to State']]]
        #         else:
        #             clean_objdict['Sales/Key Acct Rep'] = "#N/A"

        #Initiate data scrub
        clean_objdict = refine_item_data(clean_objdict,
                                         region_by_state_dict,
                                         state_by_distributor_dict)
        if 'Sales/Key Acct Rep' not in clean_objdict:
            clean_objdict = fix_sales_rep(clean_objdict, sales_rep_by_region, naluai_brand_list,
                                          gangel_brand_list, swift_brand_list, jackson_brand_list,
                                          owens_brand_list, nelson_brand_list)
            if 'Sales/Key Acct Rep' not in clean_objdict:
                if raw_dict['Salesperson Code'] in region_spelling_dict:
                    clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_spelling_dict[raw_dict['Salesperson Code']]]
                else:
                    state = raw_dict['Ship-to State']
                    if state is not "":
                        clean_objdict['Sales/Key Acct Rep'] = sales_rep_by_region[region_by_state_dict[raw_dict['Ship-to State']]]
                    else:
                        clean_objdict['Sales/Key Acct Rep'] = "#N/A"

        scrubbed_data_list.append(clean_objdict)
    return scrubbed_data_list



def refine_item_data(obj_dict, region_by_state_dict, state_by_distributor_dict):
    if 'Portfolio' in obj_dict:
        pf = obj_dict['Portfolio'].lower()
    else:
        pf = ""
        print("Item did no have a portfolio field")
        prettyprint(obj_dict)

    #switch monterey area items to 'Southwest' region
    if 'monterey' in obj_dict['Distributor'].lower():
        obj_dict['Sales Rep'] = 'Southwest'
        print('switched monterey to SW')

    #fix canada region - Step 11
    if 'canada' in obj_dict['Sales Rep'].lower():
        obj_dict['State'] = 'Canada'
        if obj_dict['Sales Rep'].lower()[0] is 'e':
            obj_dict['Sales Rep'] = 'East Canada'
        else:
            obj_dict['Sales Rep'] = 'West Canada'

    #fix international obj_dict['Sales Rep'].lower() - Step 12
    if obj_dict['Sales Rep'].lower() == 'in':
        obj_dict['Sales Rep'] = 'International'
        #cover the instance where there is a new distributor
        if obj_dict['Distributor'] in state_by_distributor_dict:
            obj_dict['State'] = state_by_distributor_dict[obj_dict['Distributor']]

    #fix AA region - Step 13
    if 'alaska airlines' in obj_dict['Distributor'].lower():
        obj_dict['Sales Rep'] = 'Airlines'

    #fix precept house regions - Step 14
    #if the item is in the 'PH' region...
    if 'ph' in obj_dict['Sales Rep'].lower():
        # and it is Core or V&E
        if 'core' in pf or 'v&e' in pf:
            # and it is a direct shipment account...
            if 'sales shipment' in obj_dict['Document Type'].lower():
                #make the region 'Precept House'
                obj_dict['Sales Rep'] = 'Precept House'
            # elif obj_dict['Brand'] in innovation_by_brand_dict:
            #     obj_dict['Sales Rep'] = 'Innovation'
            else:
                #otherwise make it the correct region according to the state
                obj_dict['Sales Rep'] = region_by_state_dict[obj_dict['State']]

    #fix TOTAL WINE region - Step 15
    if 'total wine' in obj_dict['Sales Rep'].lower():
        obj_dict['Sales Rep'] = 'Total Wine'

    #fix region spelling - Step 16
    if obj_dict['Sales Rep'] in region_spelling_dict:
        obj_dict['Sales Rep'] = region_spelling_dict[obj_dict['Sales Rep']]

    #fix Glazer's distributor title - Step 18
    if 'Glazer' in obj_dict['Distributor'] and 'Texas' in obj_dict['Distributor']:
        obj_dict['Distributor'] = 'Glazer\'s of Texas'

    #fix Odom Name and Region - Step 19
    if 'Odom Corporation - Alaska' in obj_dict['Distributor']:
        obj_dict['State'] = 'Alaska'
        obj_dict['Sales Rep'] = "Mountain"
    if 'Odom Corporation - Cour D\'Alen' in obj_dict['Distributor'] or 'Odom Corporation - Lewiston' in obj_dict['Distributor']:
        obj_dict['State'] = 'ID'
        obj_dict['Sales Rep'] = 'NW WA'
    #fix HWB - Step 31
    if 'house wine box' in obj_dict['SKU Tag'].lower():
        obj_dict['Brand'] = 'House Wine Box'
        obj_dict['Category'] = '3L BIB'
    if 'house wind lone star' in obj_dict['SKU Tag'].lower():
        obj_dict['Brand'] = 'House Wine Lone Star'
        obj_dict['Category'] = 'Innovation'
    if 'ste chapelle box' in obj_dict['SKU Tag'].lower():
        obj_dict['Brand'] = 'Ste Chappelle Box'
        obj_dict['Category'] = '3L BIB'
    if 'wtso' in obj_dict['SKU Tag'].lower():
        obj_dict['Brand'] = 'WTSO'
        obj_dict['Category'] = 'Innovation'
        #Fix Grape and grain portfolio items - Step 36
    if 'grape & grain' in pf:
        if 'total wine' not in obj_dict['Sales Rep'].lower() and 'airlines' not in obj_dict['Sales Rep'].lower() and 'precept house' not in obj_dict['Sales Rep'].lower():
            if obj_dict['Sales Rep'] in innovation_east:
                obj_dict['Sales Rep'] = 'Innovation East'
            else:
                obj_dict['Sales Rep'] = 'Innovation West'
    if 'airlines' in obj_dict['Sales Rep'].lower():
        obj_dict['Category'] = 'Alaska Airlines'
    if 'canada' in obj_dict['Sales Rep'].lower():
        obj_dict['Category'] = 'Canada'
    if 'international' in obj_dict['Sales Rep'].lower():
        obj_dict['Category'] = 'International'
    if 'total wine' in obj_dict['Sales Rep'].lower():
        if 'red knot' in obj_dict['Brand'].lower():
            obj_dict['Brand'] = 'Red Knot TWM'
        if 'shingleback' in obj_dict['Brand'].lower():
            obj_dict['Brand'] = 'Shingleback TWM'
        if 'apex' in obj_dict['Brand'].lower():
            obj_dict['Brand'] = 'Apex TWM'
    if 'grocery outlet' in obj_dict['Distributor'].lower():
        obj_dict['Category'] = 'Closeout'
    if 'gruet' in obj_dict['Brand'].lower():
        obj_dict['Category'] = 'Gruet'
    if 'dsv' in obj_dict['Brand'].lower():
        obj_dict['Category'] = 'Gruet'
    if 'closeout' in pf:
        obj_dict['Category'] = 'Closeout'
    if 'core' in pf:
        obj_dict['Category'] = 'Core'
    if 'v&e' in pf:
        obj_dict['Category'] = 'V&E'
    if 'alaska airlines' in obj_dict['Distributor'].lower():
        if 'canoe ridge' in obj_dict['Brand'].lower():
            obj_dict['Brand'] = 'Canoe Ridge Exploration'
    return obj_dict


def write_to_excel(filename, datalist):
    wb = openpyxl.Workbook()
    ws = wb.get_active_sheet()
    rownum = 2
    # fill in the column header row
    for i in range(0, len(uselesslist)):
        ws.cell(row=1, column=i+1).value = uselesslist[i]
    for item in datalist:
        try:
            for i in range(0, len(uselesslist)):
                ws.cell(row=rownum, column=i+1).value = item[uselesslist[i]]
        except KeyError as ke:
            print("IT was himmm>>> " +ws.cell(row=rownum, column = 1).value)


        rownum +=1
    wb.save(filename)


def main():

    wbtest = openpyxl.load_workbook('Input/Item Ledger Precept - 7.21.15 - Copy.xlsx')
    # wbtest = openpyxl.load_workbook('Input/Item Ledger Precept - 7.21.15.xlsx')
    shtest = wbtest.get_active_sheet()

    raw_data_list = build_raw_list(shtest)

    write_to_excel('Cleaned data file/megatest16.xlsx', generate_clean_data_list(raw_data_list))


if __name__ == '__main__':
    main()