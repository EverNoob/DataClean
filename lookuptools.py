__author__ = 'Rowbot'

import openpyxl
import simplejson as json

def build_brand_lookup():

    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    brand_dict = {}
    for rownum in range(2, ws.get_highest_row()+1):
        brand_dict[ws.cell(row=rownum, column=1).value] = {'Brand Code': ws.cell(row=rownum, column=2).value,
                                                           'Brand': ws.cell(row=rownum, column=3).value,
                                                           'Varietal Code': ws.cell(row=rownum, column=4).value,
                                                           'Varietal': ws.cell(row=rownum, column=5).value,
                                                           'SKU Tag': ws.cell(row=rownum, column=11).value,
                                                           'Portfolio': ws.cell(row=rownum, column=24).value,
                                                           'Category': ws.cell(row=rownum, column=25).value,
                                                           'Sales/Key Acct Rep': ws.cell(row=rownum, column=26).value,
                                                           'ISM': ws.cell(row=rownum, column=27).value,
                                                           'IBM': ws.cell(row=rownum, column=28).value,
                                                           'SKU Cost': ws.cell(row=rownum, column=31).value
                                                           }
    with open('JSON Files/brandlookup.json', 'w') as f:
        f.write(json.dumps(brand_dict, sort_keys=True, indent=4 * ' '))


def build_innovation_lookup():
    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    innovation_dict = {}
    for rownum in range(2, ws.get_highest_row()+1):
        #look for whether the brand in a "core brand" and not record the brand/innovation relationship
        if 'core' not in ws.cell(row=rownum, column=24).value.lower() and 'innovation' in ws.cell(row=rownum, column=8).value.lower():
            innovation_dict[ws.cell(row=rownum, column=3).value] = ws.cell(row=rownum, column=8).value
    with open('JSON Files/innovation_brands.json', 'w') as f:
        f.write(json.dumps(innovation_dict, sort_keys=True, indent= 4 * ' '))

def build_state_by_distributor_lookup():
    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    state_by_distributor_dict = {}
    #TODO: account for the distributors that operate in multiple states,
    #  perhaps by making each dist a ref. to a list
    for rownum in range(2, ws.get_highest_row()+1):
        state_by_distributor_dict[ws.cell(row=rownum, column=6).value] = ws.cell(row=rownum, column=7).value
    with open('JSON Files/state_by_distributor.json', 'w') as f:
        f.write(json.dumps(state_by_distributor_dict, sort_keys=True, indent= 4 * ' '))

def build_region_by_state_lookup():
    #TODO automate the process of eliminating all other regions, stub list below
    regions_to_delete = ['Airlines', 'California Region']
    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    region_by_state_dict = {}
    for rownum in range(2, ws.get_highest_row()+1):
        region = ws.cell(row=rownum, column=8).value
        state = ws.cell(row=rownum, column=7).value
        if region in region_by_state_dict:
            if state not in region_by_state_dict[region]:
                region_by_state_dict[region].append(state)
        else:
            region_by_state_dict[region] = [state]
    return region_by_state_dict
    with open('JSON Files/regionbystatelookup.json', 'w') as f:
        f.write(json.dumps(region_by_state_dict, sort_keys=True, indent=4 * ' '))



build_brand_lookup()


