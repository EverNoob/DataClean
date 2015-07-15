__author__ = 'Rowbot'

import openpyxl
import simplejson as json

def build_brand_lookup():

    wb = openpyxl.load_workbook('Lookup Tables/2014 and YTD 2015 Shipment Data File.xlsx')
    ws = wb.get_sheet_by_name('Sheet1')
    brand_dict = {}
    for rownum in range(2, ws.get_highest_row()+1):
        brand_dict[ws.cell(row=rownum, column=1).value] = {'Brand Code': ws.cell(row=rownum, column=2).value, 'Brand': ws.cell(row=rownum, column=3).value,
                                               'Varietal Code': ws.cell(row=rownum, column=4).value, 'Varietal': ws.cell(row=rownum, column=5).value,
                                               'Portfolio': ws.cell(row=rownum, column=24).value, 'Category': ws.cell(row=rownum, column=25).value}
    with open('JSON Files/brandlookup.json', 'w') as f:
        f.write(json.dumps([brand_dict], sort_keys=True, indent=4 * ' '))



build_brand_lookup()



