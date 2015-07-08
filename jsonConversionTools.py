__author__ = 'Rowbot'

import simplejson as json
import openpyxl

def json_to_csv(jsonfilepath, outfilepath):
   json_dict = json.load(jsonfilepath)
   outfile = open(outfilepath,mode='w')
   column_headers = list(json_dict[0].keys())
   for header in column_headers:
       outfile.write(str(header)+',')




def main():
    return