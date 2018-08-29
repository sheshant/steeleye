

import os
import xlrd
import logging
import json


def get_stats_xls_file(filename, output_json_file_name):
    try:
        book = xlrd.open_workbook(filename=filename)
    except Exception as e:
        logging.error('Failed to open and read file '+ str(e))
        return

    try:
        sheet = book.sheet_by_name('MICs List by CC')
    except (ValueError, XLRDError):
        logging.error('Failed to read tab "MICs List by CC" '+ str(e))
        return

    if sheet.nrows and sheet.ncols:
        json_list = []
        header = sheet.row_values(0)
        for row_idx in range(1, sheet.nrows):
            json_dict = {}
            for col_idx in range(0, sheet.ncols):
                json_dict[header[col_idx]] = sheet.cell(row_idx, col_idx).value

            json_list.append(json_dict)

        with open(output_json_file_name, 'w') as outfile:
            json.dump(json_list, outfile)

    else:
        logging.error('no records found in sheet')


        



if __name__ == '__main__':
    import argparse

    arg_parser = argparse.ArgumentParser(description='Function get_stats_xls_file will take one '
    'xls file name with path as input and a json file path to store output data. It will parse the '
    'xls file and display content of the tab "MICs List by CC" in a json list of dict and save it '
    'in a json file')
    arg_parser.add_argument('-f', '--xls_filename', 
                            type=str, help='filename with path for .xls file', required=True)
    arg_parser.add_argument('-j', '--output_json_filename', type=str, 
                            help='filename with path for the output json', required=True)
    xls_filename = arg_parser.parse_args().xls_filename
    output_json_filename = arg_parser.parse_args().output_json_filename
    get_stats_xls_file(xls_filename, output_json_filename)

