#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Explanation: Sample XLS report generator
Usage:
    $ python ms-xls-gen-apps [ options ]
Style:
    Google Python Style Guide:
    http://google.github.io/styleguide/pyguide.html
    @name           ms-xls-gen-apps
    @version        1.0.0
    @author-name    Wayne Schmidt
    @author-email   wschmidt@sumologic.com
    @license-name   GNU GPL
    @license-url    http://www.gnu.org/licenses/gpl.html
"""

__version__ = 1.0
__author__ = "Wayne Schmidt (wschmidt@sumologic.com)"

import argparse
import sys
import datetime
import io
import pandas
import pandas.io.formats.excel
import requests
import xlsxwriter

pandas.io.formats.excel.header_style = None
sys.dont_write_bytecode = 1

PARSER = argparse.ArgumentParser(description="""
Sample script to build a powerpoint from CSV items
""")

PARSER.add_argument('-c', metavar='<clientdata>', dest='clientdata', help='specify client data')
PARSER.add_argument('-d', metavar='<reportdata>', dest='reportdata', help='specify report data')
PARSER.add_argument('-o', metavar='<outputfile>', dest='outputfile', help='specify output file')

ARGS = PARSER.parse_args()

NOW = datetime.datetime.now()
TSTAMP = NOW.strftime("%B %-d, %Y")

SUMOURL = "https://logo.clearbit.com/www.sumologic.com"
SUMOIMG = io.BytesIO(requests.get(SUMOURL, stream = True).raw.read())

dataframe = pandas.read_csv(ARGS.clientdata)
dataframe.fillna(0, inplace=True)

transposed = dataframe.transpose()
transposed.reset_index(level=0, inplace=True)
my_rows = transposed.shape[0]

START = 0
OFFSET = 8
final = my_rows

workbook = xlsxwriter.Workbook(ARGS.outputfile)

worksheet = workbook.add_worksheet('Client_Info')
worksheet.set_column(0,10,30)

worksheet.insert_image('A1', SUMOURL, {'image_data': SUMOIMG})

cell_formatV = workbook.add_format()
cell_formatV.set_font_name('Calibri')
cell_formatV.set_font_size('14')
cell_formatV.set_align('left')
cell_formatV.set_align('vcenter')
cell_formatV.set_border()

cell_formatK = workbook.add_format()
cell_formatK.set_font_name('Calibri')
cell_formatK.set_font_size('14')
cell_formatK.set_align('left')
cell_formatK.set_align('vcenter')
cell_formatK.set_bg_color('#333399')
cell_formatK.set_font_color('white')
cell_formatK.set_border()

for my_row in range(START,my_rows,1):
    row = my_row + OFFSET
    key = transposed['index'][my_row]
    value = transposed[0][my_row]
    worksheet.write(row, 0, key, cell_formatK)
    worksheet.write(row, 1, value, cell_formatV)

worksheet = workbook.add_worksheet('AppFinder_Data')
worksheet.set_column(0,10,30)

dataframeApp = pandas.read_csv(ARGS.reportdata)

START = 0
COL = START

for (columnName, columnData) in dataframeApp.iteritems():
    ROW = START
    cell_value = columnName
    worksheet.write(ROW, COL, cell_value, cell_formatK)
    ROW += 1
    for cell_value in columnData.values:
        worksheet.write(ROW, COL, cell_value, cell_formatV)
        ROW += 1
    COL += 1

worksheet.autofilter(START, START, ROW - 1 , COL - 1)

worksheet = workbook.add_worksheet('AppFinder_Graph')
worksheet.set_column(0,10,30)

dataframeGraph = dataframeApp.groupby('category').count()['key'].reset_index(name="count")

dataframeGraph.columns = ["category", "count"]

START = 0
COL = START

for (columnName, columnData) in dataframeGraph.iteritems():
    ROW = START
    cell_value = columnName
    worksheet.write(ROW, COL, cell_value, cell_formatK)
    ROW += 1
    for cell_value in columnData.values:
        worksheet.write(ROW, COL, cell_value, cell_formatV)
        ROW += 1
    COL += 1

worksheet.autofilter(START, START, ROW - 1 , COL - 1)

pie_chart = workbook.add_chart({'type': 'pie'})

pie_chart.add_series({
    'name': 'App_Finder_Breakdown',
    'categories': '=AppFinder_Graph!$A$2:$A$' + str(ROW),
    'values':     '=AppFinder_Graph!$B$2:$B$' + str(ROW),
})

pie_chart.set_style(18)

worksheet.insert_chart('C2', pie_chart, {'x_offset': 5, 'y_offset': 5})

workbook.close()
