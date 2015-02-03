# -*- encoding: utf-8 -*-
#
# Author: Massimo Menichinelli
# Homepage: http://www.openp2pdesign.org
# License: MIT
#

import xlsxwriter

# Create document -------------------------------------------------------------

# Create the file
workbook = xlsxwriter.Workbook('FabLab-BusinessPlan.xlsx')

# Create the worksheets
executive = workbook.add_worksheet('Executive Summary')
company = workbook.add_worksheet('Company Summary')
market = workbook.add_worksheet('Market Analysis')

# Add a worksheet for each year
starting_year = 2015
years_number = 5
years = {}
for i in range(years_number):
    years[2015+i] = workbook.add_worksheet('Plan for '+str(2015+i))


# Create styles -------------------------------------------------------------

# Add a bold style to highlight heading cells
bold_style = workbook.add_format()
bold_style.set_font_color('white')
bold_style.set_bg_color('F56A2F')
bold_style.set_bold()

# Add a total style to highlight total cells
total_style = workbook.add_format()
total_style.set_font_color('red')
total_style.set_bg_color('FAECC5')
total_style.set_bold()

# Add a style for money
money_style = workbook.add_format()
money_style.set_num_format('$ 0.00')

# Add a style for money totals
money_total_style = workbook.add_format()
# Add green/red color for positive/negative numbers
money_total_style.set_num_format('[Green]$ 0.00;[Red]$ -0.00;$ 0.00')
total_style.set_bg_color('FAECC5')


# Add content -------------------------------------------------------------

# Dummy values
item_value = 0.00

# Set the width of the columns
executive.set_column('A:A', 10)

# Add content to the Expenses worksheet
executive.write('A1', 'Hello world', bold_style)
executive.write('A2', item_value, money_style)
executive.write('A3', item_value, money_style)
executive.write('A4', '=SUM(A2:A3)', money_total_style)

# Notes for the structure:

# Executive summary
# Introduction
# Bar chart for years
# Objectives
# Mission
# Keys to success

# Company summary
# Company structure/ownership
# Start-up summary
# Bar chart for start-up costs

# Market analysis
# Market segmentation
# Pie chart for market segmentation
# Target Market Segment Strategy

# Strategy
# Communication
# Years forecast
# Chart years forecast

# Management
# Persons
# Years forecast

# Financial plan

# Save document -------------------------------------------------------------

# Save and close the file
workbook.close()

# Open the file in OSX, for debug
import subprocess
import os
where = os.path.dirname(os.path.abspath(__file__))
subprocess.check_output(["./soffice",where+"/FabLab-BusinessPlan.xlsx"], cwd="/Applications/LibreOffice.app/Contents/MacOS")