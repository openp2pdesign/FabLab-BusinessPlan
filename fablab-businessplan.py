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
expenses = workbook.add_worksheet('Expenses')
activities = workbook.add_worksheet('Activities')
membership = workbook.add_worksheet('Membership')
total = workbook.add_worksheet('Total')


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
expenses.set_column('A:A', 10)

# Add content to the Expenses worksheet
expenses.write('A1', 'Hello world', bold_style)
expenses.write('A2', item_value, money_style)
expenses.write('A3', item_value, money_style)
expenses.write('A4', '=SUM(A2:A3)', money_total_style)


# Save document -------------------------------------------------------------

# Save and close the file
workbook.close()

# Open the file in OSX, for debug
import subprocess
import os
where = os.path.dirname(os.path.abspath(__file__))
subprocess.check_output(["./soffice",where+"/FabLab-BusinessPlan.xlsx"], cwd="/Applications/LibreOffice.app/Contents/MacOS")