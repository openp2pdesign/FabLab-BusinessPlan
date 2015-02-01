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
money_style = workbook.add_format({'num_format': u'€#,##0'})
# Add green/red color for positive/negative numbers
#money_style.set_num_format('[Green]General;[Red]-General;General')
# Add a number format for cells with money 
#money_style.set_num_format('0 "dollar and" .00 "cents"')


# Add content -------------------------------------------------------------

# Add content to the Expenses worksheet
expenses.write('A1', 'Hello world', bold_style)
expenses.write('A2', '12.33', money_style)
expenses.write('A3', 'Total', total_style)


# Save document -------------------------------------------------------------

# Save and close the file
workbook.close()