# -*- encoding: utf-8 -*-
#
# Author: Massimo Menichinelli
# Homepage: http://www.openp2pdesign.org
# License: MIT
#

import xlsxwriter

# Create the file
workbook = xlsxwriter.Workbook('FabLab-BusinessPlan.xlsx')

# Create the worksheets
expenses = workbook.add_worksheet('Expenses')
activities = workbook.add_worksheet('Activities')
membership = workbook.add_worksheet('Membership')
total = workbook.add_worksheet('Total')

# Add content to the Expenses worksheet
expenses.write('A1', 'Hello world')

# Save and close the file
workbook.close()