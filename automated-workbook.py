
# Packages

# A package is a container for multiple modules. In file system terms, a package is a directory or folder.

# Packages extremely important for a framework like Django; used for building web applications with Python.


# Files & Directories 

# You can iterate over all the spreadsheets in a directory, open them and process them.

 # Absolute path
 # c:\Program Files\Microsoft
 # usr/local/bin

# Relative path

# from pathlib import Path

# Path instantiates a concrete Path for the platform the code is running on.

# print(path.rmdir()) 

# Current directory 
# path = Path()

# String defines search pattern 
# path = Path()
# print(path.glob(''))

# All files, all directories
# path = Path()
# print(path.glob('*'))

# All files only
# path = Path()
# print(path.glob('*.*'))

# All the excel spread sheets 
# path = Path()
# print(path.glob('*.xls'))

# All the python files 
# path = Path()
# print(path.glob('*.py'))

# We can iterate over generator objects.
# path = Path()
# for file in path.glob('*'):
#     print(file)


# Pypi & Pip

# The Python Package Index (PyPi) is the default, vast software repository of open-source Python packages supplied by the Worldwide community of Python developers. 

# Pip manager is a package manager for installing external modules or programs.

# Web scraping. Building an engine - a web crawler - and browsing websites, extracting information from HTML files. Same technique Google uses to index websites for its search engine. When user publishes a blog post, Google's search engine extracts it's headline, keywords etc. 

# Browser automation. Extremely powerful coz you can automate testing of your web applications. PyPi package Selenium.

# Excel Spreadsheets

# Python program can automate thousands of excel spreadsheets in under a second. Can be built in less than 30 minutes.

# Openpyxl - package for working with excel spreadsheets.

# 1. Import openpyxl 2. Load Workbook 3. Iterating over rows 4. Fixing prices         5. Selecting values to add BarChart 6. Saving Workbook 

# Automated Workbook (code)

import openpyxl as xl
from openpyxl.chart import BarChart, Reference 

wb = xl.load_workbook('transactions.xlsx')
sheet = wb['Sheet1']

for row in range(2, sheet.max_row + 1):
    cell = sheet.cell(row, 3)
    corrected_price = cell.value * 0.9
    corrected_price_cell = sheet.cell(row, 4)
    corrected_price_cell.value = corrected_price

values = Reference(sheet,
         min_row=2, 
         max_row=sheet.max_row,
         min_col=4,
         max_col=4)

chart = BarChart()
chart.add_data(values)
sheet.add_chart(chart, 'e2')

wb.save('transactions2.xlsx')
    
# Automated Workbook Function

import openpyxl as xl
from openpyxl.chart import BarChart, Reference 

def process_workbook(filename):
    wb = xl.load_workbook(filename)
    sheet = wb['Sheet1']

    for row in range(2, sheet.max_row + 1):
        cell = sheet.cell(row, 3)
        corrected_price = cell.value * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price

    values = Reference(sheet,
            min_row=2, 
            max_row=sheet.max_row,
            min_col=4,
            max_col=4)

    chart = BarChart()
    chart.add_data(values)
    sheet.add_chart(chart, 'e2')

    wb.save(filename)

# We’ve created a bar chart by calling openpyxl.chart.BarChart(). 
# You can also create line charts, scatter charts, and pie charts by calling openpyxl.charts.LineChart(), openpyxl.chart.ScatterChart(), and openpyxl.chart.PieChart().

