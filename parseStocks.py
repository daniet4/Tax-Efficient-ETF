# This program parses Stock Screener data from Fidelity located at https://research2.fidelity.com/pi/stock-screener
# The goal is to determine the best stocks to hold in a taxable brokerage account to minimize dividend yields.
# It is desirable to concentrate investment returns in the stock price to avoid capital gains taxes.
# For tax efficiency, the remaining high dividend yielding stocks will be held in a tax advantaged account such as a
# 401k, IRA, or HSA.

# Updated: 8/15/18


import os
import glob

from bs4 import BeautifulSoup
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import win32com.client as win32

"""
0. Grab Fidelity Stock Screener data or something similar.
1. [DONE] Parse it and convert $X.XXB [str] - > X.XXe9 [float].
2. [DONE] Sort by dividend yield.
3. [DONE] Determine market capitalization percentage.
4. [DONE-ish] -- Printing to separate file for testing | Print to original file.
[4a]. Use the xlsxwriter module to improve the .xlsx file formatting. Refer to: http://pbpython.com/improve-pandas-excel-output.html
[5]. [DONE-ish] -- Add a total dividend yield | Print a separate file that meets a dividend yield goal; e.g. all stocks that have a dividend yield of less than 0.5%
    sorted from least to most dividend yield.
[5a]. Set a cut-off for the getGuide(div, >cutoff<) such that stocks are excluded if the optimal holding is less than a certain percentage;
    e.g. all stocks that have an optimal holding of less than 0.1%.
"""

class Stocks(object):
    def __init__(self, path, div):
        # Initializations
        self.colMC = 'Market Capitalization'
        self.colDY = 'Dividend Yield'
        # Storing inputs
        self.path = path
        self.data = None

        # Loading data
        self.loadXLSData()

        # Sort companies by dividend yield
        self.sortDividendYield()

        # Add a Market Capitalization % column
        self.getWeightedMarketCap()

        # Write out an Excel file
        self.writeStocks()

        # Return a stock investment Excel file
        self.makeGuide(div)


    def grabXLSData(self):
        pass

    def loadXLSData(self):
        """Read in data from an Excel spreadsheet"""
        self.data = pd.read_excel(self.path)
        self.data[self.colDY] = self.data[self.colDY].fillna(0) / 100
        self.deleteData()
        self.str2int()

    def deleteData(self):
        uselessCols = ['S&P 500 (R)', 'Security Type']
        for col in uselessCols:
            if col in self.data.columns:
                self.data = self.data.drop(col, axis=1)

    def str2int(self):
        suffixDict = {'K': 10**3, 'M': 10**6, 'B': 10**9, 'T': 10**12}
        col = self.colMC
        if col in self.data.columns:
            self.data[col] = self.data[col].str.replace('$', '')
            scientificFloat = pd.to_numeric(self.data[col].str[:-1])
            power = self.data[col].str[-1].replace(suffixDict)
            self.data[col] = scientificFloat * power

        else:
            msg = "'{}' column not found in '{}' data set.".format(col, os.path.split(self.path)[-1])
            raise NameError(msg)

    def getWeightedMarketCap(self):
        totalMarketCap = self.data['Market Capitalization'].sum()
        weightedMarketCap = self.data['Market Capitalization'] / totalMarketCap
        self.data['Weighted Market Capitalization'] = weightedMarketCap

    def sortDividendYield(self):
        col = self.colDY
        if col in self.data.columns:
            self.data = self.data.sort_values(by=[col])

        else:
            msg = "'{}' column not found in '{}' data set.".format(col, os.path.split(self.path)[-1])
            raise NameError(msg)

    def writeStocks(self):
        try:
            fName = 'test.xlsx'
            writer = pd.ExcelWriter(fName)# engine='xlsxwriter')
            self.data.to_excel(writer, 'Sheet1')
            writer.save()
            self.formatExcel(fName)
        except PermissionError:
            msg = 'Please close {}. The program is trying to write to it.'.format(fName)
            raise PermissionError(msg)

    def writeGuide(self):
        # Write to guide.xlsx
        try:
            fName = 'guide.xlsx'
            writer = pd.ExcelWriter(fName)#, engine='xlsxwriter')
            cols = ['Symbol', 'Company Name', 'Optimal Holding %', 'Dividend Yield', 'Price Performance (52 Weeks)',
                    'Weighted Market Capitalization', 'Market Capitalization', 'Security Price']
            self.guide[cols].to_excel(writer, 'Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Formats
            centsFormat = workbook.add_format({'num_format': '$0.00'})
            dollarsFormat = workbook.add_format({'num_format': '$#,###,###,###,###'})
            percentFormat = workbook.add_format({'num_format': '0.00%'})

            # Formatting
            worksheet.set_column('D:G', 1, percentFormat)
            worksheet.set_column('H:H', 1, dollarsFormat)
            worksheet.set_column('I:I', 1, centsFormat)

            writer.save()
        except IOError:
            msg = 'Please close {}. The program is trying to write to it.'.format(fName)
            raise IOError(msg)
        return fName

    def formatExcel(self, fName):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(os.path.join(os.path.dirname(self.path), fName))
        ws = wb.Worksheets('Sheet1')
        ws.Columns.AutoFit()
        wb.Save()
        excel.Application.Quit()

    def makeGuide(self, div):
        self.guide = self.data[self.data[self.colDY] < div]
        weightedMarketCapLowDiv = self.guide['Weighted Market Capitalization'].sum()
        self.guide['Optimal Holding %'] = self.guide['Weighted Market Capitalization'] / weightedMarketCapLowDiv
        self.guide['Average Dividend Yield %'] = (self.guide['Optimal Holding %'] * self.guide[self.colDY]).sum() / 100
        fName = self.writeGuide()
        self.formatExcel(fName)

def main():
    # Initialization
    pd.set_option('display.expand_frame_repr', False)
    path = r"C:\Users\Daniel\Desktop\Python Scripts\ETF"
    # fName = glob.glob('sp.xls')[0]
    fName = 'sp500_data.xlsx'
    fullPath = os.path.join(path, fName)
    div = 1

    stocks = Stocks(fullPath, div)
    return stocks


if __name__ == '__main__':
    SP500 = main()

