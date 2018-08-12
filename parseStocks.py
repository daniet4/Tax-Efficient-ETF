# This program parses Stock Screener data from Fidelity located at https://research2.fidelity.com/pi/stock-screener
# The goal is to determine the best stocks to hold in a taxable brokerage account to minimize dividend yields.
# It is desirable to concentrate investment returns in the stock price to avoid capital gains taxes.
# For tax efficiency, the remaining high dividend yielding stocks will be held in a tax advantaged account such as a
# 401k, IRA, or HSA.

# Updated: 8/11/18

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

import os

from bs4 import BeautifulSoup
import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell


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

        # Removing extraneous columns
        self.deleteData()

        # Convert abbreviations to integers
        self.str2int()

        # Sort companies by dividend yield
        self.sortDividendYield()

        # Add a Market Capitalization % column
        self.getWeightedMarketCap()

        # Write out an Excel file
        self.writeStocks()

        # Return a stock investment Excel file
        self.getGuide(div)


    def grabXLSData(self):
        pass

    def loadXLSData(self):
        """Read in data from an Excel spreadsheet"""
        self.data = pd.read_excel(self.path)

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
        # writer = pd.ExcelWriter(self.path)
        try:
            writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
            self.data.to_excel(writer, 'Sheet1')
            writer.save()
        except PermissionError:
            msg = 'Please close {}. The program is trying to write to it.'.format('test.xlsx')
            raise PermissionError(msg)

    def writeGuide(self):
        try:
            writer = pd.ExcelWriter('guide.xlsx', engine='xlsxwriter')
            self.guide.to_excel(writer, 'Sheet1')
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            writer.save()
        except PermissionError:
            msg = 'Please close {}. The program is trying to write to it.'.format('test.xlsx')
            raise PermissionError(msg)

    def getGuide(self, div):
        self.guide = self.data[self.data[self.colDY] < div]
        weightedMarketCapLowDiv = self.guide['Weighted Market Capitalization'].sum()
        print(weightedMarketCapLowDiv)
        self.guide['Optimal Holding %'] = self.guide['Weighted Market Capitalization'] / weightedMarketCapLowDiv * 100
        self.guide['Average Dividend Yield %'] = (self.guide['Optimal Holding %'] * self.guide[self.colDY]).sum() / 100
        self.writeGuide()

def main():
    # Initialization
    pd.set_option('display.expand_frame_repr', False)

    path = r"E:\Daniel\Downloads"
    fName = r"sp500 data.xls"
    fullPath = os.path.join(path, fName)
    div = 1

    stocks = Stocks(fullPath, div)
    return stocks


if __name__ == '__main__':
    SP500 = main()
    stop = 1
