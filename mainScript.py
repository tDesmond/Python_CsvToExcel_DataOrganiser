__author__ = 'Timmy Desmond'

from DataAnalyser import CsvDataAnalyser


##  Data analyser variables
##  Varaibles that can be called from the class (e.g analysedData.dataHeadings)
##  dataHeadings        -   Headings of the data to be analysed
##  numberOfRows        -   The number of rows of data in the csv file
##  csvData             -   All csv information converted to an array
##  columnLetters       -   Letters of the columns used for the excel file
##  numberOfColumns     -   Number of columns of data

def main():
    analysedData = CsvDataAnalyser('PCMMYM79013')

if __name__ == '__main__':
    main()
