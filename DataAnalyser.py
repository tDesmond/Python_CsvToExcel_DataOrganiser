__author__ = 'Timmy Desmond'

##  Data analyser variables
##  Varaibles that can be called from the class (e.g analysedData.dataHeadings)
##  dataHeadings        -   Headings of the data to be analysed
##  numberOfRows        -   The number of rows of data in the csv file
##  csvData             -   All csv information converted to an array
##  columnLetters       -   Letters of the columns used for the excel file
##  numberOfColumns     -   Number of columns of data


import csv
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

class CsvDataAnalyser:

    def __init__(self, csvFileName):
        self.csvFileName = csvFileName + '.csv'
        self.workbook = xlsxwriter.Workbook(csvFileName + '_results.xlsx')
        self.worksheet = self.workbook.add_worksheet()
        self.csvToArray()
        self.dataHeadings = self.csvData[0][9:]
        self.numberOfRows = len(self.csvData)
        self.resultsToExcel()
        self.setupGraph()
        self.workbook.close()


    def csvToArray(self):
        self.csvData = []
        self.columnLetters = []
        self.numberOfColumns = 0

        with open(self.csvFileName, 'rb') as f:
            reader = csv.reader(f)
            for row in reader:
                for i, element in enumerate(row):
                    try:
                        row[i] = float(row[i])
                    except ValueError:
                        pass
                self.csvData.append(row)
                
        self.numberOfColumns = len(self.csvData[0])
        
        for i in range(0, self.numberOfColumns - 1):
            sLetter = xl_rowcol_to_cell(0, i)
            sLetter = sLetter[:-1]
            self.columnLetters.append(sLetter)

    def arrayToExcel(self):
        self.csvData
        for i, element in enumerate(self.csvData):
            column = 'A' + str(i + 1)
            self.worksheet.write_row(column, self.csvData[i])

    def resultsToExcel(self):
        for i in self.columnLetters[9:]:
            row = []
            maxValCheck = ['Max Value']
            avgValCheck = ['Avg Value']
            minValCheck = ['Min Value']
            titles = ['Heading', 'Standard Deviation', 'Max Limit', 'Max Value', 'Average Value', 'Min value',
                      'Limit 75%', 'Limit 50%', 'Limit 25%', 'Min Limit 25%', '100%', '90%', '75%', 'Max Value %',
                      'Avg Value %', 'Min Value %', '25%', '10%', '0%']
            for j in range(5, 12):
                row.append(i + str(self.numberOfRows + j))
                self.worksheet.write('A' + str(self.numberOfRows + j), titles[j - 5])
            for j in range(14, 17):
                row.append(i + str(self.numberOfRows + j))
                self.worksheet.write('A' + str(self.numberOfRows + j), titles[j - 8])
            for j in range(19, 28):
                row.append(i + str(self.numberOfRows + j))
                self.worksheet.write('A' + str(self.numberOfRows + j), titles[j - 9])

            # Add excel formula's to output results
            headingFormula = '=' + i + str(1)
            stdDevFormula = '=STDEV(' + i + str(5) + ':' + i + str(19) + ')'
            maxLimit = '=' + i + str(4)
            minLimit = '=' + i + str(3)
            maxFormula = '=MAX(' + i + str(5) + ':' + i + str(19) + ')'
            avgFormula = '=AVERAGE(' + i + str(5) + ':' + i + str(19) + ')'
            minFormula = '=MIN(' + i + str(5) + ':' + i + str(19) + ')'
            maxLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.75) +' + row[6]
            minLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.25) +' + row[6]
            avgLimPCT = '=((' + row[2] + '-' + row[6] + ') * 0.5) +' + row[6]
            minValPCT = '=(((' + row[5] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
            avgValPCT = '=(((' + row[4] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
            maxValPCT = '=(((' + row[3] + '-' + row[6] + ') / (' + row[2] + '-' + row[6] + ')) * 100 )'
            pct75 = 75
            pct25 = 25
            pct90 = 90
            pct10 = 10
            pct100 = 100
            pct0 = 0
            formulas = [headingFormula, stdDevFormula, maxLimit, maxFormula,
                        avgFormula, minFormula, minLimit, maxLimPCT, avgLimPCT, minLimPCT, pct100, pct90, pct75, maxValPCT,
                        avgValPCT, minValPCT, pct25, pct10, pct0]
            for j in range(0, len(formulas)):
                self.worksheet.write(row[j], formulas[j])

    def setupGraph(self):
        lastRow = self.numberOfRows
        chart = self.workbook.add_chart({'type': 'line'})
        for i in range((lastRow + 19), (lastRow + 28)):
            lColors = ['red', 'orange', 'yellow', 'blue', 'green', 'purple', 'yellow', 'orange', 'red']
            color = lColors[i - (lastRow + 19)]
            name = '==Sheet1!$A$' + str(i)
            values = '==Sheet1!$J$' + str(i) + ':$DA$' + str(i)
            self.chart.add_series({
                'name': name,
                'categories': '=Sheet1!$J$1:$DA$1',
                'values': values,
                'line': {'color': color, 'width': 1.5},
            })

        # self.chart.set_size({'x_scale': 1.5, 'y_scale': 2})
        chart.set_x_axis({'major_gridlines': {
            'visible': True,
            'line': {'width': 1.25},
        },
            'interval_unit': 1,
            'label_position': 'low',
        })
        chart.set_y_axis({
            'min': -40,
            'max': 140,
        })

        chartPos = 'A' + str(lastRow + 30)
        self.worksheet.insert_chart(chartPos, chart, {'x_scale': 5, 'y_scale': 1.5})
        



        

    
        
        
