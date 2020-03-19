import xlrd
import numpy as np
import csv
import urllib.request
from datetime import datetime
from pathlib import Path

class Excel:
    """
    column sequence:
    1. Timestamp
    2. Date
    3. Month
    4. Year
    5. Cases
    6. Deaths
    7. Country
    8. GeoId
    """
    def __init__(self, path_to_file):
        self.COUNTRY_INDEX = 6
        self.CASES_INDEX = 4
        self.DEATHS_INDEX = 5
        self.path = path_to_file
        self.wb = xlrd.open_workbook(self.path)
        self.sheet = self.wb.sheet_by_index(0)
        self.cols = self.sheet.ncols
        self.rows = self.sheet.nrows
        data = []
        self.countries = {}
        self.geocodes = {}
    def TransformData(self):
        for i in range(1,self.rows):
            country = self.sheet.cell_value(i,self.COUNTRY_INDEX)
            if self.countries.get(country) == None:
                self.countries[country] = [[self.sheet.cell_value(i,j) for j in range(self.COUNTRY_INDEX)]]
            else:
                self.countries[country].append([self.sheet.cell_value(i,j) for j in range(self.COUNTRY_INDEX)])
        # TODO: Sort and update indexes
        for country in self.countries:
            data = np.array(self.countries[country], dtype=int)
            sorted_data = data[data[:,0].argsort()]
            index_cases = 0
            index_deaths = 0
            cumulative_cases = 0
            cumulative_deaths = 0
            new_data = np.zeros((len(sorted_data),4), dtype=int)
            for i in range(len(sorted_data)):
                if sorted_data[i,self.CASES_INDEX] > 0 or index_cases > 0:
                    index_cases += 1
                    cumulative_cases += sorted_data[i,self.CASES_INDEX]
                    new_data[i,0] = int(index_cases)
                    new_data[i,1] = cumulative_cases
                if sorted_data[i,self.DEATHS_INDEX] > 0 or index_deaths > 0:
                    index_deaths += 1
                    cumulative_deaths += sorted_data[i,self.DEATHS_INDEX]
                    new_data[i,2] = int(index_deaths)
                    new_data[i,3] = cumulative_deaths
            self.countries[country] = np.hstack((sorted_data,new_data))

    def CreateCSV(self):
        with open('data/COVID-19_Data.csv', 'w', newline='') as csvfile:
            writer = csv.writer(csvfile, delimiter=',',
                                    quotechar='|', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(['Country','Date','Day','Month','Year','Cases','Deaths','Index Cases','Cumulative Cases','Index Deaths','Cumulative Deaths'])
            for country in self.countries:
                writer.writerows([[country] + self.countries[country][i].tolist() for i in range(len(self.countries[country]))])

def DownloadFile():
    now = datetime.now()
    url = 'https://www.ecdc.europa.eu/sites/default/files/documents/COVID-19-geographic-disbtribution-worldwide-%s.xlsx'%(now.isoformat().split('T')[0])
    path = 'data/COVID-19_Data-%s.xlsx'%(now.isoformat().split('T')[0])
    if not Path(path).is_file():
        print('Beginning file download ...')
        urllib.request.urlretrieve(url, path)
        print('Downloaded')
    else:
        print('File Already Exists')
    return path


#path = 'D:\\Github\\COVID-19-Analysis\\data\\COVID-19_Data.xlsx'
path = DownloadFile()
print(path)
excel = Excel(path)
excel.TransformData()
#print(excel.countries['Afghanistan'])
#print(excel.countries['Italy'])
excel.CreateCSV()