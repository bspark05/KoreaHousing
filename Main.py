'''
Created on Jun 22, 2015

@author: Bumsub
'''

import FileIO.Excel as excel
import Web.APIs.Geocoding as geocoding
import geocoder

def findUniqueAddr(existFilepath, existSheetname, newFilepath, newSheetname):
    excelResult = excel.excelRead(existFilepath, existSheetname)
    
    temp_addr0 = ''
    temp_addr1 = '' 
    temp_addr2 = ''
    temp_addr3 = ''
    
    insertList = []
    
    for row in excelResult[1:]:
        insertListTemp = []
        addr0 = str(row[0].value)
        addr1 = str(int(row[1].value))
        addr2 = str(int(row[2].value))
        addr3 = str(row[3].value)
        
        if (addr0+addr1+addr2+addr3) != (temp_addr0+temp_addr1+temp_addr2+temp_addr3) :
            
            insertListTemp.append(addr0)
            insertListTemp.append(addr1)
            insertListTemp.append(addr2)
            insertListTemp.append(addr3)
            
            insertList.append(insertListTemp)
            
            print(addr3)
            
            temp_addr0 = addr0
            temp_addr1 = addr1
            temp_addr2 = addr2
            temp_addr3 = addr3
              
    excel.excelWriteNewFile(newFilepath, newSheetname, insertList)
    
if __name__ == '__main__':
    
    filename = '200601SaleApartment.xls'
    sheetname = 'Seoul'
    newfile = 'apartment_test.xlsx'
    newsheet = 'Sheet1'
    
    #excel.xlsToXlsx(filename, sheetname)
    
    #findUniqueAddr(filename, sheetname, newfile, newsheet)
    
    excelResult2 = excel.excelRead(newfile, newsheet)
    
    geocodingResult = geocoding.geocodeList(excelResult2)
    
    excel.excelWriteOnExistingFile(newfile, newsheet, 'E', geocodingResult)
    
    
    