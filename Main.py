#-*- coding: utf-8 -*-

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
            
            temp_addr0 = addr0
            temp_addr1 = addr1
            temp_addr2 = addr2
            temp_addr3 = addr3
              
    excel.excelWriteNewFile(newFilepath, newSheetname, insertList)
    
    print('(find unique address) saved successfully!')
    
if __name__ == '__main__':
    
    filename = '201501전월세아파트.xls'
    sheetname = '서울'
    
    fileInfoList = excel.xlsToXlsx(filename.decode('utf-8'), sheetname.decode('utf-8'))

    newfile = fileInfoList[0][:-5]+'_unique.xlsx'
    newsheet = fileInfoList[1]
    
    findUniqueAddr(fileInfoList[0], fileInfoList[1], newfile, newsheet)
    
    excelResult2 = excel.excelRead(newfile, newsheet)
    
    #geocodingResult = geocoding.geocodeList(excelResult2)
    
    #excel.excelWriteOnExistingFile(newfile, newsheet, 'E', geocodingResult)
    
    
    