'''
Created on Jun 22, 2015

@author: Bumsub
'''
#-*- coding: utf-8 -*-
import FileIO.Excel as excel
import Web.APIs.Geocoding as geocoding
import geocoder
    
if __name__ == '__main__':
        
    filename = '200601SaleApartment.xls'
    sheetname = 'Seoul'
    excelResult = excel.excelRead(filename.encode('utf-8'), sheetname.encode('utf-8'))
    
    temp_addr0 = ''
    temp_addr1 = '' 
    temp_addr2 = ''
    temp_addr3 = ''
    
    insertList = []
    
    for apartment in excelResult[1:]:
        insertListTemp = []
        addr0 = str(apartment[0].value)
        addr1 = str(int(apartment[1].value))
        addr2 = str(int(apartment[2].value))
        addr3 = str(apartment[3].value)
        
        #print(type(addr0))
        #print(addr1)
        
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
            
            
            #print(addr0+" "+addr1+" "+addr2+" "+addr3)
              
    #print(len(insertList[0]))
    #print(len(insertList))
    #print(insertList[1][3])
    
    excel.excelWriteNewFile('apartment.xlsx', 'Sheet1', insertList)
    
    excelResult2 = excel.excelRead('apartment.xlsx', 'Sheet1')
    
    geocodingResult = geocoding.geocodeList(excelResult2)
    
    #print(geocodingResult)
    
    #excel.excelWriteOnExistingFile('200601SaleApartment.xls', 'Seoul', 10, geocodingResult)
    
    