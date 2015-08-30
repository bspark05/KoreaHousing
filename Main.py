#-*- coding: utf-8 -*-

import FileIO.Excel as excel

def findUniqueAddr(existFilepath, existSheetname):
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
              
    return insertList

def findUniqueAddr2(existFilepath, existSheetname, inputList):
    excelResult = excel.excelRead(existFilepath, existSheetname)
     
    temp_addr0 = ''
    temp_addr1 = '' 
    temp_addr2 = ''
    temp_addr3 = ''
     
    insertList = []
     
    for inputRow in inputList:
        temp_addr0 = inputRow[0]
        temp_addr1 = str(inputRow[1])
        temp_addr2 = str(inputRow[2])
        temp_addr3 = inputRow[3]
        
        insertListTemp = []
        
        index = 0
        for existingRow in excelResult:
            index+=1
            
            addr0 = str(existingRow[0].value)
            addr1 = str(int(existingRow[1].value))
            addr2 = str(int(existingRow[2].value))
            addr3 = str(existingRow[3].value)
            
            if (addr0+addr1+addr2+addr3) == (temp_addr0+temp_addr1+temp_addr2+temp_addr3) :
                break
            
            if index == len(excelResult):
                insertListTemp.append(temp_addr0)
                insertListTemp.append(temp_addr1)
                insertListTemp.append(temp_addr2)
                insertListTemp.append(temp_addr3)    
         
                insertList.append(insertListTemp)
        
    return insertList
    
if __name__ == '__main__':
    
#     ##Part 1
#      
#     #step 1 - finding unique address        
#              
#     filename = '201412매매아파트.xls'
#     sheetname = '서울'
#                
#     fileInfoList = excel.xlsToXlsx(filename.decode('utf-8'), sheetname.decode('utf-8'))
#         
#     newfile = fileInfoList[0][:-5]+'_unique.xlsx'
#     newsheet = fileInfoList[1]
#         
#     uniqueAddr = findUniqueAddr(fileInfoList[0], fileInfoList[1])
#         
#     excel.excelWriteNewFile(newfile, newsheet, uniqueAddr)
#         
#     print('saved successfully (step1)')
#      
#       
#     #step 2 - update dictionary
#         
#     filenameDic = 'Dictionary.xlsx'
#     sheetnameDic = 'Sheet1'
#        
#         #uniqueAddr2 - dictionary에 들어갈 unique 값
#     uniqueAddr2 = findUniqueAddr2(filenameDic, sheetnameDic, uniqueAddr)
#     print(len(uniqueAddr))
#        
#         #2-1 - EMD code matching
#     filenameEMD = 'Seoul_EMD_code.xlsx'
#     sheetnameEMD = 'Seoul_EMD_code'
#        
#     resultEMD = excel.excelRead(filenameEMD, sheetnameEMD)
#        
#     for addr2 in uniqueAddr2:
#         tempEMD = '-1'
#         addrStrip = addr2[0].strip()
#              
#         for rowEMD in resultEMD:
#             rowEMDStr = str(rowEMD[1].value)
#                
#             if addrStrip == rowEMDStr:
#                 tempEMD = str(int(rowEMD[0].value))
#                 break
#         addr2.append(tempEMD)
#            
#     print len(uniqueAddr2)
#     if len(uniqueAddr2) != 0:
#         excel.excelWriteOnExistingFile2(filenameDic, sheetnameDic, uniqueAddr2)
#         print('saved successfully in Dictionary')
      
             
#   ##Part 2 - #3 Geocoding - AddressMatching project

    ## Part 3
    #4 - Matching excel file and BD_MGT_SN in Dictionary
      
    filenameMat = '201412SaleApartment.xlsx'
    sheetnameMat = 'Seoul'
        
    matResult = excel.excelRead(filenameMat, sheetnameMat)
        
        #4-1 - EMD code matching
    filenameEMD = 'Seoul_EMD_code.xlsx'
    sheetnameEMD = 'Seoul_EMD_code'
        
    resultEMD = excel.excelRead(filenameEMD, sheetnameEMD)
      
    emdMatList = []  
    for addrMat in matResult[1:]:
        tempEMD = '-1'
        addrStrip = str(addrMat[0].value).strip()
              
        for rowEMD in resultEMD:
            rowEMDStr = str(rowEMD[1].value)
                
            if addrStrip == rowEMDStr:
                tempEMD = str(int(rowEMD[0].value))
                break
        emdMatList.append(tempEMD)
        
    excel.excelWriteOnExistingFile3(filenameMat, sheetnameMat, 'k', emdMatList)
    print('saved successfully in '+ filenameMat)    
     
        #4-2 - Matching excel file and dictionary
     
    filenameDicRe = 'Dictionary_result.xlsx'
    sheetnameDicRe = 'Sheet2'
     
    resultDicRe = excel.excelRead(filenameDicRe, sheetnameDicRe)
     
    matResult2 = excel.excelRead(filenameMat, sheetnameMat)
     
    snList = []
    for addrMat1 in matResult2[1:]:
        tempBd_SN = '-1'
        tempeqb_SN = '-1'
        tempPointX = '-1'
        tempPointY = '-1'
        tempSnList = []
         
        addr1 = str(addrMat1[10].value)[:5]
        addr2 = str(addrMat1[10].value)[5:8]
        addr3 = str(int(addrMat1[1].value))
        addr4 = str(int(addrMat1[2].value))
         
        #print(addr1+' '+addr2+' '+addr3+' '+addr4)
         
        for addrDicRe in resultDicRe:
            addrD1 = str(addrDicRe[4].value)[:5]
            addrD2 = str(addrDicRe[4].value)[5:8]
            addrD3 = str(int(addrDicRe[1].value))
            addrD4 = str(int(addrDicRe[2].value))
             
            if addr1 == addrD1 and addr2 == addrD2 and addr3 == addrD3 and addr4 == addrD4 :
                tempBd_SN = str(addrDicRe[5].value)
                tempeqb_SN = str(int(addrDicRe[6].value))
                tempPointX =str(addrDicRe[7].value)
                tempPointY =str(addrDicRe[8].value)
                break
        tempSnList = [tempBd_SN, tempeqb_SN, tempPointX, tempPointY]
        snList.append(tempSnList)
     
    excel.excelWriteOnExistingFile4(filenameMat, sheetnameMat, 'l', snList)
    print('saved successfully in '+ filenameMat)
                
                