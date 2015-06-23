'''
Created on Jun 1, 2015

@author: Bumsub
'''


import xlrd
import openpyxl
import xlwt



def excelRead(filepath, sheetname):
    workbook = xlrd.open_workbook(filepath)
    worksheet = workbook.sheet_by_name(sheetname)
    
    num_rows = worksheet.nrows -1
    curr_row = -1
    result = []
    
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        result.append(row)
    
    return result

def excelWriteOnExistingFile(filepath, sheetname, columnNum, insert): 
    wb = xlrd.open_workbook(filepath)
    ws = wb.sheet_by_name(sheetname)
    workbook = openpyxl.load_workbook(filepath)
    worksheet = workbook.active
    
    num_rows = ws.nrows -1
    curr_row = -1
    
    while curr_row < num_rows:
        curr_row += 1
        
        worksheet[columnNum+str(curr_row+1)] = insert[curr_row][0]
        asciiNum = ord(columnNum)+1
        columnNumPlus = chr(asciiNum)
        worksheet[columnNumPlus+str(curr_row+1)] = insert[curr_row][1]
        
    workbook.save(filepath)
    print('saved successfully!')
    
def excelWriteNewFile(filepath, sheetname, insert):
    wb = xlwt.Workbook()
    ws = wb.add_sheet(sheetname)
    
    i=0
    while i<len(insert[0]):
        j=0
        while j<len(insert):
            print(insert[j][i])
            #print(type(insert[j][i]))
            ws.write(j, i, unicode(insert[j][i]))
            j+=1
        i+=1
        
    wb.save(filepath)
    
    