import openpyxl
import xlrd
from configparser import ConfigParser

# load exce file
usercfg = openpyxl.load_workbook('makefiles/userCfg.xlsx')
userSheet = usercfg.get_sheet_names()
userSheets = usercfg.get_sheet_by_name(userSheet[0])
SourceFileName = userSheets.cell(row=8, column=2).value
print('Value  = ', (SourceFileName != None))
while SourceFileName == None :
    print("11")
