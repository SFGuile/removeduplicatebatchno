import xlrd
import pandas as pd
from datetime import datetime

excelfilename=r'C:\Users\redmond\Desktop\架位批次库存报表.xls'

exportexcelname=excelfilename[0:-4]+datetime.today().strftime('%Y%m%d%H%M%S')+".xlsx"


data = xlrd.open_workbook(excelfilename)

table = data.sheet_by_index(0)
 
nrows = table.nrows
 
ncols = table.ncols 

unsortList = []

for i in range(1,nrows ):
  unsortList.append(table.row_values(i))


unsortList.sort(key=lambda t: (t[0], t[9]))

preprodno=""

sortlist=[]
sortlist.append(table.row_values(0))

myindex=0

for everyrow in unsortList:
     prodno=everyrow[0]
     batchno=everyrow[9]

     if prodno!=preprodno:
         sortlist.append(everyrow)
         myindex=myindex+1
         preprodno=prodno
     else:
         sortlist[myindex][9]=batchno
         preprodno=prodno

    


df = pd.DataFrame(sortlist)
writer = pd.ExcelWriter(exportexcelname, engine='xlsxwriter')
print(df)
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

print(r"完成")


 

     