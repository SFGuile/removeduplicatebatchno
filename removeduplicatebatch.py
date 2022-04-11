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


unsortList.sort(key=lambda t: (t[0], t[10]))

preprodno=""

sortlist=[]
sortlist.append(table.row_values(0))

myindex=0

for everyrow in unsortList:
     prodno=everyrow[0]
     batchno=everyrow[10]
     prodname=everyrow[1]
     prodsize=everyrow[2]
     monad=everyrow[3]
     prodmade=everyrow[4]
     prodadd=everyrow[5]
     prdoucedate=everyrow[6]
     availdate=everyrow[7]
     wareno=everyrow[8]
     deppos=everyrow[9]

     if prodno!=preprodno:
         sortlist.append(everyrow)
         myindex=myindex+1
         preprodno=prodno
     else:
         sortlist[myindex][1]=prodname
         sortlist[myindex][2]=prodsize
         sortlist[myindex][3]=monad
         sortlist[myindex][4]=prodmade
         sortlist[myindex][5]=prodadd
         sortlist[myindex][6]=prdoucedate
         sortlist[myindex][7]=availdate
         sortlist[myindex][8]=wareno
         sortlist[myindex][9]=deppos
         sortlist[myindex][10]=batchno
         preprodno=prodno

    


df = pd.DataFrame(sortlist)
writer = pd.ExcelWriter(exportexcelname, engine='xlsxwriter')
print(df)
df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

print(r"完成")


 

     