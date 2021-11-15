#from openpyxl import load_workbook
#from openpyxl import Workbook
import pandas as pd
import os

#current directory
cwd = os.getcwd()

#R564211A
salesFile = "R564211A"
prFile = "President Report.xlsx"
workingCap = "Book1"
operReport = "R5531001"
pr = ""
sales = ""
jdePR = ""
workCap = ""
operations = ""

#find the file
for file in os.listdir(cwd):
	if len(file) >= len(prFile) and file[:len(prFile)] == prFile:
		pr = file
	if len(file) >= len(operReport) and file[:len(operReport)] == operReport:
		operations = file
	

print ("found: ", pr)
print ("found: ", operations)

#create dataframe of operations PR
prSales = pd.read_excel(pr, sheet_name="EU & Pounds Produced-R5531001", index_col=False)
rowNum = len(prSales)+1

#create dataframe of operations

operationDF = pd.read_csv(operations, skiprows=3)


#find previous business day

previousDay = "11/11/2021"
newOp = (operationDF.loc[operationDF['Work Date'] == previousDay])

#write to pres xlsx

newOp.to_excel(pr, sheet_name="EU & Pounds Produced-R5531001", startrow=rowNum)


#workCap = "Book1 - 2021-11-04T082239.542.xls"


'''

#read workcapital raw data
workCapDF = pd.read_excel(workCap)


res= (workCapDF.groupby(['Location BP Cat-02','Record Type'])['Total Amount'].sum())
print (res['AT', 'A/P'])

#jdePrelimDF = pd.read_excel("President's Report_Final_JDE.xlsm",sheet_name='R5541021_REH9001_WorkCap')

#read sales
df = pd.read_csv(sales, skiprows=1)

#fix HQ Branch name
df.loc[df['Branch Location'] == 'HQ', 'ShippingBranchName'] = 'RPC HQ Branch Plant' 

#print (jdePrelimDF.columns)

			



#print (workCapDF.columns)

#print first row
#print (workCapDF[20:25])

df.loc[df['Branch Location'] == 'HQ', 'ShippingBranchName'] = 'RPC HQ Branch Plant' 

print (df['ShippingBranchName'][647:650])

#print (df[0:2])

#print (df.columns)



'''
