import clr
clr.AddReferenceByName('Microsoft.Office.Interop.Excel, Version=11.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c')
from Microsoft.Office.Interop import Excel

ex = Excel.ApplicationClass()   
ex.Visible = False
ex.DisplayAlerts = False

#open EXCEL file
workbook = ex.Workbooks.Open(r'D:\Documents\GitHub\HFSS-Automation\readEXCEL\example.xlsx')

#get worksheet
worksheet=workbook.Worksheets[1]

#read single cell text
total_score = 0
for i in [1, 2, 3, 4]:
    course = worksheet.Range['A'+str(i)].Text
    score = worksheet.Range['B'+str(i)].Text
    total_score += float(score)
    AddWarningMessage('{}: {}'.format(course, score))

AddWarningMessage('Total: {}'.format(total_score))

#read range
xlrange = worksheet.Range["A1", "B4"]
AddWarningMessage(str(list(xlrange.Value2)))

''' Shown in HFSS Message Manager
*Global - Messages
  [warning] Math: 80
  [warning] Physics: 70
  [warning] History: 60
  [warning] Biology: 50
  [warning] Total: 260.0
  [warning] ['Math', 80.0, 'Physics', 70.0, 'History', 60.0, 'Biology', 50.0]
'''
#AddWarningMessage(str(dir(workbook)))
workbook.Close()