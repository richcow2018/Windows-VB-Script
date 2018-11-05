'Dim objExcel

'ExcelFile = "C:\Users\user.name\Desktop\report.xlsm"
'Set objExcel = CreateObject("Excel.Application")
' Set objWorkbook = objExcel.workbooks.open(ExcelFile)

'objExcel.Workbooks.Open(ExcelFile)

'objExcel.Application.run "C:\Users\user.name\Desktop\report.xlsm!ReportMacro"

' objExcel.Application.Run "C:\Users\user.name\Desktop\report.xlsm!ReportMacro"
run_macro

Sub run_macro()
    Dim objExcel
    Dim xlBook
    Dim FolderFromPath
	Dim dteNow
dteNow = Date()
dteNow = DateAdd("d", dteNow, -1)

    Set objExcel = CreateObject("Excel.application")
	objExcel.DisplayAlerts = False
	'set thisFileName = "C:\Users\user.name\Desktop\test.xlsx"
	'Set objWorkbook = objExcel.Workbooks.Add()
    'FolderFromPath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") 
    set xlBook = objExcel.Workbooks.Open("C:\Users\user.name\Desktop\report.xlsm")
    objExcel.Application.run "'" & xlBook.Name & "'!ReportMacro"
	
  
   objExcel.ActiveWorkbook.SaveAs "C:\Users\user.name\Desktop\" &  Right("0" & Month(dteNow), 2) & "-" & Right("0" & Day(dteNow), 2) & "-" _
    & Year(dteNow) & ".xlsx", 51
 
'objExcel.close

    objExcel.Application.Quit
End Sub
