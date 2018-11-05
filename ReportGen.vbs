Dim objExcel

ExcelFile = "C:\Users\jimmy.chu\Desktop\report.xlsm!ReportMacro"
Set objExcel = CreateObject("Excel.Application")
' Set objWorkbook = objExcel.workbooks.open(ExcelFile)

objExcel.Run ExcelFile

' objExcel.Application.Run "C:\Users\jimmy.chu\Desktop\report.xlsm!ReportMacro"

