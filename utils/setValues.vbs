Option Explicit

Dim objExcel, objWorkbook, FilePath

' Set the file path of the workbook
FilePath = "C:\Users\kulka\Downloads\SumMacro.xlsm"

' Create an instance of Excel
Set objExcel = CreateObject("Excel.Application")

' Make Excel visible so you can see what's happening
objExcel.Visible = True

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open(FilePath)

' Set the values of cells A1, B1, and C1
objWorkbook.Sheets(1).Cells(1,1).Value = 100
objWorkbook.Sheets(1).Cells(1,2).Value = 200
objWorkbook.Sheets(1).Cells(1,3).Value = 300

' Save the workbook
objWorkbook.Save

' Close the workbook and Excel
objWorkbook.Close
objExcel.Quit

' Clean up
Set objWorkbook = Nothing
Set objExcel = Nothing
