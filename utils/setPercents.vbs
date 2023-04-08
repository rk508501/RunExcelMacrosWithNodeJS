Option Explicit

Dim objExcel, objWorkbook, FilePath, AValue, BValue, CValue, DValue

' Check if three command line arguments were provided
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript scriptname.vbs AValue BValue CValue"
    WScript.Quit
End If

' Set the values of A, B, and C from the command line arguments
AValue = WScript.Arguments(0)
BValue = WScript.Arguments(1)
CValue = WScript.Arguments(2)

' Set the file path of the workbook
FilePath = "C:\Users\kulka\Downloads\SumMacro.xlsm"

' Create an instance of Excel
Set objExcel = CreateObject("Excel.Application")

' Make Excel visible so you can see what's happening
objExcel.Visible = True

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open(FilePath)

' Set the values of cells A1, B1, and C1 as decimal values
objWorkbook.Sheets(1).Cells(1,1).Value = AValue
objWorkbook.Sheets(1).Cells(1,2).Value = BValue
objWorkbook.Sheets(1).Cells(1,3).Value = CValue

' Calculate the sum of cells A1, B1, and C1
objWorkbook.Sheets(1).Cells(1,4).Formula = "=SUM(A1:C1)"

' Get the value of cell D1 which contains the sum of A1, B1, and C1
DValue = objWorkbook.Sheets(1).Cells(1,4).Value

' Save the workbook
objWorkbook.Save

' Close the workbook and Excel
objWorkbook.Close
objExcel.Quit

' Clean up
Set objWorkbook = Nothing
Set objExcel = Nothing

' Return the value of cell D1 as output
WScript.Echo DValue
