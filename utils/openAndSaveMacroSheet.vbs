Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\kulka\Downloads\SumMacro.xlsm")

' Get the command line arguments for A1, B1, and C1
AValue = WScript.Arguments.Item(0)
BValue = WScript.Arguments.Item(1)
CValue = WScript.Arguments.Item(2)

' Set the values of cells A1, B1, and C1
objWorkbook.Sheets(1).Range("A1").Value = AValue
objWorkbook.Sheets(1).Range("B1").Value = BValue
objWorkbook.Sheets(1).Range("C1").Value = CValue

' Run the SumColumns macro
objExcel.Run "SumColumns"

' Save and close the workbook
objWorkbook.Save

' Return the value of cell D1
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\kulka\Downloads\SumMacro.xlsm")
OutputValue = objWorkbook.Sheets(1).Range("D1").Value
objWorkbook.Close
objExcel.Quit
WScript.StdOut.Write(OutputValue)
