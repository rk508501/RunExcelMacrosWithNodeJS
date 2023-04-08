Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Automation\RunExcelMacrosWithNodeJS\data\SumMacro.xlsm")

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
objWorkbook.Close
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing
