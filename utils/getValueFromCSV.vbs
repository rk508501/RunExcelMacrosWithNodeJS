Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
Set objWorkbook = objExcel.Workbooks.Open("C:\Automation\RunExcelMacrosWithNodeJS\data\SumMacro.xlsm")

' Get the cell address from the command line argument
CellAddress = WScript.Arguments.Item(0)

' Read the value of the specified cell
OutputValue = objWorkbook.Sheets(1).Range(CellAddress).Value

objWorkbook.Close
objExcel.Quit

' Output the value of the specified cell to the command line
WScript.StdOut.Write(OutputValue)
