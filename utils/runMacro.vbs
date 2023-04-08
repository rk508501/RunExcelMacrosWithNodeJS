Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\kulka\Downloads\SumMacro.xlsm")		
objExcel.Run "SumColumns"
objWorkbook.Save
objWorkbook.Close
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
