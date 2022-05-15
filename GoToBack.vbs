FilePath="D:\Accounting Software\"
FileName="Accounting_Software.xlsm"
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
Err.Clear
On Error Resume Next
Set objExcel_WB = GetObject(FilePath & FileName, "Excel.Workbook")
objExcel.DisplayAlerts = False
objExcel_WB.Activate
objExcel.Application.Run FileName & "!ShowUserForm"
Set objExcel_WB = Nothing
Set objExcel = Nothing
WScript.Quit
