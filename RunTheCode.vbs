FilePath ="D:\Project\"
FileName = "create_database1.xlsm"
On Error Resume Next
Set objExcel = GetObject(,"Excel.Application")
Err.Clear
On Error Resume Next
Set objExcel_WB = GetObject(FilePath & FileName,"Excel.Workbook")
objExcel.DisplayAlerts = False
objExcel_WB.Activate
objExcel.visible=True
objExcel.Application.Run FileName & "!ShowUserForm"