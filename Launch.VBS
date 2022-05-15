Set objFSO = CreateObject("Scripting.FileSystemObject")
strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)

sPath = strFolder & "\Accounting_Software.xlsm"

Set oExcelApp = CreateObject("Excel.Application")

With oExcelApp
    .Visible = False
    .DisplayAlerts = False
    Set oWB = .Workbooks.Open(sPath)
End With
SET oExcelApp = NOTHING
Wscript.Quit
