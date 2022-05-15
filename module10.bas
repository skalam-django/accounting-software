Attribute VB_Name = "Module10"
Sub InteriorTask(ByVal SheetName As String, ByVal Class As String, ByVal Route As String, ByVal PresentLeft As Boolean)
 Dim LastRow As Long, i As Long, StudentName As String
 Sheets(Route).Activate
 Application.EnableEvents = False
 Application.ScreenUpdating = False
 
 LastRow = ThisWorkbook.ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0).Row

 For i = 11 To LastRow
  On Error GoTo ExitSub
  StudentName = ActiveSheet.Range("B" & i).Value
  If StudentName = SheetName Then
   ClassIs = ActiveSheet.Range("E" & i).Value
   If ClassIs = Class Then
    Sheets(ActiveSheet.Name).Unprotect password:="00alam.arduino00"
    If PresentLeft Then
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "A" & LastRow, , 12, , False, 192, 192, 192)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "B" & LastRow, , 12, , False, 255, 204, 153)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "E" & LastRow, , 12, , False, 153, 204, 255)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "F" & LastRow, , 12, , False, 204, 153, 102)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "I" & LastRow, , 12, , False, 204, 255, 255)
     
    Else
     
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "A" & LastRow, , 12, , False, 255, 0, 0)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "B" & LastRow, , 12, , False, 255, 0, 0)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "E" & LastRow, , 12, , False, 255, 0, 0)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "F" & LastRow, , 12, , False, 255, 0, 0)
     Call Application.Run("'Accounting_Software.xlsm'!Interior", "I" & LastRow, , 12, , False, 255, 0, 0)
     
    End If
    'Sheets(ActiveSheet.Name).Protect password:="00alam.arduino00"
   Exit Sub
   End If
  End If
 Next
ExitSub:
 
End Sub



