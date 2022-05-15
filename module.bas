Attribute VB_Name = "Module1"
Public PresentLeft As Boolean

Sub BackButton()
 
 Dim UserIs As String
 
 Sheets(ActiveSheet.Name).Unprotect password:="00alam.arduino00"
 
 On Error Resume Next
 Sheets(ActiveSheet.Name).cmbShtBack.Visible = False
 Sheets(ActiveSheet.Name).shtcmbFind.Visible = False
 Sheets(ActiveSheet.Name).spbtnGradeUp.Visible = False
 
 On Error Resume Next
 Call Application.Run("'Accounting_Software.xlsm'!Unload_EditForm")
 On Error Resume Next
 Call Application.Run("'Accounting_Software.xlsm'!Unload_FindAndReplace")
 On Error GoTo Errhandller
 Call Application.Run("'Accounting_Software.xlsm'!SetExcelToNormal", False)
 
 On Error GoTo Errhandller
 UserIs = Application.Run("'Accounting_Software.xlsm'!User")
 Shell "wscript C:\Users\" & UserIs & "\AppData\Roaming\vbaTemp\GoToBack.vbs", vbNormalFocus

' On Error GoTo Errhandller
' Call Application.Run("'Accounting_Software.xlsm'!RetriveWorkingSheet", ThisWorkbook.Name, ActiveSheet.Name)
 
 Sheets(ActiveSheet.Name).Protect password:="00alam.arduino00"
 
 ThisWorkbook.Close SaveChanges:=True

 Exit Sub
 
Errhandller:
 MsgBox "The Workbook : Accounting_Software.xlsm isn't Available", vbCritical + vbOKOnly, "Error"
 ThisWorkbook.Close SaveChanges:=True

End Sub


Sub SetPic()

 Dim vbaTempLocation As String, SheetPassword As String
 On Error GoTo Errhandller
 UserIs = Application.Run("'Accounting_Software.xlsm'!User")
 vbaTempLocation = "C:\Users\" & UserIs & "\AppData\Roaming\vbaTemp\"
 SheetPassword = Application.Run("'Accounting_Software.xlsm'!ReadData", vbaTempLocation & "SheetPass.dat")
 Feedback = Application.Run("'Accounting_Software.xlsm'!InputBoxDK", "Please Type Password", "Password Required")
 Err.Clear
 If Feedback <> SheetPassword Then
  MsgBox "Wrong Password Entered" & vbCrLf & "Try Again Later", vbCritical + vbOKOnly, "Wrong Password"
  Exit Sub
 End If
 Set StudentPic = Application.FileDialog(msoFileDialogOpen)
 With StudentPic
  .AllowMultiSelect = False
  .Filters.Clear
  .Filters.Add "Image", "*.bmp;*.gif;*.jpg;*.jpeg;*.png"
  If .Show <> 0 Then
   For Each FileName In .SelectedItems
    PictureName = FileName
   Next
   Sheets(ActiveSheet.Name).Unprotect password:="00alam.arduino00"
   Sheets(ActiveSheet.Name).WS_StdntPic.Picture = LoadPicture(PictureName)
   Sheets(ActiveSheet.Name).WS_StdntPic.PictureSizeMode = 1
   Sheets(ActiveSheet.Name).Protect password:="00alam.arduino00"
  End If
 End With
 ThisWorkbook.Save
 MsgBox "Picture has been set sucessfully", vbInformation + vbOKOnly, "Status"
 Exit Sub
Errhandller:
 MsgBox "The Workbook : Accounting_Software.xlsm isn't Available", vbCritical + vbOKOnly, "Error"
 ThisWorkbook.Close SaveChanges:=True
 
End Sub

Function StatusOfShtObjects(ByVal objName As String, ByVal checkPara, ByVal checkRef, Optional ByVal Wb As String, Optional ByVal Sht As String) As Boolean
 
 Dim o As OLEObject
 If Wb = "" Then
  Wb = ActiveWorkbook.Name
 End If
 If Sht = "" Then
  Sht = ActiveSheet.Name
 End If
 
 StatusOfShtObjects = False
 
 For Each o In Workbooks(Wb).Sheets(Sht).OLEObjects
  
  If TypeName(o.Object) = objName Then
   With o
    If checkPara = checkRef Then
     StatusOfShtObjects = True
     Exit Function
    End If
   End With
  End If
 Next
 
End Function

