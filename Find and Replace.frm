VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FindReplace 
   Caption         =   "Find and Replace"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4800
   OleObjectBlob   =   "Find and Replace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Finding As String, FindTarget As Object, FirstAddress As String, ReplaceValue As String

Private Sub cmbClear_Click()
Me.tbFind.Value = ""
tbReplace.Value = ""
Finding = ""
Set FindTarget = Nothing
Me.cmbNext.Visible = False
Me.cmbNext.Enabled = False
Me.cmbPrevious.Visible = False
Me.cmbPrevious.Enabled = False
End Sub

Private Sub cmbFind_Click()
Finding = Me.tbFind.Value
If Finding <> "" Then
With Sheets(ActiveSheet.Name).Range("A:U")
Set FindTarget = .Find(Finding, LookIn:=xlValues)
If Not FindTarget Is Nothing Then
FirstAddress = FindTarget.Address
Application.EnableEvents = False
Sheets(ActiveSheet.Name).Range(FirstAddress).Select
Me.cmbNext.Visible = True
Me.cmbNext.Enabled = True
Else
GoTo DoneFinding
End If

Application.EnableEvents = True

End With
Else
MsgBox "Please Input The Data to Find", vbInformation + vbOKOnly, "Empty Input Box"
End If
GoTo ExitSub
DoneFinding:
MsgBox "No Such Data Found More"
Application.EnableEvents = True
ExitSub:
End Sub
Private Sub cmbNext_Click()

Application.EnableEvents = False
With Sheets(ActiveSheet.Name).Range("A:U")

On Error GoTo ExitSub
Set FindTarget = .FindNext(FindTarget)

If FindTarget Is Nothing Then
GoTo DoneFinding
End If

Sheets(ActiveSheet.Name).Range(FindTarget.Address).Select
Me.cmbPrevious.Visible = True
Me.cmbPrevious.Enabled = True

Set FindTarget = .FindNext(FindTarget)
If FindTarget.Address = FirstAddress Then
Me.cmbNext.Enabled = False
MsgBox "Finsihed"
End If
Set FindTarget = .FindPrevious(FindTarget)

End With
GoTo ExitSub
DoneFinding:
MsgBox "No Such Data Found More"
ExitSub:
Application.EnableEvents = True
End Sub
Private Sub cmbPrevious_Click()
Application.EnableEvents = False
With Sheets(ActiveSheet.Name).Range("A:U")

On Error GoTo ExitSub
Set FindTarget = .FindPrevious(FindTarget)
If FindTarget Is Nothing Then
GoTo DoneFinding
End If

Sheets(ActiveSheet.Name).Range(FindTarget.Address).Select
Me.cmbNext.Visible = True
Me.cmbNext.Enabled = True

If FindTarget.Address = FirstAddress Then
Me.cmbPrevious.Enabled = False
MsgBox "Finsihed"
End If

End With

GoTo ExitSub
DoneFinding:
MsgBox "No Such Data Found More"
ExitSub:
Application.EnableEvents = True
End Sub
Private Sub cmbReplace_Click()
Pass = InputBox("Password", "Enter Password")
If Pass = "0ALAM0" Then
Sheets(ActiveSheet.Name).Unprotect Password:="0ALAM0"
Application.EnableEvents = False
With Sheets(ActiveSheet.Name).Range("A:U")
CellValue = Sheets(ActiveSheet.Name).Range(FindTarget.Address).Value
ReplaceValue = tbReplace.Value
ReplacedValue = Replace(CellValue, Finding, ReplaceValue, 1, , vbTextCompare)
Sheets(ActiveSheet.Name).Range(FindTarget.Address).Value = ReplacedValue
End With
Sheets(ActiveSheet.Name).Protect Password:="0ALAM0"
Application.EnableEvents = True
Else
MsgBox "Try Again", vbCritical + vbOKOnly, "Wrong Password"
End If
End Sub
Private Sub cmbReplaceAll_Click()
Finding = Me.tbFind.Value
If Finding <> "" Then
Pass = InputBox("Password", "Enter Password")
If Pass = "0ALAM0" Then
Sheets(ActiveSheet.Name).Unprotect Password:="0ALAM0"
Application.EnableEvents = False
With Sheets(ActiveSheet.Name).Range("A:U")
Set FindTarget = .Find(Finding, LookIn:=xlValues)

If Not FindTarget Is Nothing Then
FirstAddress = FindTarget.Address
Do
CellValue = Sheets(ActiveSheet.Name).Range(FindTarget.Address).Value
ReplaceValue = tbReplace.Value
ReplacedValue = Replace(CellValue, Finding, ReplaceValue, 1, , vbTextCompare)
Sheets(ActiveSheet.Name).Range(FindTarget.Address).Value = ReplacedValue
Set FindTarget = .FindNext(FindTarget)
If FindTarget Is Nothing Then
Exit Do
End If
Loop Until FindTarget.Address = FirstAddress
End If
End With
End If
Sheets(ActiveSheet.Name).Protect Password:="0ALAM0"
Application.EnableEvents = True
Else
MsgBox "Try Again", vbCritical + vbOKOnly, "Wrong Password"
End If

End Sub
Private Sub tbFind_Change()
If Me.tbFind.Value <> "" Then
tbReplace.Visible = True
cmbReplace.Visible = True
cmbReplaceAll.Visible = True
cmbClear.Enabled = True
Else
tbReplace.Visible = False
cmbReplace.Visible = False
cmbReplaceAll.Visible = False
cmbClear.Enabled = False
End If

Me.cmbNext.Visible = False
Me.cmbNext.Enabled = False
Me.cmbPrevious.Visible = False
Me.cmbPrevious.Enabled = False

End Sub

Private Sub tbReplace_Change()
If Me.tbReplace.Value <> "" Then
cmbClear.Enabled = True
Else
cmbClear.Enabled = False
End If

End Sub

Private Sub UserForm_Initialize()
cmbPrevious.Visible = False
cmbNext.Visible = False
tbReplace.Visible = False
cmbReplace.Visible = False
cmbClear.Enabled = False

End Sub


