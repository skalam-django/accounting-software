Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Private Declare Function BlockInput Lib "USER32.dll" (ByVal fBlockIt As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const VK_CAPITAL = &H14
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Dim MyPassword As String

Sub CheckExpiry(ByVal ExpiryDate As String)

Call SetVBProjectPassword(ThisWorkbook, "0alam0")

End Sub
Sub CheckExpiry1(ByVal ExpiryDate As String)
Dim VBComp As VBIDE.VBComponent
Dim VBComps As VBIDE.VBComponents
'
'ExpiryDate = #3/3/2018#
MsgBox ExpiryDate
If Date >= CDate(replace(ExpiryDate, "#", "")) Then
MsgBox "Unlocking"
MsgBox "ProtectedVBProject = " & ProtectedVBProject(ThisWorkbook)

Call SyncVBAEditor
MsgBox "SyncVBAEditor"
Call SetVBProjectPassword(ThisWorkbook, "0alam0")
'Call SetVBProjectPassword(ThisWorkbook, "0alam0")
MsgBox "ProtectedVBProject = " & ProtectedVBProject(ThisWorkbook)
MsgBox "Unlocked"
Err.Clear

Set VBComps = ThisWorkbook.VBProject.VBComponents
MsgBox "set"
For Each VBComp In VBComps
MsgBox "VBComp= " & VBComp
Select Case VBComp.Type
'[MsgBox "VBComp.Type= " & VBComp.Type
Case vbext_ct_StdModule, vbext_ct_MSForm, vbext_ct_ClassModule
VBComps.Remove VBComp
MsgBox "Remove VBComp"
ThisWorkbook.Save
Case Else
With VBComp.CodeModule
.DeleteLines 1, CountOfLines
MsgBox "DeleteLines"
ThisWorkbook.Save
End With
ThisWorkbook.Save
End Select
Next VBComp
End If
Set VBComps = Nothing
Set VBComp = Nothing
End Sub

Function ProtectedVBProject(ByVal Wb As Workbook) As Boolean
' returns TRUE if the VB project in the active document is protected
Dim VBC As Integer
    VBC = -1
    On Error Resume Next
    VBC = Wb.VBProject.VBComponents.Count
    On Error GoTo 0
    If VBC = -1 Then
        ProtectedVBProject = True
    Else
        ProtectedVBProject = False
    End If
End Function

Sub SetVBProjectPassword(Wb As Workbook, ByVal Password As String)
 Const BreakIt As String = "%{F11}%TE+{TAB}{RIGHT}%V{+}{TAB}"
 Dim VBP As VBProject
 Dim OpenWin As VBIDE.Window
 Dim i As Integer
 Dim CapsLockState As Boolean
 Dim keys(0 To 255) As Byte
 MsgBox "Wb= " & Wb.Name & " Pass= " & Password
 If ProtectedVBProject(Wb) = True Then
   MsgBox "ProtectedVBProject(wb)= " & ProtectedVBProject(Wb)
 'BlockInput True
 
 info = WorkBookOpen(ThisWorkbook.Path & "\", ThisWorkbook.Name)
 MsgBox "info= " & info
 If info = False Then
  MsgBox NameBook & "File is being used"
 End If
 Call SetExcelToNormal
 Windows(replace(ThisWorkbook.Name, ".xlsm", "")).Visible = True
 Windows(replace(ThisWorkbook.Name, ".xlsm", "")).Activate
 ThisWorkbook.Sheets(1).Activate
 MsgBox "Activated"
 GetKeyboardState keys(0)
 CapsLockState = keys(VK_CAPITAL)
 Set VBP = Wb.VBProject
 Application.ScreenUpdating = False
 For Each OpenWin In VBP.VBE.Windows
  'MsgBox "OpenWin= " & OpenWin
  If InStr(OpenWin.Caption, "(") > 0 Then OpenWin.Close
 Next OpenWin
 Wb.Activate
 MsgBox "CapsLockState= " & CapsLockState
 If CapsLockState = True Then
  Call MakeCapsLockOff
 End If
  MsgBox "CapsLockState= " & CapsLockState
  MsgBox "BreakIt= " & BreakIt
 BlockInput True
 With Application
 '.ScreenUpdating = False
 Sleep 500
 .SendKeys BreakIt, True
 Sleep 500
 .SendKeys Password, True
 Sleep 500
 .SendKeys "{tab}", True
 Sleep 500
 .SendKeys Password, True
 Sleep 500
 .SendKeys "~", True
 Sleep 500
 .SendKeys "%{F11}~", True
 Sleep 500
 End With
 'SendKeys BreakIt & Password & "{tab}" & Password & "~" & "%{F11}~", True
 Application.ScreenUpdating = True
 Wb.Activate
 'SendKeys "%{F11}", True
 BlockInput False
 End If
End Sub

Sub SyncVBAEditor()
'=======================================================================
' SyncVBAEditor
' This syncs the editor with respect to the ActiveVBProject and the
' VBProject containing the ActiveCodePane. This makes the project
' that conrains the ActiveCodePane the ActiveVBProject.
'=======================================================================
With Application.VBE
If Not .ActiveCodePane Is Nothing Then
    Set .ActiveVBProject = .ActiveCodePane.CodeModule.Parent.Collection.Parent
End If
End With
End Sub

Private Sub MakeCapsLockOff()
 Dim keys(0 To 255) As Byte
 GetKeyboardState keys(0)
 keys(VK_CAPITAL) = 0
 SetKeyboardState keys(0)
End Sub









