VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditForm 
   Caption         =   "Edit Form"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8115
   OleObjectBlob   =   "Edit Form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbtnEdit_Click()
With Application
.ScreenUpdating = True
.DisplayAlerts = True
Sheets(ActiveSheet.Name).Unprotect Password:="0ALAM0"
ActiveCell.Value = Me.tbValue.Value
Sheets(ActiveSheet.Name).Protect Password:="0ALAM0"
tbValue.SetFocus
.ScreenUpdating = False
.DisplayAlerts = False
End With
Me.Hide
End Sub
Private Sub txtPassword_Change()
If Len(txtPassword.Text) > 0 Then
Me.lblValue.Visible = True
Me.tbValue.Visible = True
Me.cmbtnEdit.Visible = True
Else
Me.lblValue.Visible = False
Me.tbValue.Visible = False
Me.cmbtnEdit.Visible = False
End If

If txtPassword.Value = "0ALAM0" Then
Me.lblValue.Enabled = True
Me.tbValue.Enabled = True
Me.cmbtnEdit.Enabled = True

Me.tbValue.Value = ActiveCell.Value
Me.tbValue.SetFocus
Else
Me.lblValue.Enabled = False
Me.tbValue.Enabled = False
Me.cmbtnEdit.Enabled = False
End If

End Sub

Private Sub UserForm_Activate()
If txtPassword.Value = "0ALAM0" Then
Me.tbValue.Value = ActiveCell.Value
End If
End Sub

Private Sub UserForm_Initialize()
txtPassword.SetFocus
End Sub

