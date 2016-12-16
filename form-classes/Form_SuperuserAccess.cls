VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SuperuserAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub LoginBtn_Click()
 Dim passStr As String
 If IsNull([Forms]![SuperuserAccess]![password]) = False Then
 passStr = [Forms]![SuperuserAccess]![password]
 If passStr = "TTNewton" Then
     DoCmd.Close
    DoCmd.OpenForm "Superuser Panel"
    
    Else
    MsgBox ("Invalid password!")
    End If
 Else
 'do Nothing
 End If
    
    End Sub
