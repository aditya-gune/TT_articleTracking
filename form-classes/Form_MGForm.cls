VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command79_Click()
Forms![MGForm]![mgsubform].Requery
End Sub

Private Sub DeleteGroupBtn_Click()
Dim sql As String
sql = "DELETE [Media Groups].* FROM [Media Groups] WHERE ((([Media Groups].ID)=[Forms]![MGForm]![MGid]));"

If MsgBox("You are about to delete a Media Group! Proceed?", vbYesNo, "Delete Group?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True
    Forms![MGForm]![mgsubform].Requery
End If
End Sub
