VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_StatusesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command79_Click()
Forms![StatusesForm]![Statussubform].Requery
End Sub



Private Sub DeleteStatusBtn_Click()
Dim sql As String
sql = "DELETE Status.* FROM Status WHERE (((Status.ID)=[Forms]![StatusesForm]![StatusID]));"

If MsgBox("You are about to delete a status! Proceed?", vbYesNo, "Delete Status?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True
    Forms![StatusesForm]![Statussubform].Requery
End If
End Sub
