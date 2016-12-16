VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TypesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command79_Click()
Forms![TypesForm]![Typessubform].Requery
End Sub



Private Sub DeleteTypeBtn_Click()
sql = "DELETE [Article Types].* FROM [Article Types] WHERE ((([Article Types].ID)=[Forms]![TypesForm]![TypeID]));"

If MsgBox("You are about to delete an article type! Proceed?", vbYesNo, "Delete Type?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True
    Forms![TypesForm]![Typessubform].Requery
End If
End Sub
