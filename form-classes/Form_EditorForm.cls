VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database




Private Sub DeleteEditorBtn_Click()
Dim sql As String
sql = "DELETE Editors.* FROM Editors WHERE (((Editors.ID)=[Forms]![EditorForm]![EditorID]));"

If MsgBox("You are about to delete an editor! Proceed?", vbYesNo, "Delete Site?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True
    Forms![EditorForm]![Editorssubform].Requery
End If
End Sub

Private Sub refresheditor_Click()
Forms![EditorForm]![Editorssubform].Requery
End Sub
