VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_newEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub NewSiteBtn_Click()
 
End Sub

Private Sub NewEditorBtn_Click()
Dim NewEd As String
 NewEd = "INSERT INTO Editors (Editor) VALUES ([Forms]![newEditor]![EdName]);"
  DoCmd.SetWarnings False
 DoCmd.RunSQL NewEd
  DoCmd.SetWarnings True
  
  
 Forms![EditorForm]![Editorssubform].Requery
 DoCmd.Close
End Sub
