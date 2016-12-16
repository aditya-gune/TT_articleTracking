VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_newType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub NewTypeBtn_Click()
 Dim NewType As String
 NewType = "INSERT INTO [Article Types] (Type) VALUES (([Forms]![newType]![TypeName]));"
  DoCmd.SetWarnings False
 DoCmd.RunSQL NewType
  DoCmd.SetWarnings True
  
  
 Forms![TypesForm]![Typessubform].Requery
 DoCmd.Close
End Sub
