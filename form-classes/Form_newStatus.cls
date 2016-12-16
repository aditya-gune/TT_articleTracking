VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_newStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub NewStatusbtn_Click()
Dim NewSite As String
 NewSite = "INSERT INTO Status (Status) VALUES (([Forms]![newStatus]![StatusName]));"
  DoCmd.SetWarnings False
 DoCmd.RunSQL NewSite
  DoCmd.SetWarnings True
  
  
 Forms![StatusesForm]![Statussubform].Requery
 DoCmd.Close
End Sub
