VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_newGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub NewSiteBtn_Click()
 Dim NewSite As String
 NewSite = "INSERT INTO [Media Groups] ([Media Group]) VALUES (([Forms]![newGroup]![GroupName]));"
  DoCmd.SetWarnings False
 DoCmd.RunSQL NewSite
  DoCmd.SetWarnings True
  
 Forms![MGForm]![mgsubform].Requery
 DoCmd.Close
End Sub
