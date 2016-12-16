VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SitesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Command79_Click()
Forms![SitesForm]![SitesSubform].Requery
End Sub

Private Sub DeleteArticleBtn_Click()
Dim sql As String
sql = "DELETE Sites.* FROM Sites WHERE (((Sites.ID)=[Forms]![SitesForm]![SiteID]));"

If MsgBox("You are about to delete a site! Proceed?", vbYesNo, "Delete Site?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL sql
    DoCmd.SetWarnings True
    Forms![SitesForm]![SitesSubform].Requery
End If
End Sub
