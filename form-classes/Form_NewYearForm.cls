VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_NewYearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AssignTab_Click()
MsgBox (Screen.ActiveControl.name)
'MsgBox (Forms!NewYearForm.PgCtrl.Rollover_assigned2015.Form![Bundle Topic])


End Sub



Private Sub Clearbtn_Click()
Me.EditorCtrl = Null
Me.StatusCtrl = Null
Call RefreshArticleBtn_Click
End Sub

Private Sub DeleteArticleBtn_Click()
Dim deletesql As String

'Delete all articles where article.id = idbox on form
deletesql = "DELETE Articles.*, Articles.ID FROM Articles WHERE (((Articles.ID)=[Forms]![NewYearForm]![IDbox]));"

'if the user presses "Yes" after the warning, delete
If MsgBox("You are about to delete an article! Proceed?", vbYesNo, "Delete Article?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL deletesql
    DoCmd.SetWarnings True
    
    'requery the subforms
    Forms![NewYearForm]![Rollover_assigned2015].Requery
    Forms![NewYearForm]![Rollover_inv2015].Requery
End If
End Sub

Private Sub EditBundleBtn_Click()
If CurrentProject.AllForms("SearchForm").IsLoaded Then
MsgBox ("Error: Please close the Search Form before attempting this operation.")
Else
'open editbundle
DoCmd.OpenForm "EditBundle", , , "EditBundle.ID = " & Me.fktopicbox
End If
End Sub

Private Sub FormHeader_Click()
Call RestoreMain
End Sub

Private Sub RefreshArticleBtn_Click()

'If the Assigned tab is selected, refresh assigned form
If PgCtrl.Value = 0 Then
    Call refreshBoxes
    Forms!NewYearForm.Rollover_assigned2015.Requery

'else refresh the invoice form
Else
    Call refreshBoxes
    Forms!NewYearForm.Rollover_inv2015.Requery
End If

End Sub

Public Function refreshBoxes()
Me.topicbox.Requery
Me.authorbox.Requery
Me.idbox.Requery
Me.fktopicbox.Requery
Me.stypebox.Requery
End Function

Private Sub Rollover_assigned2015_Enter()
Forms!NewYearForm.Rollover_assigned2015.Requery
End Sub


Private Sub Rollover_inv2015_Enter()
Forms!NewYearForm.Rollover_inv2015.Requery
End Sub
