VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'printf debugger for invoice date
Private Sub Command21_Click()
If IsNull(Me.InvStart.Value) Then
    MsgBox "Null!"
Else
    MsgBox Me.InvStart.Value
End If
End Sub
'printf debugger for WIP Check box
Private Sub Command41_Click()
If Me.WIPCheckBox = True Then MsgBox "WIP IS CHECKED"
End Sub
'test debugger to enable statusctrl
Private Sub Command42_Click()
MsgBox Me.StatusCtrl.Enabled
End Sub

'OnClick for all articles button
'Set all fields to Null or blank
'And then call Search_Click
Private Sub AllArticlesBtn_Click()
Me.EditorCtrl = Null
Me.AuthorCtrl = Null
Me.MGBox = Null
Me.MGCtrl = Null
Me.TopicCtrl = Null
Me.SiteCtrl = Null
Me.StatusCtrl.Enabled = True
Me.StatusCtrl = Null
Me.WIPCheckBox = False
Me.CompleteCheckBox = False
Me.WIPCheckBox.Enabled = True
Me.CompleteCheckBox.Enabled = True
Me.InvStart.Value = Null
Me.InvEnd.Value = Null

'set pub dates to current date + 2 years
Me.PubStart.Value = "1/1/2014"
Me.PubEnd.Value = DateAdd("yyyy", 2, date)

Call Search_Click
End Sub

Private Sub Command78_Click()
Me!Text75.SetFocus
DoCmd.RunCommand acCmdCopy
Me.SetFocus
End Sub

Private Sub Command81_Click()

Me.Command81.SetFocus

DoCmd.OpenForm "InvoiceForm", , , , , , Me.idbox


End Sub

Private Sub EditBundleBtn_Click()
'If the new year carryover form is loaded, require user to close it
'to avoid edit conflicts
If CurrentProject.AllForms("NewYearForm").IsLoaded Then
MsgBox ("Error: Please close the 2016 Carryover Report before attempting this operation.")
Else
DoCmd.OpenForm "EditBundle", , , "EditBundle.ID = " & Me.Text56
End If
End Sub

Private Sub DeleteArticleBtn_Click()
'if the user presses "Yes" after the warning, delete
If MsgBox("You are about to delete an article! Proceed?", vbYesNo, "Delete Article?") = vbYes Then
    DoCmd.SetWarnings False
    DoCmd.OpenQuery ("ArticleDelete")
    DoCmd.SetWarnings True
    Forms![SearchForm]![QuerySubform].Requery
End If
End Sub

Private Sub Form_Load()
If Me.OpenArgs > 10 Then
Me.TopicCtrl = Me.OpenArgs
Me.[QuerySubform].Requery
ElseIf Me.OpenArgs < 10 And Me.OpenArgs > 0 Then
Me.EditorCtrl = Me.OpenArgs
Me.[QuerySubform].Requery
ElseIf IsNull(Me.OpenArgs) Then
'do nothing
End If

If IsNull(Me.PubStart.Value) Then
   Me.PubStart.Value = "1/1/2014"
End If
If IsNull(Me.PubEnd.Value) Then
    'set pub dates to current date + 2 years
   Me.PubEnd.Value = DateAdd("yyyy", 2, date)
End If
Me.[QuerySubform].Requery
End Sub

'If text123 on the main menu does not have an int
'then open form normally
'else set the editor chosen to the int in text123
Private Sub Form_Open(Cancel As Integer)

If (CurrentProject.AllForms("Main Menu").IsLoaded = False) Then
MsgBox ("The Main Menu is closed. Please close and reopen the database to reopen the Main Menu.")

ElseIf IsNull(Forms![Main Menu]!Text123) Then
    'nothing
Else
    Me.EditorCtrl = Forms![Main Menu]!Text123
    Me.PubEnd.Value = DateAdd("yyyy", 2, date)
    Me.QuerySubform.Requery
End If

End Sub

Private Sub FormHeader_Click()
Call RestoreMain
End Sub

Private Sub MGBox_AfterUpdate()
Me.MGCtrl = Me.MGBox.Value
End Sub

Public Function invoiceTable(updateSql As String)
Call Search_Click
Dim Db As Database
Dim recordAffect As Integer

Set Db = CurrentDb
On Error Resume Next
With Db

  .Execute updateSql, dbFailOnError
  recordAffect = .RecordsAffected

  Debug.Print ("after exec recordaffect= " & recordAffect)

Forms![SearchForm].QuerySubform.Requery

End With
End Function



'The main search query function
Private Sub Search_Click()

'if invoice fields are null
If IsNull(Me.InvStart) Or IsNull(Me.InvEnd) Then
    'if the wip check box is checked
    If Me.WIPCheckBox = "-1" Then
        'then set the subform's recordsource to NoInvoiceWIP1 query
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "NoInvoiceWIP1"
        
    'else if the complete box is checked
    ElseIf Me.CompleteCheckBox = "-1" Then
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "NoInvoiceComplete1"
    
    'else just run the no invoice query
    Else
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "NoInvoiceQuery1"
    End If
    
'else, if the invoice fields are NOT null, follow the same logic
'but with queries that include invoices as a searh criteria
Else
If Me.WIPCheckBox = "-1" Then
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "SearchInvoiceWIP1"
    ElseIf Me.CompleteCheckBox = "-1" Then
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "SearchInvoiceComplete1"
    Else
        Forms![SearchForm]![QuerySubform].Form.RecordSource = "SearchInvoiceQuery1"
    End If
End If

'requery the subform when finished
Forms![SearchForm]![QuerySubform].Requery
End Sub

Private Sub SiteCtrl_AfterUpdate()
'If a site has been selected, update
'the media group accordingly
Me.MGCtrl.Requery
Me.MGCtrl = Me.MGCtrl.ItemData(0)

End Sub

'if the complete box is checked,
'disable the WIP and status controls
Private Sub CompleteCheckBox_AfterUpdate()
Me.StatusCtrl.Enabled = True
Me.WIPCheckBox.Enabled = True
Select Case Me.CompleteCheckBox
    Case "-1"
        Me.StatusCtrl.Enabled = False
        Me.WIPCheckBox.Enabled = False
    Case "0"
       Me.StatusCtrl.Enabled = True
       Me.WIPCheckBox.Enabled = True
End Select
End Sub

'if the WIP box is checked,
'disable the complete and status controls
Private Sub WIPCheckBox_AfterUpdate()
Me.StatusCtrl.Enabled = True
Me.CompleteCheckBox.Enabled = True
Select Case Me.WIPCheckBox
        Case "-1"
            Me.StatusCtrl.Enabled = False
            Me.CompleteCheckBox.Enabled = False
        Case "0"
            Me.StatusCtrl.Enabled = True
            Me.CompleteCheckBox.Enabled = True
End Select
End Sub


