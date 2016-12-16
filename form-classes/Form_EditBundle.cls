VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_EditBundle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub AddBundleSubform_Exit(Cancel As Integer)
'Dim bgexists As Integer
Dim nodue As Integer
Dim notarget As Integer
Dim nostatus As Integer
Dim noauthor As Integer

'bgexists = 0
nodue = 0
notarget = 0
nostatus = 0
noauthor = 0

'Create the ADODB recordset
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Dim sqlstr As String
sqlstr = "SELECT * FROM Articles WHERE Articles.fk_topic = " & Me.ID
rst.Open sqlstr, CurrentProject.Connection
Do While Not rst.EOF
    
    'Iterate through the rst while not End of File(EOF),
    'checking due dates, pub dates, authors, and statuses
    If IsNull(rst![Due Date].Value) Then
    nodue = 1
    End If
    
    If IsNull(rst![Target].Value) Then
    notarget = 1
    End If
    
    If IsNull(rst![Author].Value) Then
    noauthor = 1
    End If
    
    If IsNull(rst![Status].Value) Then
    nostatus = 1
    End If
    
    
    If Not rst.EOF Then rst.MoveNext
Loop
   
'If any of the values are NULL or "", throw an error and Cancel
If nostatus = 1 Then
    MsgBox "Error: All articles must have a status!", , "Error!"
    Cancel = True
End If

If noauthor = 1 Then
    MsgBox "Error: All articles must have an author! (Put your name for the Buyer's Guide article)", , "Error!"
    Cancel = True
End If

If nodue = 1 Then
    MsgBox "Error: All articles must have a Due Date!", , "Error!"
    Cancel = True
End If

If notarget = 1 Then
    MsgBox "Error: All articles must have a Publication Date!", , "Error!"
    Cancel = True
End If

End Sub

Private Sub Form_Close()
'If EditBundle was opened through SearchForm, then
'requery SearchForm's subform
If CurrentProject.AllForms("SearchForm").IsLoaded Then
    Form_SearchForm.QuerySubform.Requery
    
'Else if it was opened through NewYearForm, then
'requery its subforms
ElseIf CurrentProject.AllForms("NewYearForm").IsLoaded Then
    Form_NewYearForm.Rollover_assigned2015.Requery
    Form_NewYearForm.Rollover_inv2015.Requery
End If
End Sub

Private Sub Form_Load()
'If EditBundle was opened through SearchForm, then
'set the ID equal to SearchForm's idbox and requery
'SearchForm's subform
If CurrentProject.AllForms("SearchForm").IsLoaded Then
    Me.ID.Value = Form_SearchForm.Text56
    Form.Requery
    Me.Topic = Me.Topic.ItemData(0)
    
'Else if it was opened through NewYearForm, then
'set the ID equal to NewYearForm's idbox and requery
'NewYearForm's subform
ElseIf CurrentProject.AllForms("NewYearForm").IsLoaded Then
    Me.ID.Value = Form_NewYearForm.fktopicbox
    Form.Requery
    Me.Topic = Me.Topic.ItemData(0)

'Else throw error and close the form
Else
    MsgBox ("Error! Please open the Search Form and select a bundle to edit.")
    DoCmd.Close acForm, "EditBundle"
End If


End Sub

