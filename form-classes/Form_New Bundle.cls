VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_New Bundle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'if bundle ID is null (that is, no bundle has been created),
'display error and force exit subform
Private Sub AddBundleSubform_Enter()
If IsNull(Me.ID) Then
MsgBox ("Please fill in series information before adding articles.")

Call AddBundleSubform_Exit(True)
Me.Editor.SetFocus
End If
End Sub

Private Sub AddBundleSubform_Exit(Cancel As Integer)
'do nothing, just a placeholder
End Sub


'validate articles in subform to ensure all data
'is entered
Private Function validateArticles()
Dim bgexists As Integer
Dim blankfields As Integer

bgexists = 0
blankfields = 0

Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Dim sqlstr As String
sqlstr = "SELECT * FROM Articles WHERE Articles.fk_topic = " & Nz(Me.ID, 0) & ";"
rst.Open sqlstr, CurrentProject.Connection


Do While Not rst.EOF

    'check that the bundle has a BG article
    If rst![Article Types].Value = 7 Then
    bgexists = 1
    End If
    
    'check that all records have a due date
    If IsNull(rst![Due Date].Value) Then
    blankfields = 1
    End If
    
    'check that all records have a pub date
    If IsNull(rst![Target].Value) Then
    blankfields = 1
    End If
    
    'check that all records have an author
    If IsNull(rst![Author].Value) Then
    blankfields = 1
    End If
    
    'check that all records have a status
    If IsNull(rst![Status].Value) Then
    blankfields = 1
    End If
    
    
    If Not rst.EOF Then rst.MoveNext
Loop

'if any of the exit flags above have been set, return 1
If blankfields = 1 Then
    MsgBox "Error: Missing information. Please fill out all information for your articles.", , "Error!"
    Me.AddBundleSubform.SetFocus
    validateArticles = 1
ElseIf bgexists = 0 Then
    MsgBox "Please enter a Buyer's Guide article!", , "Error!"
    Me.AddBundleSubform.SetFocus
    validateArticles = 1
End If

End Function


'if the form is closed in any other way, discard changes
Private Sub Form_Unload(Cancel As Integer)

'if no bundle has been created then just close the form
If IsNull(Me.ID) = True Or IsEmpty(Me.ID) = True Then
GoTo endsub
Else
'else get the user's permission before deleting all data associated with this series
'and then close the form
If MsgBox("Closing the form will discard your changes. Do you want to proceed?", vbYesNo, "Warning") = vbYes Then
Dim cancelAddBundle As String
Dim removeArticles As String

removeArticles = "DELETE * FROM Articles WHERE Articles.fk_topic=" & Me.ID & ";"
cancelAddBundle = "DELETE * FROM Bundle WHERE Bundle.ID=" & Me.ID & ";"

'because table has no ON DELETE CASCADE setting,
'remove the articles first and then the series
DoCmd.SetWarnings False
DoCmd.RunSQL removeArticles
DoCmd.RunSQL cancelAddBundle
DoCmd.SetWarnings True

'if no is pressed, cancel the unload operation
Else
Cancel = True
End If
End If
endsub:
End Sub

Private Sub FormHeader_Click()
Call RestoreMain
End Sub


'Autopopulate the media group after site has been updated
Private Sub Site_AfterUpdate()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim MG_ID As Integer
Dim siteSql As String

siteSql = "SELECT Sites.[Media Group] FROM Sites WHERE Sites.ID = " & Nz(Me.Site.Value, 0) & ";"

rs.Open siteSql, CurrentProject.Connection
MG_ID = rs![Media Group].Value
Me.Media_Group.Value = MG_ID
Me.Media_Group.Requery
rs.Close
End Sub

'Add the new series to the database
Private Sub Addbutton_Click()

'If any of the bundle fields are null, cancel the add operation
If (IsNull(Me.Editor) = True Or IsNull(Me.Site) = True Or _
IsNull(Me.Media_Group) = True Or IsNull(Me.Bundle_Topic) = True Or _
IsNull(Me.Bundle_Price) = True) Then
MsgBox "Please fill out all information before adding.", vbOKOnly, "Error"
GoTo Addbutton_Click_Exit
End If

On Error GoTo Addbutton_Click_Err

    On Error Resume Next
    
    'if validate articles returns 1, then article fields are blank
    'cancel add operation and go to end of sub
    If validateArticles > 0 Then GoTo Addbutton_Click_Exit
    
    'add the new record
    DoCmd.GoToRecord , "", acNewRec
    
    
    'error handler
    If (MacroError <> 0) Then
        Beep
        MsgBox MacroError.description, vbOKOnly, ""
    Else
    
    'display success message and close
    DoCmd.Close acForm, "New Bundle"
    MsgBox "Bundle added successfully!", vbOKOnly, "Success!"
    End If
    


Addbutton_Click_Exit:
    Exit Sub

Addbutton_Click_Err:
    MsgBox Error$
    Resume Addbutton_Click_Exit

End Sub

