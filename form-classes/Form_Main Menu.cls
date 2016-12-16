VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Main Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
'Check to see if the invoice reporting form is open
'if it is, close it before opening the search form
Public Function TblArticleConflict() As Integer
If (CurrentProject.AllForms("InvoiceReporting1").IsLoaded = True) Then
    DoCmd.Close acForm, "InvoiceReporting1"
End If
End Function


Private Sub AllSeriesQuery_Click()
DoCmd.OpenQuery "BundleQuery"
End Sub

Private Sub ArticlesBevansButton_Click()
Me.Text123 = 3
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"

End Sub

Private Sub ArticlesLandsbergButton_Click()
Me.Text123 = 11
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"
End Sub

Private Sub ArticlesGuneButton_Click()
Me.Text123 = 5
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"

End Sub

Private Sub ArticlesHoupesButton_Click()
Me.Text123 = 6
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"

End Sub

Private Sub ArticlesMillerButton_Click()
Me.Text123 = 1
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"
End Sub

Private Sub ArticlesNovotnyButton_Click()
Me.Text123 = 2
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"
End Sub

Private Sub checkGlobals()
Dim userInfo As String
userInfo = "Username: " & GBL_Username & vbCrLf _
        & "Name: " & GBL_Name & vbCrLf _
        & "User ID: " & GBL_UserID & vbCrLf _
        & "Editor ID: " & GBL_EditorID & vbCrLf _
        & "First name: " & GBL_Fname & vbCrLf _
        & "Last name: " & GBL_Lname & vbCrLf _
        & "Title: " & GBL_Title & vbCrLf
        
MsgBox (userInfo)
End Sub

Private Sub CustomQueryBtn_Click()
DoCmd.Minimize
DoCmd.RunCommand acCmdNewObjectDesignQuery
End Sub

'Check to see if the search form is loaded before opening the invoice
'form. If it is, prompt use to close it before Search opens
Private Sub Reporting_Click()
If (CurrentProject.AllForms("SearchForm").IsLoaded = True) Then
    If MsgBox("Opening the Reporting form will close the Search Form. Continue?", vbOKCancel, "Warning") = vbOK Then
        DoCmd.Close acForm, "SearchForm"
        DoCmd.OpenForm "InvoiceReporting1"
    End If
Else
DoCmd.OpenForm "InvoiceReporting1"
End If
End Sub


'User permissions check before opening the superuser panel
Private Sub Superuserbtn_Click()
If GBL_UserID = 88 Or GBL_UserID = 1897 Or GBL_UserID = 1826 Then
    
    'if the form is already loaded then
    'bring it to the front
    If CurrentProject.AllForms("Superuser Panel").IsLoaded Then
        DoCmd.OpenForm "Superuser Panel"
    Else
        'else, open login form
        DoCmd.OpenForm "SuperuserAccess"
    End If
Else
    MsgBox ("Access Denied!")
End If
End Sub

Private Sub PreviousMonthQuery_Click()
Call getReport
End Sub


Private Sub BudgetMgmt_Click()
Call TblArticleConflict
DoCmd.OpenForm "BudgetForm", acNormal
End Sub


Private Sub CompletedGuidesQuery_Click()
DoCmd.OpenQuery "PublishedGuides"
End Sub

Private Sub Detail_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Shift > 2 Then
Call checkGlobals
ElseIf Shift > 0 And Shift < 3 Then
If MsgBox("Does this unit have a soul?", vbYesNo) = vbYes Then 'do nothing
End If
End If
End Sub

Private Sub Form_Close()
Call logOut
End Sub

Private Sub Form_Load()

'add trusted locations
Call AddTrustedLocation

'get userdata and log in
Call userFunction
Call logIn

'if user has superuser rights, show navpane and superuser panel button
If GBL_UserID = 1897 Or GBL_UserID = 1826 Then
Superuserbtn.Visible = True
CustomQueryBtn.Visible = True
Else
'else, hide them
Superuserbtn.Visible = False
CustomQueryBtn.Visible = False
DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide
End If

'notify user of series updates
Call seriesUpdate
Call UnmuteNotifications



End Sub

Private Sub Form_Timer()
If GBL_EditorID < 1 And GBL_UserID < 1 Then
Call userFunction
End If
End Sub

Private Sub SearchBtn_Click()
Me.Text123 = Null
Call TblArticleConflict
DoCmd.OpenForm "SearchForm"
End Sub

