VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InvoiceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'Checks to ensure that the value entered for amount is a number
Private Sub amt_AfterUpdate()
Dim amount As String
amount = Me.amt

'if Me.amt is not numeric, reject value
'and display warning
If IsNumeric(Me.amt) = False Then
MsgBox ("Error: Invoice amount must be a number.")
Me.amt = Null
Me.approval.SetFocus
Me.amt.SetFocus
End If

'If me.amt contains a comma or a dollar sign,
'remove them before proceeding
If (InStr(Me.amt, ",") > 0) Or (InStr(Me.amt, "$") > 0) Then
amount = Replace(amount, ",", "")
amount = Replace(amount, "$", "")
End If

Me.amt = amount
End Sub


Private Sub Auto_Header0_Click()
Call RestoreMain
End Sub
'Autoset "month to go" as date's month
Private Sub date_AfterUpdate()
Me.monthtogo = MonthName(DatePart("m", Me.date))
End Sub
'File the invoice
Private Sub fileInvoice_Click()
Dim recipient As String
Dim subject As String
Dim msgbody As String
Dim attachment As String
Dim cc As String
Dim updateSql As String
'If fields are blank, end call without doing antyhing
If IsNull(Me.invoicenum) = True Or IsNull(Me.amt) = True Then
MsgBox ("Error: Please fill out all fields.")
GoTo endcall
End If

'If no invoice is attached, end call without doing anything
If IsNull(Me.docpath) = True Then
MsgBox ("Error: Please attach an invoice.")
GoTo endcall
End If




'Generate the invoice email and close the form
Call generateEmail(recipient, cc, subject, msgbody, attachment)
Call sendEmail(recipient, cc, subject, msgbody, attachment)

'Update tables and requery the search form
Call generateSql(updateSql)
Call Forms("SearchForm").invoiceTable(updateSql)

DoCmd.Close acForm, "InvoiceForm"
endcall:

End Sub

'Find a file to attach as invoice
Private Sub findFile_Click()

Dim f As Object

Set f = Application.FileDialog(3)

'Selecting multiple files not allowed
f.AllowMultiSelect = False

If f.Show Then
    For i = 1 To f.SelectedItems.Count
        sFile = Filename(f.SelectedItems(i), sPath)
        Me.docpath = sPath & sFile
        
    Next
End If

End Sub

Private Sub Form_Load()

'Ensure search form is loaded
If CurrentProject.AllForms("SearchForm").IsLoaded = False Then
    MsgBox ("Error: Cannot open invoice form when Search is closed")
    DoCmd.Close acForm, "InvoiceForm"
    GoTo closeform
End If



'Get ID of article to invoice
Me.ID.Value = Me.OpenArgs

'autopopulate site, payee, and "month to go" fields
Call getSite
Me.payee.Value = Form_SearchForm.Text54
Me.monthtogo = MonthName(DatePart("m", Me.date))

closeform:

End Sub

'Fetch site based based on site id from search form
Private Function getSite()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sqlstr As String

sqlstr = "SELECT Sites.Site FROM Sites WHERE Sites.ID = " & Form_SearchForm.Text66
rs.Open sqlstr, CurrentProject.Connection

Do While Not rs.EOF
Me.Site = rs![Site].Value
If Not rs.EOF Then rs.MoveNext
Loop


End Function

'Generate the text of the invoice email
Private Function generateEmail(recipient As String, cc As String, subject As String, msgbody As String, attachment As String)
recipient = "accountspayable@techtarget.com"
cc = "cgagne@techtarget.com"
subject = "Centralized Editorial, PRC (Product Content), Invoice - " & Me.payee

msgbody = "Title: " & Me.Title & vbCrLf & _
            "Content Type: " & Me.contenttype & vbCrLf & _
            "Site Name: " & Me.Site & vbCrLf & _
            "Date Submitted: " & Me.date & vbCrLf & _
            "Month going toward: " & Me.monthtogo & vbCrLf & vbCrLf & _
            "Contributor: " & Me.payee & vbCrLf & _
            "Invoice Amount: " & FormatCurrency(Me.amt, 0, False, False, True) & vbCrLf & _
            "Invoice Number: " & Me.invoicenum & vbCrLf & _
            "Budget Type: " & Me.budget & vbCrLf & vbCrLf & _
            "Approval Status: " & Me.approval & vbCrLf & _
            "Quarter: Q" & DatePart("q", Me.date) & " " & DatePart("yyyy", Me.date) & vbCrLf & _
            "Content Description (Accounts Payable):" & vbCrLf & _
            "See attached invoice" & vbCrLf & vbCrLf & _
            GBL_Fname & " " & GBL_Lname & vbCrLf & _
            GBL_Title
            
attachment = Me.docpath
End Function

'Update the tables with the invoice ID, amount and date values
Public Function generateSql(updateSql As String)

Dim datesub As Date
Dim invnum As String

invnum = Me.invoicenum
datesub = DateValue(Me.date)
updateSql = "UPDATE Articles" & _
            " SET Articles.[Invoice Amount] = " & Me.amt & ", Articles.[Invoice ID] = '" & invnum & "', [Invoice Submitted] = #" & datesub & "# WHERE ID = " & Me.OpenArgs & ";"

End Function




Private Sub FormHeader_Click()
Call RestoreMain
End Sub


