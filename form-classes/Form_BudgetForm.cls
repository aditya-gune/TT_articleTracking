VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BudgetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public warning As Integer
Public currentHeader As Integer
Public myBudget As Long
'Public amtSpent As Integer
Public myBudgetImpact As Long

Private Sub BudgetSubform_Exit(Cancel As Integer)
Call populateBoxes
End Sub

'If superuser, export as admin. Else, export as regular user
Private Sub Command44_Click()
If GBL_UserID = 88 Then
Call adminExcel
Else
Call exportExcel
End If
End Sub

'Add a new record to budget plan
Private Sub Command52_Click()
Dim insertSql As String

insertSql = "INSERT INTO Budget_details (Title, header_id, amount, site, quarter, priority) " & _
            "VALUES ('', " & currentHeader & ", 0, '', '', '');"
DoCmd.SetWarnings False
DoCmd.RunSQL insertSql
DoCmd.SetWarnings True
Me.BudgetSubform.Requery
Me.BudgetSubform.SetFocus
Call Me.BudgetSubform.Form.goToNewRec

End Sub

'Delete the selected record from budget plan
Private Sub Command54_Click()
Dim deletesql As String

deletesql = "DELETE Budget_details.* FROM Budget_details WHERE Budget_details.ID = " & Me.selectedrecord.Value
DoCmd.SetWarnings False
DoCmd.RunSQL deletesql
DoCmd.SetWarnings True
Me.BudgetSubform.Requery
Call populateBoxes
Call populateQBoxes

End Sub


Private Sub Command56_Click()
DoCmd.OpenForm "ChangeBudget", acNormal
End Sub

Private Sub Form_Load()

'On form load, if superuser, load as admin
If GBL_UserID = 88 Then
Call adminLoader
GoTo endcall
End If

'Else, load existing series, budget plan, and
'calculate amounts
Dim amtSpent As Integer
Call getCurrentHeader
Call getMyBudget

Call loadExisting
Call loadBudgetSubform
Call updateBudgetImpact
Call getAmtSpent(amtSpent)

Call populateBoxes
Call populateQBoxes
Call checkDuplicates

endcall:

End Sub

'Get the current budget header
Public Function getCurrentHeader()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Dim idSql As String

Dim curY As Integer
curY = DatePart("yyyy", date)

idSql = "SELECT Months.id FROM Months WHERE Months.editor_ID = " & GBL_EditorID & " AND Months.Year = " & curY & ";"
rst.Open idSql, CurrentProject.Connection
Do While Not rst.EOF
    currentHeader = Nz(rst![ID].Value, 0)
If Not rst.EOF Then rst.MoveNext
Loop
End Function

'Use the current budget header to get the budget plan
Private Function getMyBudget()
Dim rec As ADODB.Recordset
Set rec = New ADODB.Recordset
Dim idSql As String

Dim curQ As Integer
Dim curY As Integer
curQ = DatePart("q", date)
curY = DatePart("yyyy", date)

budgetSql = "SELECT Months.amount FROM Months WHERE Months.ID = " & currentHeader & ";"
rec.Open budgetSql, CurrentProject.Connection
Do While Not rec.EOF
    myBudget = Nz(rec![amount].Value, 0)
If Not rec.EOF Then rec.MoveNext
Loop
End Function

'Use globals to getthe amount of budget that has already been spent
Private Function getAmtSpent(amtSpent As Integer)
Dim sqlstr As String

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

curY = DatePart("yyyy", date)

sqlstr = "SELECT Sum([Bundle].[Bundle Price]) AS spent FROM Bundle " & _
"WHERE (((DatePart('yyyy',[Bundle].[Created]))=" & curY & ") AND ((Bundle.Editor)=" & GBL_EditorID & "));"

rs.Open sqlstr, CurrentProject.Connection

Do While Not rs.EOF
  amtSpent = Nz(rs![spent].Value, 0)
    
If Not rs.EOF Then rs.MoveNext
Loop


End Function

'Load existing series
Private Function loadExisting()
Dim curY As Integer
curY = DatePart("yyyy", date)

Me.AssignedSeries.SourceObject = "Query.adminBudgetQuery"
Me.AssignedSeries.Form.RecordSource = "SELECT Bundle.Editor, Bundle.[Bundle Topic], Bundle.Site, Bundle.[Bundle Price] FROM Bundle WHERE ((DatePart('yyyy',[Bundle].[Created]))=" & curY & " AND Bundle.editor = " & GBL_EditorID & ");"
Me.AssignedSeries.Requery

'set subform permissions
Me.AssignedSeries.Form.AllowAdditions = False
Me.AssignedSeries.Form.AllowEdits = False
Me.AssignedSeries.Form.AllowDeletions = False
Me.AssignedSeries.Form.NavigationButtons = False
Me.AssignedSeries.Form.ShortcutMenuBar = False


End Function

'load the budget plan
Private Function loadBudgetSubform()

Dim pSql As String
Dim filterstr As String

pSql = "SELECT Budget_details.Title, Budget_details.site, Budget_details.amount, Budget_details.quarter, Budget_details.priority, Budget_details.ID, Budget_details.header_id" & _
        " FROM Budget_details" & _
        " WHERE (((Budget_details.header_id)=" & currentHeader & "));"

filterstr = "[header_id] = " & currentHeader

Me.BudgetSubform.Requery
Me.BudgetSubform.Form.filter = filterstr
Me.BudgetSubform.Form.FilterOn = True

'hide the budget header column, which is auto added for new records
For Each ctl In Me.BudgetSubform.Form.Controls
    If (ctl.name = "Budget Item") Then
            ctl.ColumnHidden = True
    End If
Next ctl

'set subform permissions
Me.BudgetSubform.Form.AllowAdditions = False
Me.BudgetSubform.Form.AllowEdits = True
Me.BudgetSubform.Form.AllowDeletions = True
Me.BudgetSubform.Form.NavigationButtons = False


End Function
'refresh boxes
Private Sub Command5_Click()
If GBL_UserID = 88 Then
Call adminLoader
GoTo endcall
End If

Me.BudgetSubform.Requery
Call populateBoxes
Call populateQBoxes
endcall:
End Sub

'update a line item's impact on the budget
Private Function updateBudgetImpact()
Dim sumrs As ADODB.Recordset
Set sumrs = New ADODB.Recordset

Dim sqlstr As String

sqlstr = "SELECT SUM(Budget_details.amount) AS projected" & _
        " FROM Budget_details" & _
        " WHERE (((Budget_details.header_id)=" & currentHeader & "));"
   
sumrs.Open sqlstr, CurrentProject.Connection

Do While Not sumrs.EOF

myBudgetImpact = Nz(sumrs![projected].Value, 0)
If Not sumrs.EOF Then sumrs.MoveNext
Loop
End Function
 
'populate the amount boxes
Public Sub populateBoxes()
Dim amtSpent As Integer
Call getAmtSpent(amtSpent)
Call updateBudgetImpact
Me.myTotalBudget.Value = FormatCurrency(myBudget, 0, vbFalse, vbFalse, vbTrue)
Me.amtRemaining.Value = FormatCurrency(myBudget - (amtSpent + myBudgetImpact), 0, vbFalse, vbFalse, vbTrue)
Me.amtSpent.Value = FormatCurrency(amtSpent + myBudgetImpact, 0, vbFalse, vbFalse, vbTrue)
End Sub

'populate the quarter boxes
Public Sub populateQBoxes()
Dim n As Integer

Dim qArray(4) As Integer
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset

Dim qSql As String

For i = 1 To 4

qSql = "SELECT SUM(Budget_details.amount) AS q" & _
        " FROM Budget_details" & _
        " WHERE (((Budget_details.header_id)=" & currentHeader & ") AND " & _
        "(Budget_details.quarter = " & i & "));"
rst.Open qSql, CurrentProject.Connection

'fill an array with the results of rs
Do While Not rst.EOF
qArray(i) = Nz(rst![Q].Value, 0)
If Not rst.EOF Then rst.MoveNext
Loop
rst.Close
Next i

'populate the quarterly boxes from the array
Me.Q1box = FormatCurrency((Nz(Me.myBudget, 0) / 4) - qArray(1), 0, vbFalse, vbFalse, vbTrue)
Me.Q2box = FormatCurrency((Nz(Me.myBudget, 0) / 4) - qArray(2), 0, vbFalse, vbFalse, vbTrue)
Me.Q3box = FormatCurrency((Nz(Me.myBudget, 0) / 4) - qArray(3), 0, vbFalse, vbFalse, vbTrue)
Me.Q4box = FormatCurrency((Nz(Me.myBudget, 0) / 4) - qArray(4), 0, vbFalse, vbFalse, vbTrue)


End Sub

Private Sub Form_Timer()
Call checkDuplicates
Call populateQBoxes
End Sub

Private Sub FormHeader_Click()
Call RestoreMain
End Sub

Private Sub exportExcel()

    Dim rst As ADODB.Recordset              'create the adodb recordset
    Set rst = New ADODB.Recordset
    Dim qSql As String                      'create the sql string
    Dim results() As String                 'create the dynamic array to hold results
    Dim r As Integer                        'r holds the recordcount
    Dim newExcelApp As Excel.Application    'create the excel application
    Dim newWbk As Excel.Workbook            'create the excel wkbook
    Dim newWkSheet As Excel.Worksheet       'create the excel wksht
    Dim str As String
    Set newExcelApp = Excel.Application
    newExcelApp.Visible = True
    Set newWbk = newExcelApp.Workbooks.add
    Set newWkSheet = newWbk.Worksheets(1)


qSql = "SELECT Budget_details.Title, Sites.site, Budget_details.amount, Budget_details.quarter, Budget_details.priority" & _
        " FROM Budget_details INNER JOIN Sites ON Budget_details.site = Sites.ID" & _
        " WHERE (((Budget_details.header_id)=" & currentHeader & "));"
rst.Open qSql, CurrentProject.Connection
r = 0

'get recordcount
Do While Not rst.EOF
r = r + 1
If Not rst.EOF Then rst.MoveNext
Loop

'redim the array to size of results
ReDim results(5, r + 1)

If Not rst.BOF Then rst.MoveFirst

'populate results() with rs values
For j = 2 To r + 1

    results(1, j) = rst![Title].Value
    results(2, j) = rst![Site].Value
    results(3, j) = rst![amount].Value
    results(4, j) = rst![quarter].Value
    results(5, j) = rst![priority].Value
    rst.MoveNext
Next j
        newWkSheet.Cells(1, 1).Value = "TITLE"
        newWkSheet.Cells(1, 2).Value = "SITE"
        newWkSheet.Cells(1, 3).Value = "AMOUNT"
        newWkSheet.Cells(1, 4).Value = "QUARTER"
        newWkSheet.Cells(1, 5).Value = "PRIORITY"

'populate worksheet with results() values
For j = 1 To r + 1
    For i = 1 To 5

        newWkSheet.Cells(j + 1, i).Value = results(i, j)
    
    Next i

Next j

End Sub

'check budget plan and existing series for duplicate items
Public Function checkDuplicates()
Dim curSql As String
Dim planSql As String
Dim curRs As ADODB.Recordset
Set curRs = New ADODB.Recordset
Dim str As String
Dim planRs As ADODB.Recordset
Set planRs = New ADODB.Recordset
Dim numCur As Integer               'holds the recordcount
Dim i As Integer

curSql = "SELECT Bundle.[Bundle Topic] FROM Bundle WHERE ((DatePart('yyyy',[Bundle].[Created]))=" & DatePart("yyyy", date) & ") AND Bundle.Editor = " & GBL_EditorID & ";"
planSql = "SELECT Budget_details.Title FROM Budget_details" & _
        " WHERE (((Budget_details.header_id)=" & Me.currentHeader & "));"

curRs.Open curSql, CurrentProject.Connection

'get the recordcount
Do While Not curRs.EOF
    numCur = numCur + 1
If Not curRs.EOF Then curRs.MoveNext
Loop

'declare a dynamic array and redim it to recordcount size
Dim currArray() As String
ReDim currArray(numCur)
i = 0

'if there was some error with the rs that resulted in numCur < 1, go to endcall
If numCur < 1 Then
GoTo endfunction
End If


'fill array with rs values
curRs.MoveFirst
Do While Not curRs.EOF
    currArray(i) = curRs![Bundle Topic].Value
    str = str & currArray(i) & vbCrLf
If Not curRs.EOF Then curRs.MoveNext
i = i + 1
Loop

curRs.Close

'reopen rs with planSql
str = ""
planRs.Open planSql, CurrentProject.Connection
Do While Not planRs.EOF
    For i = 0 To numCur
    
        'compare the strings, if they are similar, show warning
        If InStr(currArray(i), planRs![Title].Value) Then
            If warning <> 1 Then
                MsgBox ("You may have some entries in your budget plan that have already been assigned. Please check your plan and remove any duplicates.")
                warning = 1
            End If
        End If

    Next i
If Not planRs.EOF Then planRs.MoveNext
Loop
endfunction:
End Function

'loader for admin users (in place of overloaded form_load() function)
Public Function adminLoader()
Dim amtSpent As Long
Dim curY As Integer
Dim pSql As String
Dim filterstr As String
curY = DatePart("yyyy", date)

Me.AssignedSeries.SourceObject = "Query.adminBudgetQuery"
Me.AssignedSeries.Form.RecordSource = "SELECT Bundle.Editor, Bundle.[Bundle Topic], Bundle.Site, Bundle.[Bundle Price] FROM Bundle WHERE ((DatePart('yyyy',[Bundle].[Created]))=" & curY & ");"
Me.AssignedSeries.Requery

Me.AssignedSeries.Form.AllowAdditions = False
Me.AssignedSeries.Form.AllowEdits = False
Me.AssignedSeries.Form.AllowDeletions = False
Me.AssignedSeries.Form.NavigationButtons = False
Me.AssignedSeries.Form.ShortcutMenuBar = False
Me.Label42.Caption = "Series Assigned in 2016"


pSql = "SELECT Budget_details.Title, Budget_details.site, Budget_details.amount, Budget_details.quarter, Budget_details.priority, Budget_details.ID, Budget_details.header_id" & _
        " FROM Budget_details" & _
        " WHERE (((Budget_details.header_id)=" & currentHeader & "));"

Me.BudgetSubform.Requery
Me.BudgetSubform.Width = (1440 * 7.3333)
Me.Command52.Visible = False
Me.Command54.Visible = False

For Each ctl In Me.BudgetSubform.Form.Controls
    If (ctl.name = "Budget Item") Then
            ctl.ColumnHidden = False
    End If
Next ctl

Me.Label43.Caption = "Series Planned in 2016"
Me.BudgetSubform.Form.AllowAdditions = False
Me.BudgetSubform.Form.AllowEdits = False
Me.BudgetSubform.Form.AllowDeletions = False
Me.BudgetSubform.Form.NavigationButtons = False

Me.Label7.Caption = "Annual Budget"

Call getTeamBudget
Call getTeamCommitted(amtSpent)

Me.myTotalBudget.Value = FormatCurrency(myBudget, 0, vbFalse, vbFalse, vbTrue)
Me.amtRemaining.Value = FormatCurrency((myBudget - (amtSpent + myBudgetImpact)), 0, vbFalse, vbFalse, vbTrue)
Me.amtSpent.Value = FormatCurrency(amtSpent + myBudgetImpact, 0, vbFalse, vbFalse, vbTrue)


Me.Label22.Visible = False
Me.Label24.Visible = False
Me.Label26.Visible = False
Me.Label28.Visible = False
Me.Q1box.Visible = False
Me.Q2box.Visible = False
Me.Q3box.Visible = False
Me.Q4box.Visible = False
Me.Box55.Visible = False
Me.TimerInterval = 0
End Function

Private Sub getTeamBudget()
Dim rsB As ADODB.Recordset
Set rsB = New ADODB.Recordset
Dim budgetSql As String


Dim curQ As Integer
Dim curY As Integer
curQ = DatePart("q", date)
curY = DatePart("yyyy", date)

budgetSql = "SELECT Months.amount FROM Months WHERE Months.Year = " & DatePart("yyyy", date) & ";"
rsB.Open budgetSql, CurrentProject.Connection
Do While Not rsB.EOF

    myBudget = myBudget + Nz(rsB![amount].Value, 0)

If Not rsB.EOF Then rsB.MoveNext
Loop


End Sub

Private Sub getTeamCommitted(amtSpent As Long)
Dim sqlstr As String

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

curY = DatePart("yyyy", date)

sqlstr = "SELECT Sum([Bundle].[Bundle Price]) AS spent FROM Bundle " & _
"WHERE (((DatePart('yyyy',[Bundle].[Created]))=" & curY & "));"
'MsgBox ("curY = " & curY & vbCrLf & "GBL_EID = " & GBL_EditorID)

rs.Open sqlstr, CurrentProject.Connection

Do While Not rs.EOF
  amtSpent = Nz(rs![spent].Value, 0)
    
If Not rs.EOF Then rs.MoveNext
Loop

End Sub

Private Sub adminExcel()

    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    Dim qSql As String
    Dim results() As String
    Dim r As Integer
    Dim newExcelApp As Excel.Application
    Dim newWbk As Excel.Workbook
    Dim newWkSheet As Excel.Worksheet
    Dim str As String
    Set newExcelApp = Excel.Application
    newExcelApp.Visible = True
    Set newWbk = newExcelApp.Workbooks.add
    Set newWkSheet = newWbk.Worksheets(1)


qSql = "SELECT Budget_details.Title, Months.Description, Budget_details.amount, Sites.Site, Budget_details.quarter, Budget_details.priority" & _
        " FROM Sites INNER JOIN (Months INNER JOIN Budget_details ON Months.ID = Budget_details.header_id) ON Sites.ID = Budget_details.site" & _
        " WHERE (((Months.year)=" & DatePart("yyyy", date) & "));"
rst.Open qSql, CurrentProject.Connection
r = 0
Do While Not rst.EOF
r = r + 1
If Not rst.EOF Then rst.MoveNext
Loop

'MsgBox ("rows: " & r)
ReDim results(6, r + 1)

If Not rst.BOF Then rst.MoveFirst

For j = 2 To r + 1

    results(1, j) = rst![Title].Value
    results(2, j) = rst![description].Value
    results(3, j) = rst![Site].Value
    results(4, j) = rst![amount].Value
    results(5, j) = rst![quarter].Value
    results(6, j) = rst![priority].Value
    rst.MoveNext
Next j
        newWkSheet.Cells(1, 1).Value = "TITLE"
         newWkSheet.Cells(1, 2).Value = "BUDGET ITEM"
        newWkSheet.Cells(1, 3).Value = "SITE"
        newWkSheet.Cells(1, 4).Value = "AMOUNT"
        newWkSheet.Cells(1, 5).Value = "QUARTER"
        newWkSheet.Cells(1, 6).Value = "PRIORITY"
For j = 1 To r + 1
    For i = 1 To 6

        newWkSheet.Cells(j + 1, i).Value = results(i, j)
    
    Next i

Next j

End Sub

