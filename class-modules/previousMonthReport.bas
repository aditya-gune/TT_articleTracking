Attribute VB_Name = "previousMonthReport"
Option Compare Database
'Constructs an ad-hoc query for to generate report
'of what was published in the preceding month
Public Function getReport()
Dim monthStart As Date
Dim monthEnd As Date
Dim nameStr As String
Dim userSql As String
Dim qd As QueryDef

'get dates
Call getDates(monthStart, monthEnd, nameStr)

'construct the query string
Call constructQuery(monthStart, monthEnd, userSql)


'clear any previous data from the query
On Error Resume Next
CurrentDb.QueryDefs.Delete nameStr
On Error GoTo 0

'define the query using the SQL created above
Set qd = CurrentDb.CreateQueryDef(nameStr, userSql)

'Open the query
DoCmd.OpenQuery nameStr, acViewNormal, acReadOnly

'Delete the query after creation to save space
On Error Resume Next
CurrentDb.QueryDefs.Delete nameStr
On Error GoTo 0


End Function
Private Sub constructQuery(monthStart As Date, monthEnd As Date, userSql As String)

userSql = "SELECT Bundle.Editor, Bundle.[Bundle Topic], Articles.[Article Types], Articles.Subtype, Bundle.Site, Articles.Target, Articles.URL " & _
            "FROM Bundle INNER JOIN Articles ON Bundle.ID = Articles.fk_topic " & _
            "WHERE (((Articles.Target) >= #" & monthStart & "# And (Articles.Target) <= #" & monthEnd & "#) And ((Articles.URL) Is Not Null)) " & _
            "ORDER BY Articles.Target, Bundle.ID, Articles.[Article Types];"

End Sub


Private Sub getDates(monthStart As Date, monthEnd As Date, nameStr As String)
Dim today As Date
Dim monthvar As Integer
Dim yearvar As Integer

today = date

'If the current month is not January, then
'monthVar equals current month minus one
'yearVar equals current year
If (Month(today) <> 1) Then
monthvar = (DatePart("m", today) - 1)
yearvar = (DatePart("yyyy", today))

'If current month is January, then
'monthVar equals 12
'yearVar equals current year minus one
Else
monthvar = 12
yearvar = (DatePart("yyyy", today) - 1)

End If

'Cat the numbers together to assemble
'the strings for the month beginning and end
monthStart = monthvar & "/01/" & yearvar
monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)

'create the name for the report
nameStr = "Monthly Report: " & MonthName(monthvar, False) & " " & yearvar

End Sub

