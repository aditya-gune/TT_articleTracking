Attribute VB_Name = "SeriesUpdateModule"
Option Compare Database

'Check whether the date is valid for a notification
'should only proc on first 10 days of the month
Private Function checkDate()

 If DatePart("d", date) <= 10 Then
 checkDate = 1
 Else
 checkDate = 0
 End If
End Function

Private Function queryDB()
Dim sql As String
Dim testsql As String
Dim targetmonth As Date
Dim targetstart As Date
Dim targetend As Date
Dim topicsFound As String
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
   
targetmonth = DateAdd("M", -9, date)
targetstart = DateSerial(Year(targetmonth), Month(targetmonth), 1)
targetend = DateSerial(Year(targetmonth), Month(targetmonth) + 1, 0)

sql = "SELECT DISTINCT Bundle.[Bundle Topic]" & _
    " FROM Bundle INNER JOIN Articles ON Bundle.ID = Articles.fk_topic" & _
    " WHERE (((Bundle.Editor)= " & GBL_EditorID & ") AND ((Articles.[Article Types])=4)" & _
    " AND ((Articles.Status)=7)" & _
    " AND ((Articles.Target)>=#" & targetstart & "# AND ((Articles.Target)<=#" & targetend & "#)));"
    
testsql = "SELECT DISTINCT Bundle.[Bundle Topic]" & _
    " FROM Bundle INNER JOIN Articles ON Bundle.ID = Articles.fk_topic" & _
    " WHERE (((Bundle.Editor)=1) AND ((Articles.[Article Types])=4)" & _
    " AND ((Articles.Status)=7)" & _
    " AND ((Articles.Target)>=#" & targetstart & "# AND ((Articles.Target)<=#" & targetend & "#)));"
 
rst.Open sql, CurrentProject.Connection

Do While Not rst.EOF

If IsNull(rst![Bundle Topic].Value) = True Or rst![Bundle Topic].Value = "" Then
GoTo NullString
Else
topicsFound = topicsFound & rst![Bundle Topic].Value & vbCrLf
End If

If Not rst.EOF Then rst.MoveNext
Loop


queryDB = topicsFound
GoTo endcall

NullString:
queryDB = Null

endcall:

End Function


Public Function seriesUpdate()


Dim updates As String

If checkDate > 0 Then
updates = queryDB
    If IsNull(updates) = False And (updates <> "") Then
    If IsMuted < 1 Then
        seriesUpdate = updates
        DoCmd.OpenForm "SeriesUpdateForm"
        DoCmd.SelectObject acForm, "Main Menu"
        DoCmd.Minimize
        Forms("SeriesUpdateForm").SetFocus
    Else
    'do nothing
    End If
    
    Else
    'do nothing
    End If
    
Else
'do nothing
End If


End Function
