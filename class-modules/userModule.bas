Attribute VB_Name = "userModule"
Option Compare Database
Option Explicit
Global GBL_Username As String
Global GBL_Name As String
Global GBL_UserID As Integer
Global GBL_EditorID As Integer
Global GBL_Fname As String
Global GBL_Lname As String
Global GBL_Title As String




Public Function userFunction()
Dim userName As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim userSql As String
Dim userID As Integer
Dim idbox As Control
Dim lname As String
Set idbox = [Forms]![Main Menu]![idbox]
Dim fnamebox As Control
Set fnamebox = [Forms]![Main Menu]![fnamebox]



userName = "agune"
[Forms]![Main Menu]![usernamebox] = userName
userSql = "SELECT UserInfo.[ID], UserInfo.[Last name], UserInfo.[Name], UserInfo.[First name], UserInfo.[Title] FROM UserInfo WHERE UserInfo.[User name] = '" & userName & "';"
rs.Open userSql, CurrentProject.Connection

'MsgBox (userName)
Do While Not rs.EOF
    'do nothing
'[Forms]![Main Menu]![usernamebox] = userName
idbox = rs![ID].Value
fnamebox = "Welcome, " & Nz(rs![First name].Value, "User")

GBL_Name = Nz(rs![name].Value, 0)
GBL_UserID = Nz(rs![ID].Value, 0)
GBL_Username = userName
GBL_Lname = Nz(rs![Last name].Value, 0)
GBL_Fname = Nz(rs![First name].Value, 0)
GBL_Title = Nz(rs![Title].Value, "Product Content Team")
lname = Nz(rs![Last name].Value, 0)

If Not rs.EOF Then rs.MoveNext
Loop

Call getEditorID(lname)

End Function

Public Function getEditorID(lname As String)
Dim userSql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
userSql = "SELECT Editors.[ID] FROM Editors WHERE Editors.Editor LIKE '" & lname & "';"
rs.Open userSql, CurrentProject.Connection
Do While Not rs.EOF
GBL_EditorID = Nz(rs!ID.Value, 0)
If Not rs.EOF Then rs.MoveNext
Loop
'MsgBox ("success")
End Function

Public Function logIn()
Dim loginSql As String
loginSql = "UPDATE Editors SET Editors.online = 1 WHERE Editors.ID = " & GBL_EditorID & ";"


DoCmd.SetWarnings False
DoCmd.RunSQL loginSql
DoCmd.SetWarnings True

End Function

Public Function logOut()
Dim logoutSql As String
logoutSql = "UPDATE Editors SET Editors.online = 0 WHERE Editors.ID = " & GBL_EditorID & ";"


DoCmd.SetWarnings False
DoCmd.RunSQL logoutSql
DoCmd.SetWarnings True

End Function

