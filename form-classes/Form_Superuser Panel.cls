VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Superuser Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
 Dim namebox As Control
    Set namebox = [Forms]![Superuser Panel]![namebox]
    Dim usernamebox As Control
    Set usernamebox = [Forms]![Superuser Panel]![usernamebox]
    Dim useridbox As Control
    Set useridbox = [Forms]![Superuser Panel]![useridbox]
    namebox = GBL_Name
    usernamebox = Environ("USERNAME")
    useridbox = GBL_UserID
    Call showUsersOnline
If GBL_UserID = 88 Or GBL_UserID = 1897 Or GBL_UserID = 1826 Then
    'do nothing
   

Else
    DoCmd.Close
    MsgBox ("Access Denied")
    
    
End If
End Sub


Private Function showUsersOnline()
Dim rst As ADODB.Recordset
Set rst = New ADODB.Recordset
Dim usersOnlineQ As String
Dim numUsers As Integer
Dim userArray As Variant
Dim usersOnline As String

numUsers = 0
usersOnlineQ = "SELECT Editors.Editor FROM Editors WHERE Editors.online > 0;"

rst.Open usersOnlineQ, CurrentProject.Connection

userArray = rst.GetRows
usersOnline = " Currently Online:" & vbCrLf & vbCrLf
For i = 0 To UBound(userArray, 2)

usersOnline = usersOnline & "   " & userArray(0, i) & vbCrLf


Next i
Me.Text167.Value = usersOnline




End Function
