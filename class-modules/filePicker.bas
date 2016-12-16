Attribute VB_Name = "filePicker"
Option Compare Database

Public Function Filename(ByVal strpath As String, sPath) As String
    sPath = Left(strpath, InStrRev(strpath, "\"))
    Filename = Mid(strpath, InStrRev(strpath, "\") + 1)
End Function
