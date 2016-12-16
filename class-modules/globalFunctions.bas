Attribute VB_Name = "globalFunctions"
Option Compare Database

Public Function RestoreMain()
DoCmd.SelectObject acForm, "Main Menu"
DoCmd.Restore
End Function


Public Function exportall()
Debug.Print (CurrentProject.Path)
For Each c In Application.VBE.VBProjects(1).VBComponents
Select Case c.Type
    Case vbext_ct_ClassModule, vbext_ct_Document
        Sfx = ".cls"
    Case vbext_ct_MSForm
        Sfx = ".frm"
    Case vbext_ct_StdModule
        Sfx = ".bas"
    Case Else
        Sfx = ""
End Select
If Sfx <> "" Then
    c.Export _
        Filename:=CurrentProject.Path & "\trackingcode\" & _
        c.name & Sfx
End If
Next c
End Function
