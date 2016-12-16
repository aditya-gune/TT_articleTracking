VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'NOTE: FORM DEPRECATED

Private Sub ADDBUNDLEBUTTON_Click()

End Sub

Private Sub Command154_Click()

End Sub

Private Sub Check169_Click()
Me!InvoiceMonth1.Enabled = True
Dim i As Integer
Select Case Me.Check169
    Case "-1"
        Me.InvoiceMonth1.Enabled = False
        Me.textSelection1.Value = 1
        Me.textSelection2.Value = 2
        Me.textSelection3.Value = 3
        Me.textSelection4.Value = 4
        Me.textSelection5.Value = 5
        Me.textSelection6.Value = 6
        Me.textSelection7.Value = 7
        Me.textSelection8.Value = 8
        Me.textSelection9.Value = 9
        Me.textSelection10.Value = 10
        Me.textSelection11.Value = 11
        Me.textSelection12.Value = 12
    Case "0"
        Me.InvoiceMonth1.Enabled = True
        Me.textSelection1.Value = Null
        Me.textSelection2.Value = Null
        Me.textSelection3.Value = Null
        Me.textSelection4.Value = Null
        Me.textSelection5.Value = Null
        Me.textSelection6.Value = Null
        Me.textSelection7.Value = Null
        Me.textSelection8.Value = Null
        Me.textSelection9.Value = Null
        Me.textSelection10.Value = Null
        Me.textSelection11.Value = Null
        Me.textSelection12.Value = Null
        For i = 0 To Me.InvoiceMonth1.ListCount
            Me!InvoiceMonth1.Selected(i) = False
            Next i
End Select
End Sub

Private Sub CompleteCheckBox_AfterUpdate()
Me.Status.Enabled = True
Me.WIPCheckBox.Enabled = True
Select Case Me.CompleteCheckBox
    Case "-1"
        Me.Status.Enabled = False
        Me.WIPCheckBox.Enabled = False
    Case "0"
       Me.Status.Enabled = True
       Me.WIPCheckBox.Enabled = True
End Select
End Sub

Private Sub InvoiceMonth1_Click()
'  MsgBox (CStr(Me.InvoiceMonth1.ItemsSelected.Count))
 Dim sTemp As String
 If Me!InvoiceMonth1.ItemsSelected.Count <> 0 Then
        For Each oItem In Me!InvoiceMonth1.ItemsSelected
            If iCount = 0 Then
                sTemp = sTemp & Me!InvoiceMonth1.ItemData(oItem)
                iCount = iCount + 1
            Else
                sTemp = sTemp & "," & Me!InvoiceMonth1.ItemData(oItem)
                iCount = iCount + 1
            End If
        Next oItem
    Else
        MsgBox "Nothing was selected from the list", vbInformation
        Exit Sub  'Nothing was selected
    End If
    'Stop
    'Me!textSelection1.Value = sTemp
    Dim strArray() As String
    strArray() = Split(sTemp, ",")
    For i = LBound(strArray) To UBound(strArray)
        If i = 0 Then
            Me!textSelection1.Value = strArray(0)
        ElseIf i = 1 Then
            Me!textSelection2.Value = strArray(1)
        ElseIf i = 2 Then
            Me!textSelection3.Value = strArray(2)
        ElseIf i = 3 Then
            Me!textSelection4.Value = strArray(3)
        ElseIf i = 4 Then
            Me!textSelection5.Value = strArray(4)
        ElseIf i = 5 Then
            Me!textSelection6.Value = strArray(5)
        ElseIf i = 6 Then
            Me!textSelection7.Value = strArray(6)
        ElseIf i = 7 Then
            Me!textSelection8.Value = strArray(7)
        ElseIf i = 8 Then
            Me!textSelection9.Value = strArray(8)
        End If
    Next i
End Sub

Private Sub WIPCheckBox_AfterUpdate()
Me.Status.Enabled = True
Me.CompleteCheckBox.Enabled = True
Select Case Me.WIPCheckBox
        Case "-1"
            Me.Status.Enabled = False
            Me.CompleteCheckBox.Enabled = False
        Case "0"
            Me.Status.Enabled = True
            Me.CompleteCheckBox.Enabled = True
End Select
End Sub

