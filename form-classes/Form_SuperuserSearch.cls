VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SuperuserSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ADDBUNDLEBUTTON_Click()

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

