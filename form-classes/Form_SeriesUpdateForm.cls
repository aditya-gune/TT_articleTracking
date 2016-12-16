VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_SeriesUpdateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub MuteButton_Click()
Call MuteNotifications
DoCmd.Close acForm, "SeriesUpdateForm"
Call RestoreMain
End Sub

Private Sub Okbutton_Click()
DoCmd.Close acForm, "SeriesUpdateForm"
Call RestoreMain
End Sub

Private Sub Form_Load()
Dim updates As String
updates = seriesUpdate
UpdateBox.Value = updates
End Sub
