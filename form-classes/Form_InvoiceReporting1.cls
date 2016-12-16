VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InvoiceReporting1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub EndDate_LostFocus()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting1!totalbox
totalbox.Requery
Forms![InvoiceReporting1]![TotalSpentEditorsubform].Requery
Forms![InvoiceReporting1]![TotalSpentMGsubform].Requery
End Sub




Private Sub FormHeader_Click()
Call RestoreMain
End Sub

Private Sub Refreshbutton_Click()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting1!totalbox
totalbox.Requery
Forms![InvoiceReporting1]![TotalSpentEditorsubform].Requery
Forms![InvoiceReporting1]![TotalSpentMGsubform].Requery
End Sub

Private Sub StartDate_LostFocus()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting1!totalbox
totalbox.Requery
Forms![InvoiceReporting1]![TotalSpentEditorsubform].Requery
Forms![InvoiceReporting1]![TotalSpentMGsubform].Requery
End Sub
