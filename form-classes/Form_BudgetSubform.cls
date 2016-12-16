VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_BudgetSubform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterInsert()
Call Forms("BudgetForm").populateBoxes
Call Forms("BudgetForm").populateQBoxes
End Sub

Private Sub Form_AfterUpdate()
'Call Forms("BudgetForm").populateBoxes
End Sub

Private Sub Form_BeforeInsert(Cancel As Integer)
Me.header_id = Forms!BudgetForm.currentHeader
End Sub


Private Sub quarter_AfterUpdate()
Call Forms("BudgetForm").populateBoxes
Call Forms("BudgetForm").populateQBoxes
End Sub

Public Function goToNewRec()
DoCmd.GoToRecord , , acLast
End Function
