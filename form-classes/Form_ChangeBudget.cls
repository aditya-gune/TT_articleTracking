VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_ChangeBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Form_Load()
If CurrentProject.AllForms("BudgetForm").IsLoaded = False Then
    MsgBox ("Error: Cannot change budget when Budget Management Center is closed")
    DoCmd.Close acForm, "ChangeBudget"
    
End If


End Sub

Private Sub newAmount_AfterUpdate()
Dim amount As String
Dim num As Long
Dim objregEx As Object
Set objregEx = CreateObject("VBScript.RegExp")
objregEx.IgnoreCase = True
objregEx.Global = True

objregEx.pattern = "[^0-9.]"

amount = objregEx.Replace(Me.newAmount.Value, "")

If InStr(amount, ".") Then
num = Round(Val(amount))
Debug.Print (num)
Me.newAmount = num
Else
Me.newAmount = amount
End If
End Sub

Private Sub SubmitBtn_Click()
Dim newBudgetAmt As String
If IsNull(Me.newAmount) = False Then

 newBudgetAmt = "UPDATE Months SET Months.amount =" & Me.newAmount & _
                " WHERE Months.ID = " & Forms![BudgetForm].currentHeader & ";"
 DoCmd.SetWarnings False
 DoCmd.RunSQL newBudgetAmt
 DoCmd.SetWarnings True
  
  
 DoCmd.Close acForm, "BudgetForm"
 DoCmd.Close acForm, "ChangeBudget"
 DoCmd.OpenForm "BudgetForm", acNormal
 Else
 MsgBox ("Error: annual budget must be a number.")
 End If


End Sub
