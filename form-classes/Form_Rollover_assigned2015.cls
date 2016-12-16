VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Rollover_assigned2015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Current()
Forms!NewYearForm.topicbox.Value = Me.Bundle_Topic
Forms!NewYearForm.authorbox.Value = Me.Author
Forms!NewYearForm.stypebox.Value = Me.Subtype
Forms!NewYearForm.idbox.Value = Me.ID
Forms!NewYearForm.fktopicbox.Value = Me.fk_topic

Call Forms("NewYearForm").refreshBoxes

End Sub
