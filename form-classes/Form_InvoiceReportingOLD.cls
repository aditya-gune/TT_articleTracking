VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_InvoiceReportingOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'NOTE: FORM DEPRECATED

Private Sub EndDate_LostFocus()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting!totalbox
totalbox.Requery

Dim jmbox As Control
Set jmbox = Forms!InvoiceReporting!jmbox
jmbox.Requery

Dim anbox As Control
Set anbox = Forms!InvoiceReporting!anbox
anbox.Requery

Dim dbbox As Control
Set dbbox = Forms!InvoiceReporting!dbbox
dbbox.Requery

Dim agbox As Control
Set agbox = Forms!InvoiceReporting!agbox
agbox.Requery

Dim thbox As Control
Set thbox = Forms!InvoiceReporting!thbox
thbox.Requery

Dim tebox As Control
Set tebox = Forms!InvoiceReporting!tebox
tebox.Requery

'By media group
Dim storagebox As Control
Set storagebox = Forms!InvoiceReporting!storagebox
storagebox.Requery

Dim netbox As Control
Set netbox = Forms!InvoiceReporting!networkingbox
netbox.Requery

Dim dcvbox As Control
Set dcvbox = Forms!InvoiceReporting!dcvirtbox
dcvbox.Requery

Dim babox As Control
Set babox = Forms!InvoiceReporting!babox
babox.Requery

Dim secbox As Control
Set secbox = Forms!InvoiceReporting!secbox
secbox.Requery

Dim ciobox As Control
Set ciobox = Forms!InvoiceReporting!ciobox
ciobox.Requery

End Sub


Private Sub Refreshbutton_Click()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting!totalbox
totalbox.Requery

Dim jmbox As Control
Set jmbox = Forms!InvoiceReporting!jmbox
jmbox.Requery

Dim anbox As Control
Set anbox = Forms!InvoiceReporting!anbox
anbox.Requery

Dim dbbox As Control
Set dbbox = Forms!InvoiceReporting!dbbox
dbbox.Requery

Dim agbox As Control
Set agbox = Forms!InvoiceReporting!agbox
agbox.Requery

Dim thbox As Control
Set thbox = Forms!InvoiceReporting!thbox
thbox.Requery

Dim tebox As Control
Set tebox = Forms!InvoiceReporting!tebox
tebox.Requery

'By media group
Dim storagebox As Control
Set storagebox = Forms!InvoiceReporting!storagebox
storagebox.Requery

Dim netbox As Control
Set netbox = Forms!InvoiceReporting!networkingbox
netbox.Requery

Dim dcvbox As Control
Set dcvbox = Forms!InvoiceReporting!dcvirtbox
dcvbox.Requery

Dim babox As Control
Set babox = Forms!InvoiceReporting!babox
babox.Requery

Dim secbox As Control
Set secbox = Forms!InvoiceReporting!secbox
secbox.Requery

Dim ciobox As Control
Set ciobox = Forms!InvoiceReporting!ciobox
ciobox.Requery

End Sub

Private Sub StartDate_LostFocus()
Dim totalbox As Control
Set totalbox = Forms!InvoiceReporting!totalbox
totalbox.Requery

Dim jmbox As Control
Set jmbox = Forms!InvoiceReporting!jmbox
jmbox.Requery

Dim anbox As Control
Set anbox = Forms!InvoiceReporting!anbox
anbox.Requery

Dim dbbox As Control
Set dbbox = Forms!InvoiceReporting!dbbox
dbbox.Requery

Dim agbox As Control
Set agbox = Forms!InvoiceReporting!agbox
agbox.Requery

Dim thbox As Control
Set thbox = Forms!InvoiceReporting!thbox
thbox.Requery

Dim tebox As Control
Set tebox = Forms!InvoiceReporting!tebox
tebox.Requery

'By media group
Dim storagebox As Control
Set storagebox = Forms!InvoiceReporting!storagebox
storagebox.Requery

Dim netbox As Control
Set netbox = Forms!InvoiceReporting!networkingbox
netbox.Requery

Dim dcvbox As Control
Set dcvbox = Forms!InvoiceReporting!dcvirtbox
dcvbox.Requery

Dim babox As Control
Set babox = Forms!InvoiceReporting!babox
babox.Requery

Dim secbox As Control
Set secbox = Forms!InvoiceReporting!secbox
secbox.Requery

Dim ciobox As Control
Set ciobox = Forms!InvoiceReporting!ciobox
ciobox.Requery

End Sub
