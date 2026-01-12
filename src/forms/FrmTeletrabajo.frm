VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTeletrabajo 
   Caption         =   "Fichaje"
   ClientHeight    =   2052
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3432
   OleObjectBlob   =   "FrmTeletrabajo.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "FrmTeletrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonOK_Click()
    If OptionButtonYes.Value = True Then
        esTeletrabajo = True
    ElseIf OptionButtonNo.Value = True Then
        esTeletrabajo = False
    End If
    Me.Hide ' Oculta el UserForm
End Sub

Private Sub UserForm_Click()

End Sub
