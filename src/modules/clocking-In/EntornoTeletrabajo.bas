'------------------------------------------------------------
' Module: EntornoTeletrabajo
' Description:
'   Provides shared state and helper procedures related to
'   teleworking workflows.
'
'   This module exposes a public flag indicating teleworking
'   status and a helper routine to display the corresponding
'   UserForm.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Public esTeletrabajo As Boolean

Sub MostrarUserFormTeletrabajo()
    FrmTeletrabajo.Show
End Sub
