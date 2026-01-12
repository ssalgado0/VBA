'------------------------------------------------------------
' Macro: FichajeSalidaTeletrabajo
' Description:
'   Registers the daily clock-out time by updating a predefined
'   timesheet workbook based on the current date.
'
'   The macro records the clock-out time in the corresponding
'   worksheet and generates an email draft with the updated
'   timesheet attached for review and manual sending.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub FichajeSalidaTeletrabajo()
    Dim OutlookApp As Object
    Dim mailItem As Object
    Dim mailBody As String
    Dim folderPath As String
    Dim i As Integer
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim mesNombre As String
    Dim maxRows As Integer

    folderPath = "Q:\Route\to\CLOCKING IN FILE.xlsx"
   
    Select Case Month(Date)
        Case 1
            mesNombre = "ENERO"
        Case 2
            mesNombre = "FEBRERO"
        Case 3
            mesNombre = "MARZO"
        Case 4
            mesNombre = "ABRIL"
        Case 5
            mesNombre = "MAYO"
        Case 6
            mesNombre = "JUNIO"
        Case 7
            mesNombre = "JULIO"
        Case 8
            mesNombre = "AGOSTO"
        Case 9
            mesNombre = "SEPTIEMBRE"
        Case 10
            mesNombre = "OCTUBRE"
        Case 11
            mesNombre = "NOVIEMBRE"
        Case 12
            mesNombre = "DICIEMBRE"
        Case Else
            ' MsgBox "No se encontró el mes actual dentro del archivo."
    End Select
    
    Set wb = Workbooks.Open(folderPath)
    Set ws = Worksheets(mesNombre)
    
    ' Set initial row
    i = 1
    
    ' Set max rows number
    maxRows = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    ' Iterate over rows on column B to find current day
    While Not ws.Range("B" & i).Value = Day(Date)
        i = i + 1
        ' If day was not found
        If i > maxRows Then
            MsgBox "El día actual no se encontró en la columna B."
            Exit Sub
        End If
    Wend
    
    ' Set clock out time at column D
    ws.Range("D" & i).Value = Time

    ' Save file
    wb.Save
    wb.Close
    
    ' Generate clocking out mail
    Set OutlookApp = CreateObject("Outlook.Application")
    Set mailItem = OutlookApp.CreateItem(0)
    mailItem.To = "supervisor1@companyname.com; supervisor2@companyname.com"
    mailItem.Subject = "fichaje"
    mailBody = "<html><body>Buenos días,<br><br>" & _
               "Adjunto plantilla de fichaje de salida del día " & Date & _
               "<br><br>Un saludo."
    
    mailItem.Attachments.Add folderPath
    mailBody = mailBody & "</body></html>"
    mailItem.htmlBody = mailBody
    
    mailItem.Display
End Sub
