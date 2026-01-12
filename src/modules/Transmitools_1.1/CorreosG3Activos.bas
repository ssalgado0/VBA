'------------------------------------------------------------
' Macro: CorreosG3Activos
' Description:
'   Prompts the user to select a CSV file and opens it in Excel.
'   Reads the header row and iterates through the data rows to
'   identify shipments with an active G3 status (excluding those
'   marked as "Revoked").
'
'   Records are grouped by customs office based on the office
'   code, building an HTML table per group. For each customs
'   office, an Outlook email draft is created with fixed "To"
'   recipients and a dynamic "CC" list, including the generated
'   table in the email body for review.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub CorreosG3Activos()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long, ultimaFila As Long
    Dim OutlookApp As Object, OutlookMail As Object
    Dim Archivo As String
    Dim lineaDatos As Variant
    Dim encabezados As Variant
    Dim tablaHTML As String
    Dim encabezadosHtml As String
    Dim correosDict As Object 
    Dim subjectDict As Object 
    Dim filasValidas As Long 

    ' Create dictionaries
    Set correosDict = CreateObject("Scripting.Dictionary")
    Set subjectDict = CreateObject("Scripting.Dictionary")

    ' File selection
    Dim FileDialog As Object
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    If FileDialog.Show = -1 Then
        Archivo = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ningún archivo.", vbExclamation
        Exit Sub
    End If

    ' Open the CSV file
    Set wb = Workbooks.Open(Archivo)
    Set ws = wb.Sheets(1)

    Dim valorCelda As String
    valorCelda = ws.Cells(1, 1).Value
    If Right(valorCelda, 1) = ";" Then valorCelda = Left(valorCelda, Len(valorCelda) - 1)
    encabezados = Split(valorCelda, ";")

    ' HTML header row
    encabezadosHtml = "<tr style='background-color: #c2c9cc; font-weight: bold;'>"
    For Each dato In encabezados
        encabezadosHtml = encabezadosHtml & "<td style='padding: 5px;'>" & dato & "</td>"
    Next dato
    encabezadosHtml = encabezadosHtml & "</tr>"

    ' Determine last row with data
    ultimaFila = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    filasValidas = 0

    ' Iterate through rows
    For i = 2 To ultimaFila
        valorCelda = ws.Cells(i, 1).Value
        If Replace(valorCelda, ";", "") = "" Then GoTo SiguienteFila
        If Right(valorCelda, 1) = ";" Then valorCelda = Left(valorCelda, Len(valorCelda) - 1)
        lineaDatos = Split(valorCelda, ";")

        If UBound(lineaDatos) >= 7 Then
            If lineaDatos(7) = "Revocado" Then GoTo SiguienteFila

            Dim copiaCorreo As String
            Dim aduanaCodigo As String
            
            Select Case lineaDatos(2)
                Case "ES002801": copiaCorreo = "2801people@companyname.com; 2801anotherpeople@companyname.com": aduanaCodigo = "Aduana 2801"
                Case "ES000801": copiaCorreo = "0801people@companyname.com; 0801anotherpeople@companyname.com": aduanaCodigo = "Aduana 0801"
                Case "ES000101": copiaCorreo = "0101people@companyname.com; 0101anotherpeople@companyname.com": aduanaCodigo = "Aduana 0101"
                Case "ES004601": copiaCorreo = "4601people@companyname.com; 4601anotherpeople@companyname.com": aduanaCodigo = "Aduana 4601"
                Case "ES000301": copiaCorreo = "0301people@companyname.com; 0301anotherpeople@companyname.com": aduanaCodigo = "Aduana 0301"
                Case "ES001507": copiaCorreo = "1507people@companyname.com; 1507anotherpeople@companyname.com": aduanaCodigo = "Aduana 1507"
                Case "ES004101": copiaCorreo = "4101people@companyname.com; 4101anotherpeople@companyname.com": aduanaCodigo = "Aduana 4101"
              ' Add more cases as needed in the future
                Case Else: copiaCorreo = "": aduanaCodigo = ""
            End Select

            If copiaCorreo <> "" Then
                If Not correosDict.exists(copiaCorreo) Then
                    correosDict.Add copiaCorreo, encabezadosHtml
                    subjectDict.Add copiaCorreo, aduanaCodigo
                End If
                correosDict(copiaCorreo) = correosDict(copiaCorreo) & "<tr>"
                For Each dato In lineaDatos
                    correosDict(copiaCorreo) = correosDict(copiaCorreo) & "<td style='padding: 5px;'>" & dato & "</td>"
                Next dato
                correosDict(copiaCorreo) = correosDict(copiaCorreo) & "</tr>"
                filasValidas = filasValidas + 1
            End If
        End If
SiguienteFila:
    Next i

    ' If no valid rows were found, notify and exit
    If filasValidas = 0 Then
        MsgBox "No se encontró ningún G3 activo en el archivo CSV", vbInformation
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Initialize Outlook
    On Error Resume Next
    Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If OutlookApp Is Nothing Then
        MsgBox "Outlook no está disponible.", vbExclamation
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    ' Create emails
    Dim key As Variant
    For Each key In correosDict.Keys
        tablaHTML = "<table border='1' style='border-collapse: collapse; width:100%; text-align:left; background-color: #d3e6ed;'>"
        tablaHTML = tablaHTML & correosDict(key) & "</table>"

        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = "somepeople@companyname.com; responsible1@companyname.com"
            .CC = key
            .Subject = "Revisar G3 Activos " & subjectDict(key)
            .htmlBody = "Buenos días,<br><br>" & _
                        "Revisar estos envíos con G3 activo sin PREH7:<br><br>" & _
                        tablaHTML & "<br><br>" & _
                        "Muchas gracias,<br><br>Un saludo."
            ' Show generated mails for the user to check
            .Display
        End With
        Set OutlookMail = Nothing
    Next key

    ' Do not save and exit
    wb.Close SaveChanges:=False
    Set OutlookApp = Nothing
End Sub
