'------------------------------------------------------------
' Macro: ExtraerFechasDeUltimacion
' Description:
'   Queries the AEAT website to obtain the
'   final complete clearance date from
'   a list of MRNs and outputs the results to the
'   completion dates worksheet.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub ExtraerFechasDeUltimacion()
    Dim http As Object
    Dim htmlDoc As Object
    Dim parameterList As Variant
    Dim parameter As String
    Dim outputRange As Range
    Dim i As Long
    Dim url As String
    Dim baseUrl As String
    Dim lastRow As Long
    Dim ws As Worksheet
    Dim liElements As Object
    Dim liElement As Object
    Dim spanElement As Object
    Dim dateText As String
    Dim dateParts() As String
    Dim dayPart As String, monthPart As String, yearPart As String
    Dim results() As Variant
    
    ' Stop updates
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Fechas de ultimación")
    
    ' Get last row with data in column B
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    
    ' Save parameters pasted in column B by the user
    parameterList = ws.Range("B8:B" & lastRow).Value
    
    ' Prepare results matrix
    ReDim results(1 To UBound(parameterList, 1), 1 To 1)
    
    ' Base URL
    baseUrl = "https://www1.agenciatributaria.gob.es/wlpl/ADTR-JDIT/Ncts5Detalle?CLAVE="
    
    ' Parse AEAT webpage searching for the data of interest (completion date)
    For i = 1 To UBound(parameterList, 1)
        parameter = parameterList(i, 1)
        url = baseUrl & parameter
        
        ' Create HTTP object
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.send
        
        ' Load and save HTML
        Set htmlDoc = CreateObject("HTMLFILE")
        htmlDoc.body.innerHTML = http.responseText
        
        ' Search for the <li> element containing the completion date. Return blank if shipment has not been completed yet
        Set liElements = htmlDoc.getElementsByTagName("li")
        dateText = ""
        
        For Each liElement In liElements
            If InStr(liElement.innerHTML, "Fecha Final de Ultimación Completa:") > 0 Then
                Set spanElement = liElement.getElementsByTagName("span")(0)
                If Not spanElement Is Nothing Then
                    dateText = Replace(spanElement.innerText, "-", "/")
                    dateParts = Split(dateText, "/")
                    If UBound(dateParts) = 2 Then
                        dayPart = dateParts(0)
                        monthPart = dateParts(1)
                        yearPart = dateParts(2)
                        If IsNumeric(dayPart) And IsNumeric(monthPart) Then
                            ' Transform the date depending on the day of the month
                            If CInt(dayPart) <= 12 Then
                                dateText = monthPart & "/" & dayPart & "/" & yearPart
                            Else
                                dateText = dayPart & "/" & monthPart & "/" & yearPart
                            End If
                        End If
                    End If
                End If
                Exit For
            End If
        Next liElement
        
        ' Save result
        results(i, 1) = dateText
    Next i
    
    ' Save final results of all of the shipments
    Set outputRange = ws.Range("C8:C" & lastRow)
    outputRange.Value = results
    
    ' Execution finished
    ws.Cells(2, 2).Value = "¡Hecho!"
    
    ' Adjust content
    ws.Columns("A:C").AutoFit
    
    ' Reactivate Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
