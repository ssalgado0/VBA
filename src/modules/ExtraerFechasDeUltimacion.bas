'------------------------------------------------------------
' Macro: ExtraerFechasDeUltimacion
' Descripción:
'   Consulta la web de la AEAT para obtener la
'   "Fecha Final de Ultimación Completa" a partir de
'   una lista de MRNs y vuelca los resultados en la hoja
'   "Fechas de ultimación".
'
' Autor: ssalgado0@uoc.edu
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
    
    ' Desactivar actualizaciones para acelerar
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Definir la hoja de cálculo
    Set ws = ThisWorkbook.Sheets("Fechas de ultimación")
    
    ' Última fila con datos en columna B
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    
    ' Leer parámetros en columna B
    parameterList = ws.Range("B8:B" & lastRow).Value
    
    ' Preparar matriz de resultados
    ReDim results(1 To UBound(parameterList, 1), 1 To 1)
    
    ' Base URL
    baseUrl = "https://www1.agenciatributaria.gob.es/wlpl/ADTR-JDIT/Ncts5Detalle?CLAVE="
    
    ' Bucle de parámetros
    For i = 1 To UBound(parameterList, 1)
        parameter = parameterList(i, 1)
        url = baseUrl & parameter
        
        ' Crear objeto HTTP
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.send
        
        ' Cargar HTML en documento
        Set htmlDoc = CreateObject("HTMLFILE")
        htmlDoc.body.innerHTML = http.responseText
        
        ' Buscar <li> con la fecha
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
        
        ' Guardar resultado en matriz
        results(i, 1) = dateText
    Next i
    
    ' Volcar resultados de una sola vez
    Set outputRange = ws.Range("C8:C" & lastRow)
    outputRange.Value = results
    
    ' Mensaje de confirmación
    ws.Cells(2, 2).Value = "¡Hecho!"
    
    ' Ajustar columnas
    ws.Columns("A:C").AutoFit
    
    ' Reactivar Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
