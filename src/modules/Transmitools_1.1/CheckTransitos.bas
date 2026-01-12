'------------------------------------------------------------
' Macro: CheckTransitos
' Description:
'   Reads a list of transit keys from the worksheet and parses
'   the AEAT transit detail page for each one.
'
'   The macro extracts recipient information, destination
'   customs office, reference number, and seal data from the
'   returned HTML, then validates the results using rule-based
'   checks.
'
'   Valid records are written to the worksheet, while
'   inconsistent or incomplete records are flagged with an
'   error message. Results are finally sorted for easier
'   review.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub CheckTransitos()
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
    Dim spanElement As Object
    Dim aduanaText As String
    Dim precintoText As String, referenciaText As String, nombreText As String, identificadorDestText As String
    Dim j As Long
    Dim foundDestinatario As Boolean
    Dim nombreFinal As String

    ' Disable updates to speed up execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Define worksheet
    Set ws = ThisWorkbook.Sheets("Listado nombres T1")
    
    ' Get last row with data in column B
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row

    ' Read parameters
    parameterList = ws.Range("B8:B" & lastRow).Value
    Set outputRange = ws.Range("C8:C" & lastRow)

    ' Base URL
    baseUrl = "https://www1.agenciatributaria.gob.es/wlpl/ADTR-JDIT/Ncts5Detalle?CLAVE="

    ' Main loop
    For i = 1 To UBound(parameterList, 1)
        parameter = parameterList(i, 1)
        url = baseUrl & parameter

        ' Download HTML
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.send

        ' Load into HTMLDocument
        Set htmlDoc = CreateObject("HTMLFILE")
        htmlDoc.body.innerHTML = http.responseText

        ' Initialize variables
        identificadorDestText = ""
        aduanaText = ""
        precintoText = ""
        referenciaText = ""
        nombreText = ""
        foundDestinatario = False

        ' Search for <li> elements that we want
        Set liElements = htmlDoc.getElementsByTagName("li")
        For j = 0 To liElements.Length - 1
            If InStr(liElements(j).innerText, "DESTINATARIO (de Cabecera).") > 0 Then
                foundDestinatario = True
            ElseIf foundDestinatario And identificadorDestText = "" And InStr(liElements(j).innerText, "Identificador:") > 0 Then
                On Error Resume Next
                Set spanElement = liElements(j).getElementsByTagName("span")(0)
                If Not spanElement Is Nothing Then identificadorDestText = spanElement.innerText
                On Error GoTo 0
            ElseIf foundDestinatario And nombreText = "" And InStr(liElements(j).innerText, "Nombre:") > 0 Then
                On Error Resume Next
                Set spanElement = liElements(j).getElementsByTagName("span")(0)
                If Not spanElement Is Nothing Then nombreText = spanElement.innerText
                On Error GoTo 0
            End If
        Next j

        ' Other fields
        For j = 0 To liElements.Length - 1
            If InStr(liElements(j).innerText, "Número de Referencia UCR:") > 0 Then
                On Error Resume Next
                Set spanElement = liElements(j).getElementsByTagName("span")(0)
                If Not spanElement Is Nothing Then referenciaText = spanElement.innerText
                On Error GoTo 0
            ElseIf InStr(liElements(j).innerText, "Aduana de Destino Declarada:") > 0 Then
                On Error Resume Next
                Set spanElement = liElements(j).getElementsByTagName("span")(0)
                If Not spanElement Is Nothing Then aduanaText = spanElement.innerText
                On Error GoTo 0
            End If
        Next j

        ' Get seal number; if it starts with "X/", remove the prefix
        On Error Resume Next
        precintoText = htmlDoc.getElementsByTagName("tbody")(3).getElementsByTagName("tr")(0).getElementsByTagName("td")(3).innerText
        On Error GoTo 0
        If InStr(precintoText, "/") = 2 Then
            precintoText = Mid(precintoText, 3)
        End If

        aduanaText = Replace(aduanaText, Left(aduanaText, 4), "", 1, 1)
        identificadorDestText = Replace(identificadorDestText, Left(identificadorDestText, 2), "", 1, 1)

        ' Evaluations (evaluation placeholder cases look -and actually are- duplicated for privacy; actual production logic uses unrepeated values)
        If (InStr(nombreText, "RECIPIEN NAME ONE") > 0 Or InStr(nombreText, "RECIP NAME ONE") > 0) _
           And (aduanaText = "1111" Or aduanaText = "1212") _
           And identificadorDestText = "A00112233" Then
            nombreFinal = "RECIPIENT NAME ONE"

        ElseIf (InStr(nombreText, "RECIPIEN NAME TWO") > 0 Or InStr(nombreText, "RECIP NAME TWO") > 0 Or InStr(nombreText, "REC NAME TWO") > 0) _
           And (aduanaText = "2222" Or aduanaText = "2323") _
           And identificadorDestText = "B11223344" Then
            nombreFinal = "RECIPIENT NAME TWO"

        ElseIf (InStr(nombreText, "RECIPI NAME THREE") > 0 Or InStr(nombreText, "THREE") > 0 Or InStr(nombreText, "NAME THREE") > 0) _
           And aduanaText = "3333" _
           And identificadorDestText = "A22334455" Then
            nombreFinal = "RECIPIENT NAME THREE"

        Else
            ' In case of error detection
            nombreFinal = "FALTA O DISCREPANCIA EN UNO O VARIOS CAMPOS: NOMBRE DESTINATARIO || ADUANA DESTINO || CIF DESTINATARIO"
        End If

        ' Save data on worksheet
        ' Error in declaration case: save just error message
        If nombreFinal = "FALTA O DISCREPANCIA EN UNO O VARIOS CAMPOS: NOMBRE DESTINATARIO || ADUANA DESTINO || CIF DESTINATARIO" Then
            outputRange.Cells(i, 2).Value = nombreFinal
        Else
        ' Succesful declaration: save transit data
            outputRange.Cells(i, 2).Value = referenciaText
            outputRange.Cells(i, 3).Value = parameter
            outputRange.Cells(i, 4).Value = precintoText
            outputRange.Cells(i, 5).Value = nombreFinal
            outputRange.Cells(i, 6).Value = identificadorDestText
        End If
    Next i

    ' Sort results
    Dim rng As Range
        
    ' Define the range from B8 to the last row with data
    Set rng = ws.Range("B8:H" & lastRow)
    
    ' Sort in ascending order
    rng.Sort Key1:=ws.Range("D8"), Order1:=xlAscending, Header:=xlNo
    
    ' Reactivate Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' Macro end msg
    With ws
        .Cells(2, 2).Value = "¡Hecho!"
        .Range("J:XFD").EntireColumn.Hidden = True
    End With
End Sub
