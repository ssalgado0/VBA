'------------------------------------------------------------
' Macro: DescargarTransitosPDF
' Description:
'   Iterates through a list of transit reference keys stored
'   in the worksheet and queries the AEAT NCTS5 detail service
'   for each record.
'
'   The macro parses the returned HTML to extract recipient
'   identification data, recipient name, destination customs
'   office, and reference number.
'
'   Based on predefined rule-based validations, the macro
'   determines the standardized recipient names and constructs
'   the corresponding AEAT document URL.
'
'   Finally, the associated transit PDF document is downloaded
'   and stored locally using a structured naming convention.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub DescargarTransitosPDF()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim url As String, html As String
    Dim baseUrl As String
    Dim http As Object
    Dim regex As Object, matches As Object
    Dim claveEE As String, claveCAT As String
    Dim nuevoEnlace As String
    Dim referencia As String
    Dim mrnText As String
    Dim baseUrl2 As String, url2 As String
    Dim parameter As String
    Dim liElements As Object
    Dim spanElement As Object
    Dim identificadorDestText As String, aduanaText As String
    Dim precintoText As String, referenciaText As String, nombreText As String
    Dim foundDestinatario As Boolean
    Dim nombreFinal As String
    Dim j As Long
    Dim htmlDoc As Object
    
    ' Disable updates
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set ws = ThisWorkbook.Sheets("Listado nombres T1")
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).Row
    
    baseUrl = "https://www1.agenciatributaria.gob.es/wlpl/inwinvoc/es.aeat.dit.adu.eeca.ee.viev.VisualizaEVSv?fProc=DB04&fReferencia="
    baseUrl2 = "https://www1.agenciatributaria.gob.es/wlpl/ADTR-JDIT/Ncts5Detalle?CLAVE="

    ' Start iteration over rows
    For i = 8 To lastRow

        ' Skip if previous declaration was erratic
        If ws.Cells(i, 4).Value = "FALTA O DISCREPANCIA EN UNO O VARIOS CAMPOS: NOMBRE DESTINATARIO || ADUANA DESTINO || CIF DESTINATARIO" Then
            GoTo NextI
        End If
        
        '================================
        ' 1) NCTS5 DETAIL REQUEST
        '================================

        parameter = Trim$(CStr(ws.Cells(i, 2).Value))
        If Len(parameter) = 0 Then GoTo NextI

        ' Assemble the first URL with the parameter
        url2 = baseUrl2 & parameter
        
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url2, False
        http.send
        
        If http.Status <> 200 Then GoTo NextI

        ' Save HTML
        Set htmlDoc = CreateObject("HTMLFILE")
        htmlDoc.body.innerHTML = http.responseText

        ' Initialize variables
        mrnText = ""
        identificadorDestText = ""
        aduanaText = ""
        precintoText = ""
        nombreText = ""
        referenciaText = ""
        foundDestinatario = False
        
        ' Search for data contained in <li> element
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
      
        aduanaText = Replace(aduanaText, Left(aduanaText, 4), "", 1, 1)
        identificadorDestText = Replace(identificadorDestText, Left(identificadorDestText, 2), "", 1, 1)
        
        '================================
        ' IDENTIFICATION RULES
        '================================
        
        ' Add more rules if needed
        If (InStr(nombreText, "NAME ONE SA") > 0 Or InStr(nombreText, "RECIPIENT NAME ONÉ") > 0) _
           And (aduanaText = "1111" Or aduanaText = "1212") _
           And identificadorDestText = "A00112233" Then
            nombreFinal = "RECIPIENT NAME ONE"

        ElseIf (InStr(nombreText, "RECIPI NAME TWO") > 0 Or InStr(nombreText, "ADDRESSEE NAME TWO") > 0 Or InStr(nombreText, "NAME TWO COMPANY") > 0) _
           And aduanaText = "2233" _
           And identificadorDestText = "B22446688" Then
            nombreFinal = "RECIPIENT NAME TWO"

        ElseIf (InStr(nombreText, "REC NAME THREE") > 0 Or InStr(nombreText, "COMPANY THREE ALT NAME") > 0 Or InStr(nombreText, "NAME THREE SL") > 0) _
           And (aduanaText = "3345" Or aduanaText = "4433") _
           And identificadorDestText = "F33669900" Then
            nombreFinal = "RECIPIENT NAME THREE"
            
        Else
            nombreFinal = "FALTA O DISCREPANCIA EN UNO O VARIOS CAMPOS: NOMBRE DESTINATARIO || ADUANA DESTINO || CIF DESTINATARIO"
        End If

        
        '=========================================
        ' 2) GET CLAVE_EE and CLAVE_CAT KEYS
        '=========================================
        referencia = parameter
        url = baseUrl & referencia

        ' Get HTML
        Set http = CreateObject("MSXML2.XMLHTTP")
        http.Open "GET", url, False
        http.send
        If http.Status <> 200 Then GoTo NextI
        
        html = http.responseText

        ' Search for "CLAVE_EE" and "CLAVE_CAT" identificators in the HTML
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = "CLAVE_EE=([^&]+).*CLAVE_CAT=([^&]+)"
        regex.Global = True
        
        Set matches = regex.Execute(html)
        
        If matches.count = 0 Then GoTo NextI
        
        claveEE = matches(0).SubMatches(0)
        claveCAT = matches(0).SubMatches(1)

        ' Second HTML webpage (associated to transit PDF file) assembly
        nuevoEnlace = "https://www1.agenciatributaria.gob.es/wlpl/inwinvoc/es.aeat.dit.adu.eeca.catalogo.vis.Visualiza?" & _
                      "COMPLETA=SI&ORIGEN=C&CLAVE_CAT=" & claveCAT & "&CLAVE_EE=" & claveEE
        
        '==============================================================
        ' 3) DESCARGA DEL PDF 
        ' Pass by parameters the data found in the previous steps
        '==============================================================

        ' PDF "sub macro" download
        Call DescargarPDFDesdeUrl( _
                nuevoEnlace, _
                parameter, _
                referenciaText, _
                nombreFinal, _
                nombreText)
        
NextI:
    Next i
    
    ' Reactivate Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    ' Process done message
    With ws
        .Cells(2, 2).Value = "¡Tránsitos descargados!"
        .Range("J:XFD").EntireColumn.Hidden = True
    End With
    
End Sub

'====================================
' SEPARATION WITH "PARENT" MACRO
'====================================

Sub DescargarPDFDesdeUrl( _
        pdfUrl As String, _
        parameter As String, _
        referenciaText As String, _
        nombreFinal As String, _
        nombreText As String)

    Dim http As Object
    Dim pdfBytes() As Byte
    Dim savePath As String
    Dim fNum As Integer
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", pdfUrl, False
    http.send
    
    If http.Status = 200 Then
        
        pdfBytes = http.responseBody
        
        ' PDF final name
        If nombreFinal = "FALTA O DISCREPANCIA EN UNO O VARIOS CAMPOS: NOMBRE DESTINATARIO || ADUANA DESTINO || CIF DESTINATARIO" Then
            savePath = "N:\mad-hub\data\CustomsIMP\Shared\TRANSITOS DESPACHO RAPIDO WPX - ESI\WPX\TransitosDescargados\" & parameter & " " & nombreText & " " & referenciaText & ".pdf"
        Else
            savePath = "N:\mad-hub\data\CustomsIMP\Shared\TRANSITOS DESPACHO RAPIDO WPX - ESI\WPX\TransitosDescargados\" & parameter & " " & nombreFinal & " " & referenciaText & ".pdf"
        End If
        
        ' Save PDF
        fNum = FreeFile
        Open savePath For Binary As #fNum
        Put #fNum, , pdfBytes
        Close #fNum
        
    Else
        Debug.Print "ERROR DESCARGANDO PDF (" & http.Status & "): " & pdfUrl
    End If
End Sub
