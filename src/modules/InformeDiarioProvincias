'------------------------------------------------------------
' Macro: InformeDiarioProvincias
' Description:
'   Opens an Excel file selected by the user and
'   collects the records corresponding to the execution
'   date by iterating the data from the last row upwards.
'   Quantities are grouped and summed by location and
'   load type, generating a daily summary in a worksheet
'   named "Resumen" where the extracted data can be
'   reviewed.
'
'   Afterwards, a second Excel file is opened and the
'   calculated values are pasted into each location
'   worksheet, positioning them in the row corresponding
'   to the current date.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub InformeDiarioProvincias()
    Dim RutaArchivo As String
    Dim libro As Workbook
    Dim hoja As Worksheet
    Dim hojaResultado As Worksheet
    Dim ultimaFila As Long
    Dim filaActual As Long
    Dim celda As Range
    Dim localizacion As String
    Dim carga As String
    Dim cantidad As Long
    Dim resumen(1 To 8, 1 To 5) As Long
    Dim i As Integer
    Dim FechaHoy As Date
    Dim fechaCelda As Date
    Dim tabla As ListObject
    Dim rango As Range

    ' Load variables
    Dim h7inspSCQ As Long
    Dim inspSCQ As Long
    Dim clrd As Long
    Dim exe5 As Long
    Dim h7rlse As Long

    ' File variables
    Dim FileDialog As FileDialog
    Dim folderPath As String
    Dim fileName As String
    Dim fileSystem As Object


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' FIRST PART: Extract wanted data from the source Excel  
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Show select file dialog
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Disallow multiselect
    FileDialog.AllowMultiSelect = False
    FileDialog.Title = "Seleccione un archivo"

    ' File picker dialog show
    If FileDialog.Show = -1 Then
        RutaArchivo = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ningún archivo.", vbExclamation
        ' User cancels macro execution
        Exit Sub
    End If

    ' Today's date
    FechaHoy = Date
    
    ' Open report file downloaded from SQL Server Reporting Services
    Set libro = Workbooks.Open(RutaArchivo)
    
    ' Data is in the first sheet
    Set hoja = libro.Sheets(1)
    
    ' Set last file with data in column A
    ultimaFila = hoja.Cells(hoja.Rows.count, "A").End(xlUp).Row
    
    ' Initialize summary matrix
    For i = 1 To 8
        resumen(i, 1) = 0
        resumen(i, 2) = 0
        resumen(i, 3) = 0
        resumen(i, 4) = 0
    Next i
    
    ' Initialize counters
    h7inspSCQ = 0
    inspSCQ = 0
    clrd = 0
    exe5 = 0
    h7rlse = 0
    
    ' Iterate starting from last row upwards
    For filaActual = ultimaFila To 2 Step -1
        Set celda = hoja.Cells(filaActual, 1)
        
        ' Check if the cell has a date
        If IsDate(celda.Value) Then
            ' Get rid of hour data, only day for our purpose 
            fechaCelda = DateValue(celda.Value)
            
            ' Stop when the row's date is different from today
            If fechaCelda <> FechaHoy Then
                Exit For
            End If
            
            ' Get and convert necessary data from the adjacent cells
            localizacion = Trim(celda.Offset(0, 1).Value)
            carga = Trim(celda.Offset(0, 3).Value) 
            cantidad = CLng(celda.Offset(0, 2).Value)
                       
            ' Assign quantities by location and type of load
            Select Case localizacion
                Case "MAD"
                    i = 1
                    ' Count CLRD
                    If carga = "CLRD" Then
                        clrd = clrd + cantidad
                    ElseIf carga = "EXE5" Then
                        exe5 = exe5 + cantidad
                    ElseIf carga = "H7RLSE" Then
                        h7rlse = h7rlse + cantidad
                    End If
                Case "BCN"
                    i = 2
                Case "VIT"
                    i = 3
                Case "VLC"
                    i = 4
                Case "ALC"
                    i = 5
                Case "SVQ"
                    i = 6
                Case "SCQ"
                    i = 7
                    ' Count INSP and H7INSP for SCQ location
                    If carga = "INSP" Then
                        inspSCQ = inspSCQ + cantidad
                    ElseIf carga = "H7INSP" Then
                        h7inspSCQ = h7inspSCQ + cantidad
                    End If
                Case "XPA"
                    i = 8
                Case Else
                    i = 0
            End Select
            
            ' Sum up quantity according to the type of load
            If i > 0 Then
                If Right(carga, 1) = "0" Then
                    ' Do nothing
                ElseIf carga = "BRKR" Then
                    resumen(i, 1) = resumen(i, 1) + cantidad
                ElseIf carga Like "ADT*" Then
                    resumen(i, 2) = resumen(i, 2) + cantidad
                ElseIf carga = "H7INSP" Or carga = "H7RLSE" Or carga = "LOW3" Or carga = "SIMPL" Then
                    resumen(i, 3) = resumen(i, 3) + cantidad
                Else
                    resumen(i, 4) = resumen(i, 4) + cantidad
                End If
            End If
            
        Else
            ' Print in console in case of date error
            Debug.Print "Invalid date in cell: " & celda.Address
        End If
    Next filaActual
    
    ' Create new sheet "Resumen" for the results
    On Error Resume Next
        Set hojaResultado = libro.Sheets("Resumen")
    On Error GoTo 0
    If hojaResultado Is Nothing Then
        Set hojaResultado = libro.Sheets.Add
        hojaResultado.Name = "Resumen"
    End If
    
    ' Write headers in the new sheet
    hojaResultado.Range("A1").Value = "Localización"
    hojaResultado.Range("B1").Value = "Cesiones"
    hojaResultado.Range("C1").Value = "ADT's"
    hojaResultado.Range("D1").Value = "HV"
    hojaResultado.Range("E1").Value = "LV"
    
    hojaResultado.Range("A2").Value = "MAD"
    hojaResultado.Range("A3").Value = "BCN"
    hojaResultado.Range("A4").Value = "VIT"
    hojaResultado.Range("A5").Value = "VLC"
    hojaResultado.Range("A6").Value = "ALC"
    hojaResultado.Range("A7").Value = "SVQ"
    hojaResultado.Range("A8").Value = "SCQ"
    hojaResultado.Range("A9").Value = "XPA"
    
    hojaResultado.Range("A11").Value = "INSP"
    hojaResultado.Range("A12").Value = "H7INSP"
    
    hojaResultado.Range("A15").Value = "CLRD"
    hojaResultado.Range("A16").Value = "EXE5"
    hojaResultado.Range("A17").Value = "H7RLSE"
    
    ' Fill summary values in the result sheet with alternating colors
    For i = 1 To 8
        With hojaResultado
            .Cells(i + 1, 2).Value = resumen(i, 1)
            .Cells(i + 1, 3).Value = resumen(i, 2)
            .Cells(i + 1, 4).Value = resumen(i, 4) + (resumen(i, 1) + resumen(i, 2)) ' Sum up BRKR and ADT
            .Cells(i + 1, 5).Value = resumen(i, 3)
            
            ' Alternating colors for range B to F
            If i Mod 2 = 0 Then
                .Range(.Cells(i + 1, 2), .Cells(i + 1, 6)).Interior.Color = RGB(255, 255, 255)
            Else
                .Range(.Cells(i + 1, 2), .Cells(i + 1, 6)).Interior.Color = RGB(220, 230, 241)
            End If
        End With
    Next i
    
    ' Write INSP and H7INSP values for SCQ
    hojaResultado.Cells(11, 2).Value = inspSCQ
    hojaResultado.Cells(12, 2).Value = h7inspSCQ
    
    hojaResultado.Cells(15, 2).Value = clrd
    hojaResultado.Cells(16, 2).Value = exe5
    hojaResultado.Cells(17, 2).Value = h7rlse
    
   ' Alternating colors for row 11 until column F   
    If 11 Mod 2 = 0 Then
        hojaResultado.Range("A11:F11").Interior.Color = RGB(255, 255, 255) 
    Else
        hojaResultado.Range("A11:F11").Interior.Color = RGB(220, 230, 241) 
    End If
    
    ' Alternating colors for row 12 until column F
    If 12 Mod 2 = 0 Then
        hojaResultado.Range("A12:F12").Interior.Color = RGB(255, 255, 255) 
    Else
        hojaResultado.Range("A12:F12").Interior.Color = RGB(220, 230, 241) 
    End If
        
     ' Define table range
    Set rango = hojaResultado.Range("A14:B17")

    ' Create table
    Set tabla = hojaResultado.ListObjects.Add(xlSrcRange, rango, , xlYes)
    tabla.Name = "MiTabla"
    tabla.TableStyle = "TableStyleLight1" 
    
    rango.Cells(1, 1).Value = "RELEVO"
    rango.Cells(1, 2).Value = "MADRID"
    
    ' Release reference to source worksheet (no longer needed)
    Set hoja = Nothing
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' SECOND PART: Parse and paste the data in target Excel  
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ' Second part variables
    Dim fecha As Date
    Dim rutaCompleta As String
    Dim libro2 As Workbook
    Dim Hoja2 As Worksheet
    Dim celdaEncontrada As Range

    ' Pick the second file
    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)

    FileDialog.AllowMultiSelect = False
    FileDialog.Title = "Seleccione un archivo"

    If FileDialog.Show = -1 Then
        rutaCompleta = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ningún archivo.", vbExclamation
        Exit Sub
    End If

    ' Assign today's date and number of special customer packages
    fecha = Date
    numSpCustName = Application.InputBox("SPECIAL CUSTOMER NAME MAD")

    Set libro2 = Workbooks.Open(rutaCompleta)

    ' Paste the data
    Set Hoja2 = libro2.Sheets("MAD")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B2").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("C2").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("D2").Value
    celdaEncontrada.Offset(0, 5).Value = hojaResultado.Range("E2").Value
    celdaEncontrada.Offset(0, 9).Value = "SPECIAL CUSTOMER NAME: " & numSpCustName

    Set Hoja2 = libro2.Sheets("BCN")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B3").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("C3").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("D3").Value
    celdaEncontrada.Offset(0, 5).Value = hojaResultado.Range("E3").Value

    Set Hoja2 = libro2.Sheets("VIT")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B4").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("C4").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("D4").Value
    celdaEncontrada.Offset(0, 5).Value = hojaResultado.Range("E4").Value

    Set Hoja2 = libro2.Sheets("VLC")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)
    
    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B5").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("D5").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("E5").Value

    Set Hoja2 = libro2.Sheets("XXA")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B6").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("D6").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("E6").Value

    Set Hoja2 = libro2.Sheets("XVQ")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B7").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("D7").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("E7").Value

    Set Hoja2 = libro2.Sheets("SCQ")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B8").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("D8").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("E8").Value
    celdaEncontrada.Offset(0, 8).Value = hojaResultado.Range("B11").Value & " insp, " & hojaResultado.Range("B12").Value & " h7insp"

    Set Hoja2 = libro2.Sheets("XPA")
    Set celdaEncontrada = Hoja2.Columns("A").Find(What:=Date)

    celdaEncontrada.Offset(0, 2).Value = hojaResultado.Range("B9").Value
    celdaEncontrada.Offset(0, 3).Value = hojaResultado.Range("D9").Value
    celdaEncontrada.Offset(0, 4).Value = hojaResultado.Range("E9").Value

    ' Autosave disabled, save file manually
    ' libro2.Save
End Sub
