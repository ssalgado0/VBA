'------------------------------------------------------------
' Macro: PrepararCorreosDesdePDFs
' Descripción:
'   Analiza una carpeta con PDFs de tránsitos, identifica el
'   destinatario y los números de albarán a partir del nombre
'   de los archivos, y genera correos de Outlook agrupados
'   por destinatario con los PDFs adjuntos.
'
' Requisitos:
'   - PDFs con nomenclatura "MRN + Destinatario + Nº Albarán"
'
' Autor: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub PrepararCorreosDesdePDFs()
    Dim OutlookApp As Object
    Dim mailItem As Object
    Dim FileDialog As Object
    Dim folderPath As String
    Dim fileName As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim recipient As String
    Dim recipientsList As Variant
    Dim recipientFound As Boolean
    Dim processedRecipient As String
    Dim i As Integer
    Dim recipientFiles As Object
    Dim recipientAlbaranes As Object
    Dim key As Variant
    Dim mailBody As String
    Dim albaranNumber As String
    Dim regex As Object
    Dim matches As Object
    Dim miImagen As String

    ' Inicializar la aplicación de Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Mostrar el diálogo para seleccionar una carpeta
    Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If FileDialog.Show = -1 Then
        folderPath = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ninguna carpeta.", vbExclamation
        Exit Sub
    End If
    
    ' Crear objeto FileSystem para trabajar con archivos y carpetas
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    
    ' Definir la lista de destinatarios
    recipientsList = Array _
    ("ADDRESSEE 1", _
    "ADDRESSEE 2", _
    "ADDRESSEE 3")

    ' Crear diccionarios para almacenar archivos y albaranes por destinatario
    Set recipientFiles = CreateObject("Scripting.Dictionary")
    Set recipientAlbaranes = CreateObject("Scripting.Dictionary")

    ' Crear un objeto Regex para encontrar números de 10 dígitos
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b\d{10}\b" ' Busca exactamente 10 dígitos
    regex.Global = True

    ' Recorrer cada archivo en la carpeta seleccionada
    For Each file In folder.Files
        If LCase(fileSystem.GetExtensionName(file.Name)) = "pdf" Then
            fileName = file.Name
            recipientFound = False
            processedRecipient = ""
            
            ' Ajustar nombre del archivo para casos especiales antes de la búsqueda
            If InStr(1, fileName, "ADDRESEE 1", vbTextCompare) > 0 Then
                processedRecipient = "ADDRESSEE 1"
                recipientFound = True
            ElseIf InStr(1, fileName, "ADRESSE 2", vbTextCompare) > 0 Then
                processedRecipient = "ADDRESSEE 2"
                recipientFound = True
            ElseIf InStr(1, fileName, "ADDRESIE 3", vbTextCompare) > 0 Then
                processedRecipient = "ADDRESSEE 3"
                recipientFound = True
            ElseIf InStr(1, fileName, "DRESSE 3", vbTextCompare) > 0 Then
                processedRecipient = "ADDRESSEE 3"
                recipientFound = True
            Else
                ' Buscar destinatario en el nombre del archivo
                For i = LBound(recipientsList) To UBound(recipientsList)
                    If InStr(1, fileName, recipientsList(i), vbTextCompare) > 0 Then
                        recipient = recipientsList(i)
                        recipientFound = True
                        processedRecipient = recipient
                        Exit For
                    End If
                Next i
            End If

            If recipientFound Then
                ' Añadir archivo al destinatario en el diccionario
                If Not recipientFiles.exists(processedRecipient) Then
                    recipientFiles.Add processedRecipient, New Collection
                    recipientAlbaranes.Add processedRecipient, New Collection
                End If
                recipientFiles(processedRecipient).Add file.Path

                ' Buscar número de albarán en el nombre del archivo
                Set matches = regex.Execute(fileName)
                If matches.count > 0 Then
                    albaranNumber = matches.item(0).Value
                    recipientAlbaranes(processedRecipient).Add albaranNumber
                End If
            End If
        End If
    Next file
    
    ' Crear y mostrar correos para cada destinatario
    For Each key In recipientFiles.Keys
        ' Crear un nuevo correo
        Set mailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
        
        If key = "ADDRESSEE 3" Then
            mailItem.Subject = "TRÁNSITOS " & key & " SOCIEDAD ANONIMA"
        Else
            mailItem.Subject = "TRÁNSITOS " & key
        End If
              
        ' mailItem.Subject = "TRÁNSITOS " & key
        mailBody = "<html><body>Buenos días,<br><br>" & _
                   "Adjuntamos tránsitos para ultimar a su llegada, AWB:<br><br>"
        
        ' Añadir números de albarán al cuerpo del correo
        If recipientAlbaranes.exists(key) Then
            For i = 1 To recipientAlbaranes(key).count
                mailBody = mailBody & recipientAlbaranes(key)(i) & "<br>"
            Next i
        End If
        
        ' Asignar la imagen correspondiente al destinatario
        Select Case key
            Case "ADDRESSEE 1"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario1.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            Case "ADDRESSEE 2"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario2.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            Case "ADDRESSEE 3"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario3.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            Case "ADDRESSEE 4"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario4.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            Case "ADDRESSEE 5"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario5.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            Case "ADDRESSEE 6"
                miImagen = "N:\ruta\a\la\imagen\con\la\direccion\del\destinatario6.jpg"
                mailItem.To = "mail1@domain.com; mail2@domain.com"
                mailItem.CC = "mail1@domain.com; mail2@domain.com; mail3@domain.com"
            ' Añade más casos según sea necesario
            Case Else
                mailItem.To = "default@example.com"
                mailItem.CC = "default_cc1@example.com; default_cc2@example.com"
        End Select
        
        ' Adjuntar la imagen con Content-ID
        Dim attachment As Object
        Set attachment = mailItem.Attachments.Add(miImagen)
        attachment.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001E", "miImagen"
        
        ' Crear el cuerpo del correo con la firma incluida
        mailBody = mailBody & "<br><br>La dirección de entrega es la siguiente:<br><br><img src='cid:miImagen'><br><br>"
        mailBody = mailBody & "<br><br>Un saludo,<br><br>"
        
        mailBody = mailBody & "</body></html>"
        
        mailItem.htmlBody = mailBody
        
        ' Añadir archivos PDF al correo
        For i = 1 To recipientFiles(key).count
            mailItem.Attachments.Add recipientFiles(key)(i)
        Next i
        
        ' Mostrar el correo (no enviar)
        mailItem.Display
    Next key
    
    ' Limpiar
    Set recipientFiles = Nothing
    Set recipientAlbaranes = Nothing
    Set fileSystem = Nothing
    Set folder = Nothing
    Set FileDialog = Nothing
    Set regex = Nothing
End Sub


