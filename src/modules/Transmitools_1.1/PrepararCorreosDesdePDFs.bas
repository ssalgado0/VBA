'------------------------------------------------------------
' Macro: PrepareEmailsFromPDFs
' Description:
'   Analyzes a folder containing transit PDFs, identifies the
'   recipient and delivery note numbers from the file names,
'   and generates Outlook emails grouped by recipient with
'   the PDFs attached.
'
' Requirements:
'   - PDFs named using the format "MRN + Recipient + AWB No."
'
' Author: ssalgado0@uoc.edu
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

    ' Initialize Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Show the dialog to select a folder
    Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If FileDialog.Show = -1 Then
        folderPath = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ninguna carpeta.", vbExclamation
        Exit Sub
    End If
    
    ' Create FileSystem object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    
    ' Define list of recipients
    recipientsList = Array _
    ("ADDRESSEE 1", _
    "ADDRESSEE 2", _
    "ADDRESSEE 3")

    ' Create dictionaries to store files and AWBs by recipient
    Set recipientFiles = CreateObject("Scripting.Dictionary")
    Set recipientAlbaranes = CreateObject("Scripting.Dictionary")

    ' Create a Regex object to find 10-digit numbers
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b\d{10}\b"
    regex.Global = True

    ' Loop through each file in the selected folder
    For Each file In folder.Files
        If LCase(fileSystem.GetExtensionName(file.Name)) = "pdf" Then
            fileName = file.Name
            recipientFound = False
            processedRecipient = ""
            
            ' Try to save some common errors when users name PDF files
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
                ' Search for recipient in the file name
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
                ' Add file to the recipient in the dictionary
                If Not recipientFiles.exists(processedRecipient) Then
                    recipientFiles.Add processedRecipient, New Collection
                    recipientAlbaranes.Add processedRecipient, New Collection
                End If
                recipientFiles(processedRecipient).Add file.Path

                ' Search for AWB number in the file name
                Set matches = regex.Execute(fileName)
                If matches.count > 0 Then
                    albaranNumber = matches.item(0).Value
                    recipientAlbaranes(processedRecipient).Add albaranNumber
                End If
            End If
        End If
    Next file
    
    ' Create and display emails for each recipient
    For Each key In recipientFiles.Keys
        ' Create a new email
        Set mailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
        
        If key = "ADDRESSEE 3" Then
            mailItem.Subject = "TRÁNSITOS " & key & " SOCIEDAD ANONIMA"
        Else
            mailItem.Subject = "TRÁNSITOS " & key
        End If
              
        ' mailItem.Subject = "TRÁNSITOS " & key
        mailBody = "<html><body>Buenos días,<br><br>" & _
                   "Adjuntamos tránsitos para ultimar a su llegada, AWB:<br><br>"
        
        ' Add AWB numbers to the email body
        If recipientAlbaranes.exists(key) Then
            For i = 1 To recipientAlbaranes(key).count
                mailBody = mailBody & recipientAlbaranes(key)(i) & "<br>"
            Next i
        End If
        
        ' Assign the corresponding delivery address image to the recipient
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
            ' Add more cases as needed, structure is in the next lines
            Case Else
                mailItem.To = "default@example.com"
                mailItem.CC = "default_cc1@example.com; default_cc2@example.com"
        End Select
        
        ' Attach the image with Content-ID
        Dim attachment As Object
        Set attachment = mailItem.Attachments.Add(miImagen)
        attachment.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001E", "miImagen"
        
        ' Create the email body including the signature
        mailBody = mailBody & "<br><br>La dirección de entrega es la siguiente:<br><br><img src='cid:miImagen'><br><br>"
        mailBody = mailBody & "<br><br>Un saludo,<br><br>"
        
        mailBody = mailBody & "</body></html>"
        
        mailItem.htmlBody = mailBody
        
        ' Add PDF files to the email
        For i = 1 To recipientFiles(key).count
            mailItem.Attachments.Add recipientFiles(key)(i)
        Next i
        
        ' Display the email (do not send automatically without the user's personal check)
        mailItem.Display
    Next key
    
    ' Clean
    Set recipientFiles = Nothing
    Set recipientAlbaranes = Nothing
    Set fileSystem = Nothing
    Set folder = Nothing
    Set FileDialog = Nothing
    Set regex = Nothing
End Sub
