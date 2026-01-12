'------------------------------------------------------------
' Macro: CorreosUltimacion
' Description:
'   The user selects a folder with transit PDF
'   files and their corresponding EPOD (Proof of Delivery
'   documents).
'
'   The macro identifies recipients and delivery note
'   numbers from file names, processes only shipments
'   with EPOD (delivered), and generates one email per 
'   recipient with the relevant attachments.
'
'   Email subject and body are adjusted based on the
'   expiration date extracted from the folder name.
'   Emails are displayed for review, not sent
'   automatically.
'
' Author: ssalgado0@uoc.edu
'------------------------------------------------------------

Sub CorreosUltimacion()

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
    Dim i As Integer
    Dim recipientFiles As Object
    Dim recipientAlbaranes As Object
    Dim key As Variant
    Dim mailBody As String
    Dim albaranNumber As String
    Dim regex As Object
    Dim matches As Object
    Dim miImagen As String
    Dim folderName As String
    Dim fechaCaducidad As String
    Dim anoActual As Integer
    
    ' Initialize Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    ' Show dialog to select a folder
    Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    If FileDialog.Show = -1 Then
        folderPath = FileDialog.SelectedItems(1)
    Else
        MsgBox "No se ha seleccionado ninguna carpeta.", vbExclamation
        Exit Sub
    End If
    
    ' Create FileSystem object to work with files and folders
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    folderName = fileSystem.GetFileName(folderPath)
    fechaCaducidad = Mid(folderName, InStr(folderName, " ") + 1)
    
    ' Adjustment so the macro works correctly during the first days of year change
    If Month(Date) = 12 And Day(Date) >= 26 Then
        anoActual = Year(Date) + 1
    Else
        anoActual = Year(Date)
    End If
        
    ' Define recipients list
    recipientsList = Array _
    ("RECIPIENT NAME ONE", _
    "RECIPIENT NAME TWO", _
    "RECIPIENT NAME THREE", _
    "RECIPIENT NAME FOUR", _
    "RECIPIENT NAME FIVE")

    ' Create dictionaries to store files and delivery notes per recipient
    Set recipientFiles = CreateObject("Scripting.Dictionary")
    Set recipientAlbaranes = CreateObject("Scripting.Dictionary")

    ' Create Regex object to find 10-digit numbers
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\b\d{10}\b"
    regex.Global = True

    ' Loop through each file in the selected folder
    For Each file In folder.Files

        ' Variable to check if shipment has an EPOD (already delivered)
        Dim tieneEpod As Boolean
        tieneEpod = False

        ' Check only PDF files
        If LCase(fileSystem.GetExtensionName(file.Name)) = "pdf" Then
            fileName = file.Name
            recipientFound = False
            processedRecipient = ""
    
            ' Adjust file name for special errors before searching
            If InStr(1, fileName, "RECIPIE NAME ONE", vbTextCompare) > 0 Then
                processedRecipient = "RECIPIENT NAME ONE"
                recipientFound = True
            ElseIf InStr(1, fileName, "ECIPIENT NAME TWO", vbTextCompare) > 0 Then
                processedRecipient = "RECIPIENT NAME TWO"
                recipientFound = True
            ElseIf InStr(1, fileName, "RECIPIENT NAME & FIVE", vbTextCompare) > 0 Then
                processedRecipient = "RECIPIENT NAME FIVE"
                recipientFound = True
            ElseIf InStr(1, fileName, "RECIPIENT & FIVE", vbTextCompare) > 0 Then
                processedRecipient = "RECIPIENT NAME FIVE"
                recipientFound = True
            Else
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
                Set matches = regex.Execute(fileName)
                If matches.count > 0 Then
                    albaranNumber = matches.item(0).Value
    
                    ' Search for related EPOD files in the same folder
                    For Each otherFile In folder.Files
                        If LCase(fileSystem.GetExtensionName(otherFile.Name)) = "pdf" Then
                            otherFileName = otherFile.Name
                            ' Naming error handling
                            If InStr(1, otherFileName, albaranNumber & " EPOD", vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, "EPOD " & albaranNumber, vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, albaranNumber & " Epod", vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, "Epod " & albaranNumber, vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, albaranNumber & " ePOD", vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, "ePOD " & albaranNumber, vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, albaranNumber & " epod", vbTextCompare) > 0 Or _
                               InStr(1, otherFileName, "epod " & albaranNumber, vbTextCompare) > 0 Then

                               ' Stop searching once an EPOD is found
                               tieneEpod = True
                               Exit For 
                            End If
                        End If
                    Next otherFile
    
                    ' Continue only if EPOD exists
                    If tieneEpod Then
                        ' Create collections for the recipient if they don’t exist
                        If Not recipientFiles.exists(processedRecipient) Then
                            recipientFiles.Add processedRecipient, New Collection
                            recipientAlbaranes.Add processedRecipient, New Collection
                        End If

                        ' Add main PDF file and delivery note number
                        recipientFiles(processedRecipient).Add file.Path
                        recipientAlbaranes(processedRecipient).Add albaranNumber
    
                        ' Add EPOD file to the same recipient
                        For Each otherFile In folder.Files
                            If LCase(fileSystem.GetExtensionName(otherFile.Name)) = "pdf" Then
                                otherFileName = otherFile.Name
                                ' Error handling
                                If InStr(1, otherFileName, albaranNumber & " EPOD", vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, "EPOD " & albaranNumber, vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, albaranNumber & " Epod", vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, "Epod " & albaranNumber, vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, albaranNumber & " ePOD", vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, "ePOD " & albaranNumber, vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, albaranNumber & " epod", vbTextCompare) > 0 Or _
                                   InStr(1, otherFileName, "epod " & albaranNumber, vbTextCompare) > 0 Then
                                   recipientFiles(processedRecipient).Add otherFile.Path
                                End If
                            End If
                        Next otherFile
                    End If
                End If
            End If
        End If
    Next file

    
    
    ' Create and display emails for each recipient
    For Each key In recipientFiles.Keys
        ' Create a new email
        Set mailItem = OutlookApp.CreateItem(0)

        ' Subject naming and special case handling
        If key = "RECIPIENT NAME THREE" Then
            mailItem.Subject = "TRÁNSITOS " & key & " EXTRA DATA" & " ***ULTIMACIÓN***"
        Else
            mailItem.Subject = "TRÁNSITOS " & key & " ***ULTIMACIÓN***"
        End If

        ' Build email body
        mailBody = "<html><body>Buenos días,<br><br>" & _
                   "Adjuntamos tránsito y EPOD de los siguientes envíos para su ultimación<br><br>"
                   
        ' Add delivery note numbers to the email body
        If recipientAlbaranes.exists(key) Then
            For i = 1 To recipientAlbaranes(key).count
                mailBody = mailBody & recipientAlbaranes(key)(i) & "<br>"
            Next i
        End If
        
        ' Assign recipients depending on destination
        Select Case key
            Case "RECIPIENT NAME ONE"
                mailItem.To = "recip1name1@companyname.com; recip2name1@companyname.com"
                mailItem.CC = "cc_recip1name1@companyname.com; cc_recip2name1@companyname.com; cc_recip3name1@companyname.com"
            Case "RECIPIENT NAME TWO"
                mailItem.To = "recip1name2@companyname.com; recip2name2@companyname.com"
                mailItem.CC = "cc_recip1name2@companyname.com; cc_recip2name2@companyname.com; cc_recip3name2@companyname.com"
            Case "RECIPIENT NAME THREE"
                mailItem.To = "recip1name3@companyname.com; recip2name3@companyname.com"
                mailItem.CC = "cc_recip1name3@companyname.com; cc_recip2name3@companyname.com; cc_recip3name3@companyname.com"
            Case "RECIPIENT NAME FOUR"
                mailItem.To = "recip1name4@companyname.com; recip2name4@companyname.com"
                mailItem.CC = "cc_recip1name4@companyname.com; cc_recip2name4@companyname.com; cc_recip3name4@companyname.com"
            Case "RECIPIENT NAME FIVE"
                mailItem.To = "recip1name5@companyname.com; recip2name5@companyname.com"
                mailItem.CC = "cc_recip1name5@companyname.com; cc_recip2name5@companyname.com; cc_recip3name5@companyname.com"
            ' Add more cases when needed
            Case Else
                mailItem.To = "default@example.com"
                mailItem.CC = "default_cc1@example.com; default_cc2@example.com"
        End Select

        ' Build email body with expiration date information
        Dim fechaConversion As String
        fechaConversion = fechaCaducidad & "-" & anoActual
        
        Dim fechaCaducidadConverted As Date
        fechaCaducidadConverted = CDate(fechaConversion)
        
        ' If expiration date is today
        If fechaCaducidadConverted - Date = 0 Then
            mailBody = mailBody & "<br><br>Fecha Límite " & "<strong style='background-color:IndianRed'>" & "HOY " & fechaCaducidad & "</strong><strong style='background-color:IndianRed'>" & "-" & anoActual & "</strong><br>"
            
                    If key = "RECIPIENT NAME THREE" Then
                        mailItem.Subject = "TRÁNSITOS " & key & " EXTRA DATA" & " ***ULTIMACIÓN URGENTE HOY***"
                    Else
                        mailItem.Subject = "TRÁNSITOS " & key & " ***ULTIMACIÓN URGENTE HOY***"
                    End If
        ' If expired
        ElseIf fechaCaducidadConverted - Date < 0 Then
            mailBody = mailBody & "<br><br><strong>CADUCADOS.</strong> Fecha Caducidad " & "<strong style='background-color:IndianRed'>" & fechaCaducidad & "</strong><strong style='background-color:IndianRed'>" & "-" & anoActual & "</strong><br>"
        ' Expiration date is in 1 or 2 days
        ElseIf fechaCaducidadConverted - Date = 1 Or fechaCaducidadConverted - Date = 2 Then
            mailBody = mailBody & "<br><br>Fecha Límite " & "<strong style='background-color:Gold'>" & fechaCaducidad & "</strong><strong style='background-color:Gold'>" & "-" & anoActual & "</strong><br>"
        ' Expiration date is 3 or more days away
        ElseIf fechaCaducidadConverted - Date >= 3 Then
            mailBody = mailBody & "<br><br>Fecha Límite " & "<strong style='background-color:SkyBlue'>" & fechaCaducidad & "</strong><strong style='background-color:SkyBlue'>" & "-" & anoActual & "</strong><br>"
        End If
 
        ' Add disclaimer text
        mailBody = mailBody & "<br><br>En caso de que la fecha esté al límite:<br>"
        mailBody = mailBody & "No asumimos sanción<br>"
        mailBody = mailBody & "<br><br>Muchas gracias,<br>"
        mailBody = mailBody & "Un saludo<br><br><br>"
        
        mailBody = mailBody & "</body></html>"
        
        mailItem.htmlBody = mailBody
        
        ' Attach PDF files
        For i = 1 To recipientFiles(key).count
            mailItem.Attachments.Add recipientFiles(key)(i)
        Next i
        
        ' Display email (no autosend)
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
