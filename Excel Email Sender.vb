Sub SendEmails()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim cell As Range

    On Error GoTo ErrorHandler

    ' Create Outlook application
    Set OutApp = CreateObject("Outlook.Application")
    OutApp.Session.Logon

    ' Loop through each cell in the range
    For Each cell In Range("A2:A" & Cells(Rows.Count, 1).End(xlUp).Row)
        If cell.Value <> "" Then ' If the cell is not empty
            ' Create a new email
            Set OutMail = OutApp.CreateItem(0)
            
            ' Fill in email details
            With OutMail
                .To = cell.Value ' Receiver's email address from column A
                .Subject = cell.Offset(0, 1).Value ' Subject from column B
                .Body = cell.Offset(0, 2).Value ' Body from column C
                
                ' Extract sender's email address from column D
                Dim senderEmail As String
                senderEmail = cell.Offset(0, 3).Value
                
                ' Set sender's email address
                If senderEmail <> "" Then
                    .SentOnBehalfOfName = senderEmail
                Else
                    ' If sender's email address is blank, use default
                    ' Replace "default@email.com" with your default email address
                    .SentOnBehalfOfName = "r.r.8@pg.com"
                End If
                
                ' You can add more properties here such as CC, etc.
                ' For example:
                '.CC = "additional@email.com"
                '...

                ' Send the email
                .Send
            End With
            
            ' Clear the email object
            Set OutMail = Nothing
        End If
    Next cell

    ' Clean up
    Set OutApp = Nothing

    MsgBox "Emails Sent Successfully!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation
End Sub

