Sub MoveToDateFolder()
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim objInbox As Outlook.Folder
    Dim objFolder As Outlook.Folder
    Dim strFolderName As String
    Dim strYear As String
    Dim strMonth As String
    
    ' Get the currently selected item
    Set objItem = Application.ActiveExplorer.Selection.Item(1)
    
    ' Check if the selected item is a mail item
    If TypeOf objItem Is Outlook.MailItem Then
        Set objMail = objItem
        
        ' Get the received date of the email
        strFolderName = Format(objMail.ReceivedTime, "dd/mm/yyyy")
        strYear = Format(objMail.ReceivedTime, "yyyy")
        strMonth = Format(objMail.ReceivedTime, "mmmm")
        
        ' Convert the month to plain English
        strMonth = Application.WorksheetFunction.Text(objMail.ReceivedTime, "[$-409]mmmm")
        
        ' Set the Inbox folder as the starting point
        Set objInbox = Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox)
        
        ' Set the destination folder
        Set objFolder = GetOrCreateFolder(objInbox, strYear & "/" & strMonth & "/" & strFolderName)
        
        ' Move the email to the destination folder
        objMail.Move objFolder
    End If
    
    Set objMail = Nothing
    Set objFolder = Nothing
    Set objInbox = Nothing
    Set objItem = Nothing
End Sub

Function GetOrCreateFolder(parentFolder As Outlook.Folder, folderPath As String) As Outlook.Folder
    Dim objFolder As Outlook.Folder
    Dim folderNames() As String
    Dim i As Integer
    
    folderNames = Split(folderPath, "/")
    
    On Error Resume Next
    
    ' Iterate through each folder level in the folder path
    Set objFolder = parentFolder
    For i = LBound(folderNames) To UBound(folderNames)
        If Not folderNames(i) = "" Then
            Set objFolder = objFolder.Folders(folderNames(i))
            ' Create the folder if it doesn't exist
            If objFolder Is Nothing Then
                Set objFolder = objFolder.Folders.Add(folderNames(i))
            End If
        End If
    Next i
    
    Set GetOrCreateFolder = objFolder
    Set objFolder = Nothing
End Function
