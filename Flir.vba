'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ______ _ _                _____  _____                      '
'   |  ___(_) |              / __  \|  _  |                     '
'   | |_   _| |_ __   __   __`' / /'| |/' |                     '
'   |  _| | | | '__|  \ \ / /  / /  |  /| |                     '
'   | |   | | | |      \ V / ./ /___\ |_/ /                     '
'   \_|   |_|_|_|       \_/  \_____(_)___/                      '
'                                                               '
'    @author    private@josh-king.co.uk                         '
'    @date      02/10/2018                                      '
'    @version   v2                                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   Application_Startup()                                                           '
'       Used for any maintenance functions that need to be run on application       '
'       initialisation. Currently used to create the utility folders and a          '
'       quick check to make sure that the custom drives are connected. Also         '
'       used to perform a cleanup of the users Outlook every Monday of the week     '
Sub Application_Startup()
    'Tidy up the debug console'
    'clearDebugConsole
    'Creates the [0 - Filed] & [0 - Sent] folders if they doesn't exist'
    CreateInboxFolder ("0 - Filed")
    CreateInboxFolder ("0 - Sent")
    
    'Checks to make sure that the Y/M drive is valid and ready'
    If Not DoesDriveExist("Y") Then
        PingMessage ("ERROR 001: [Y:\] Drive not found")
        End
    End If
    
    If Not DoesDriveExist("M") Then
        PingMessage ("ERROR 001: [M:\] Drive not found")
        End
    End If
    
    'If today is a Monday then perform cleanup'
    If Weekday(Now(), vbMonday) = 1 Then
        PerformCleanUp
    End If

    DebugPrintOut ("Outlook initialised correctly.")
End Sub

'   ManualFilingTrigger()                                                                  '
'       Basic sub function so it can be called from the macro and then pass the     '
'       argument to the main FilingEngine Sub                                       '
Sub ManualFilingTrigger()
    FilingEngine 0
End Sub

'   SentFilingTrigger()                                                                  '
'       Basic sub function so it can be called from the macro and then pass the     '
'       argument to the main FilingEngine Sub                                       '
Sub SentFilingTrigger()
    FilingEngine 1
End Sub

'   FilingEngine(typeOfFiling[integer])                                             '
'       The engine behind the filing macros - this is the re-written version to make'
'       it as object oriented as possible however VBA is very restrictive as a base '
'       and doesn't allow a lot of polymorphic techniques                           '
'           @typeOfFiling - Takes a 1 or a 0 depending on the type of filing the    '
'                           use wants to perform (this changes the selection)       '
Sub FilingEngine(ByVal typeOfFiling As Integer)
    Dim emailObj, inboxFolder As Object, currentFolder As Object, objFiledFldr As Object, objSntFldr As Object, explorerObj As Object, selectionObj As Object, nameSpaceObj As Object, folderObj As Object
    Dim counter As Long, matterCounter As Long, totalSelection As Long, i As Integer
    Dim subject As String, matterNo As String, saveLocation As String, e_type As String, sentOnFormatted As String, displayName As String, senderName As String, recipientName As String
    
    counter = 0
    matterCounter = 0
    Set nameSpaceObj = Application.GetNamespace("MAPI")
    Set folderObj = nameSpaceObj.GetDefaultFolder(olFolderInbox)
    Set objFiledFldr = folderObj.Folders("0 - Filed")
    Set objSntFldr = folderObj.Folders("0 - Sent")

    If typeOfFiling = 0 Then
        Set explorerObj = Application.ActiveExplorer
        Set selectionObj = explorerObj.Selection

        totalSelection = selectionObj.Count
    Else
        totalSelection = objSntFldr.Items.Count
    End If

    DebugPrintOut ("Found " & totalSelection & " email(s).")

    ProgressForm.Progressbar_amount.Caption = counter & "/" & totalSelection
    ProgressForm.progressBar.Width = 0
    ProgressForm.Show

    For i = totalSelection To 1 Step -1
    On Error GoTo SelectionHandler
        If (typeOfFiling = 0) Then
            Set emailObj = selectionObj.Item(i)
        Else
            Set emailObj = objSntFldr.Item(i)
        End If
        
        If TypeOf emailObj Is MailItem Then
            'Retrieve information from the emailObj'
            matterNo = ExtractMatter(StripIllegalCharacters(emailObj.subject))
            senderName = emailObj.senderName
            recipientName = emailObj.Recipients(1)
            'Format the date'
            sentOnFormatted = Format(emailObj.sentOn, "yyyy-mm-dd hh-mm-ss")
            
            'Check if the email was received or sent and then change the e_type'
            If InStr(recipientName, "The Partnership") = 0 Then
                e_type = "S"
                displayName = StripIllegalCharacters(recipientName)
            Else
                e_type = "R"
                displayName = StripIllegalCharacters(senderName)
            End If
            
            'Simple check if there is a matterNo within the subject line'
            If matterNo <> "" Then
                If DoesMatterExist(matterNo) Then
                    SaveEmail emailObj, matterNo, sentOnFormatted, e_type, displayName, objFiledFldr
                    
                    'Counter for how many matters have been processed'
                    matterCounter = matterCounter + 1
                End If
            
                If Not DoesMatterExist(matterNo) Then
                    PingMessage "ERROR 003: " & matterNo & " does not exist. Skipped."
                End If
            'If theres no matter number in subject'
            Else
                 If Not MultipleForm.confirm_label = "skip" Then
                    'Call a userform for userinput
                    PopulateFilingForm (emailObj)
                    matterNo = FilingForm.inputted_matter.Text
                    FilingForm.inputted_matter.Text = ""
    
                    'Initial check if the matter Exists'
                    If Not DoesMatterExist(matterNo) Then
                        PingMessage "ERROR 0063: " & matterNo & " does not exist. Try again."
                        PopulateFilingForm (emailObj)
                        matterNo = FilingForm.inputted_matter.Text
                        FilingForm.inputted_matter.Text = ""
                    End If
    
                    'Secondary check if the matter Exists'
                    If Not DoesMatterExist(matterNo) Then
                        PingMessage "ERROR 003: " & matterNo & " does not exist. Skipped."
                    Else
                        If FilingForm.confirm_label = "add" Then
                            SaveEmail emailObj, matterNo, sentOnFormatted, e_type, displayName, objFiledFldr

                            'Counter for how many matters have been processed'
                            matterCounter = matterCounter + 1
                        Else
                            DebugPrintOut ("Successfully skipped e-mail.")
                        End If
                    End If
                End If
            End If
        End If
        
        If Not TypeOf emailObj Is MailItem Then
            PopulateFilingFormErr (emailObj)
        End If
        
        'Counter for how many files have been processed'
        counter = counter + 1
        
        'Make the progress bar update()'
        ProgressForm.Progressbar_amount.Caption = counter & "/" & totalSelection
        ProgressForm.progressBar.Width = Round(((220 * counter * 100) / totalSelection) / 100)
        ProgressForm.Repaint
    Next
    
    ProgressForm.Hide
    DebugPrintOut ("Selected files processed! (" & matterCounter & "/" & counter & ")")
    GoTo FinishedFiling
    
SelectionHandler:
    'Error 0 is macro ran fine'
    If (Err.Number <> 0) Then
        DebugPrintOut "Error number: " & Err.Number & " " & Err.Description
        PingMessage ("ERROR 002: Failed when building collection")
        End
    End If
FinishedFiling:
End Sub

'   PerformCleanUp()                                                                '
'       Cleans up the users filed and sent folder. The conditions below are for each'
'       folder - using the iteration backwards which I will use throughout          '
'           @objFiledFldr - If the email is older than 7 days then it will be       '
'                           deleted, if it is younger then it will be set to read   '
'           @objSntFldr - If the email is older than 14 days then it will be        '
'                         deleted, if it is younger then it will be set to read     '
Sub PerformCleanUp()
    Dim nameSpaceObj As Outlook.NameSpace, folderObj As Outlook.Folder, objFiledFldr As Outlook.Folder, objSntFldr As Outlook.Folder
    Set nameSpaceObj = Application.GetNamespace("MAPI")
    Set folderObj = nameSpaceObj.GetDefaultFolder(olFolderInbox)
    Set objFiledFldr = folderObj.Folders("0 - Filed")
    Set objSntFldr = folderObj.Folders("0 - Sent")

    'Iterates from the end backwards - you will see me doing this a lot'                                                                                                                         '
    For i = objFiledFldr.Items.Count To 1 Step -1
        Set emailObj = objFiledFldr.Items(i)
        'If they are older than 7 days then delete, otherwise set as read'
        If DateDiff("d", emailObj.CreationTime, Date) > 7 Then
            DebugPrintOut emailObj.subject & " [> 7 Days]"
            emailObj.Delete
        Else
            emailObj.UnRead = False
        End If
    Next

    For i = objSntFldr.Items.Count To 1 Step -1
        Set emailObj = objSntFldr.Items(i)
        If DateDiff("d", emailObj.CreationTime, Date) > 14 Then
            DebugPrintOut emailObj.subject & " [> 14 Days]"
            emailObj.Delete
        Else
            emailObj.UnRead = False
        End If
    Next
    
    DebugPrintOut ("Cleaned the users Folders.")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ______ _   _ _   _ _____ _____ _____ _____ _   _  _____     '
'   |  ___| | | | \ | /  __ \_   _|_   _|  _  | \ | |/  ___|    '
'   | |_  | | | |  \| | /  \/ | |   | | | | | |  \| |\ `--.     '
'   |  _| | | | | . ` | |     | |   | | | | | | . ` | `--. \    '
'   | |   | |_| | |\  | \__/\ | |  _| |_\ \_/ / |\  |/\__/ /    '
'   \_|    \___/\_| \_/\____/ \_/  \___/ \___/\_| \_/\____/     '
'                                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'   SaveEmail [emailObj(object), matterNo(string), date(string),                '
'              e_type(String), displayName(string), objFiledFldr(object)]       '
'       Save email is a specific function that formats the file name of the     '
'       email calls the GenerateFiles function and moves the emails once the    '
'       emailObj has been saved properly                                        '
Function SaveEmail(ByVal emailObj As MailItem, matterNo As String, parsedDate As String, e_type As String, displayName As String, objFiledFldr As Object)
    'Creates the saveName path'
    Dim saveName As String
    saveName = "M:\" & matterNo & "@@" & parsedDate & "@@" & e_type & "@@" & displayName
    
    'Checks if there is already a file with the same name, and then adds a 1'
    If DoesFileExist(saveName, ".msg") Then
        parsedDate = Format(DateAdd("s", 1, parsedDate), "yyyy-mm-dd hh-mm-ss")
        saveName = "M:\" & matterNo & "@@" & parsedDate & "@@" & e_type & "@@" & displayName
        DebugPrintOut ("Duplicate file detected, applying change [" & saveName & "].")
    End If
    
    'Save all the attachments from the email if the email is recieved'
    If e_type = "R" Then
        SaveAttachments emailObj, matterNo
    End If
    
    'Generate the files needed to be saved'
    GenerateFiles saveName, emailObj, e_type, matterNo

    'Move the email to the 0 - Filed folder as long as it's not skipped'
    emailObj.UnRead = False
    emailObj.Move objFiledFldr
End Function

'   ExtractMatter [search(string)]                                              '
'       Function that uses REGEX to check if a string contains a MatterNo       '
'           ([A-Z]{3}\d{4}\s)   |   (\w+.[A-Z]\d+)                              '
Function ExtractMatter(ByVal search As String) As String
On Error GoTo MatterException
    Dim result As String, matterMatches As Object, regEx As Object
    Set regEx = CreateObject("vbscript.regexp")

    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "([A-Z]{3}\d{4})\b"
    End With

    Set matterMatches = regEx.Execute(search)

    If matterMatches.Count <> 0 Then
        If matterMatches.Count > 1 Then
            'Call a userform for userinput
            MultipleForm.Show
            result = MultipleForm.inputted_matter.Text
            MultipleForm.inputted_matter.Text = ""
        Else
            result = matterMatches.Item(0).submatches.Item(0)
        End If
    End If

    'Check if the extracted matter is on the other side'
    If IsOtherSideMatter(result) Then
        MultipleForm.Caption = "Acting on both sides"
        MultipleForm.information = "Acting on both sides detected (" & result & "), please confirm the matter number below."
        MultipleForm.Show
        
        result = MultipleForm.inputted_matter.Text
        MultipleForm.inputted_matter.Text = ""
        MultipleForm.Caption = "Multiple Matters Detected"
        MultipleForm.information = "Multiple matter numbers have been detected in this email, please confirm the matter number below."
    End If

    ExtractMatter = result
    DebugPrintOut ("Found [" & result & "] in search from " & matterMatches.Count & " match(es).")
GoTo FoundMatter
MatterException:
    If (Err.Number <> 0) Then
        DebugPrintOut "Error number: " & Err.Number & " " & Err.Description
        PingMessage "ERROR 004: Failed to check for MatterNo."
        End
    End If
FoundMatter:
End Function

'   PopulateFilingForm [emailObj(object)]                                       '
'       Used to fill in the blanks for the filing form with information         '
'       from the emailObj                                                       '
Function PopulateFilingForm(ByVal emailObj As MailItem)
    FilingForm.email_sentFrom_label = emailObj.senderName
    FilingForm.email_subject_label = emailObj.subject
    FilingForm.email_folder.Caption = emailObj.Parent
    FilingForm.conversationID_label.Caption = emailObj.ConversationID
    FilingForm.conversationIndex_label.Caption = emailObj.ConversationIndex
    FilingForm.email_sentTo_label = emailObj.Recipients(1)
    FilingForm.email_date_label = emailObj.sentOn
    FilingForm.Show
End Function

'   GenerateFiles [location(string), emailObj(object), e_type(string)]          '
'       Function that is used to generate the 3 files needed (XML, MSG, MHTML)  '
Function GenerateFiles(ByVal location As String, ByVal emailObj As Object, ByVal e_type As String, ByVal matterNo As String)
    Dim fsoObj As Object, textObj As Object

On Error GoTo GenerateFilesHandler
    'Save the emails in a certain location'
    emailObj.SaveAs location & ".msg", olMSG
    emailObj.SaveAs location & ".mht", olMHTML

    'Create a file writer instance'
    Set fsoObj = CreateObject("scripting.filesystemobject")
    Set textObj = fsoObj.CreateTextFile(location & ".xml", True)
    textObj.WriteLine (GenerateXML(emailObj, matterNo, e_type))
    'Make sure to close it!'
    textObj.Close

    DebugPrintOut ("Files generated correctly.")
    GoTo GeneratedFiles
GenerateFilesHandler:
    If (Err.Number <> 0) Then
        DebugPrintOut "Error number: " & Err.Number & " " & Err.Description
        PingMessage "ERROR 005: Could not generate files"
        End
    End If
GeneratedFiles:
End Function

'Function that generates an XML file, currently the fasest way to do it       '
'       GenerateXML [emailObj(object), matterNo(string), e_type(String)]   '
'               String with correct XML tags                                  '
Function GenerateXML(ByVal emailObj As Object, ByVal matterNo As String, ByVal e_type As String) As String
    Dim attachmentName As String, html As String, attachmentCounter As Integer, hasAttachment As Integer, attachmentObj As Attachment
    
    attachmentCounter = 0
    hasAttachment = 0

    'Loop though each attachment'
    For Each attachmentObj In emailObj.Attachments
        'Get the attachment name & HTML body'
        attachmentName = attachmentObj.FileName
        html = emailObj.HTMLBody

        'Search the HTML body for attachment names'
        If InStr(html, attachmentName) = 0 Then
            'if it has found the attachment then add 1'
            attachmentCounter = attachmentCounter + 1
        End If
    Next attachmentObj

    'If the counter is greater than one then there is an actual attachment'
    If attachmentCounter > 0 Then
        hasAttachment = 1
    End If

    GenerateXML = "<?xml version='1.0' encoding='utf-8'?>" & vbNewLine _
                & "<email Version='1.1'>" & vbNewLine _
                & "  <subject>" & StripIllegalCharacters(emailObj.subject) & "</subject>" & vbNewLine _
                & "  <recipient_name>" & StripIllegalCharacters(emailObj.Recipients(1)) & "</recipient_name>" & vbNewLine _
                & "  <recipient_address>" & StripIllegalCharacters(emailObj.Recipients.Item(1).Address) & "</recipient_address>" & vbNewLine _
                & "  <sender_name>" & StripIllegalCharacters(emailObj.senderName) & "</sender_name>" & vbNewLine _
                & "  <sender_address>" & StripIllegalCharacters(emailObj.SenderEmailAddress) & "</sender_address>" & vbNewLine _
                & "  <cc>" & StripIllegalCharacters(emailObj.CC) & "</cc>" & vbNewLine _
                & "  <date>" & StripIllegalCharacters(emailObj.sentOn) & "</date>" & vbNewLine _
                & "  <attachment>" & hasAttachment & "</attachment>" & vbNewLine _
                & "  <importance>" & StripIllegalCharacters(emailObj.Importance) & "</importance>" & vbNewLine _
                & "  <sendreceive>" & e_type & "</sendreceive>" & vbNewLine _
                & "  <body>" & StripIllegalCharacters(emailObj.Body) & "</body>" & vbNewLine _
                & "</stops_email>"
End Function

'   SaveAttachments (emailObj[object], matterNo[string])                    '
'       Loops through all the attachments that are linked to the emailObj   '
'       and then saves them into the users upload directory (N:\)           '
Function SaveAttachments(ByVal emailObj As Object, ByVal matterNo As String)
    Dim objAttachments As Attachments, attachmentCount As Integer, i As Integer, attachmentName As String, upload As String
    
    Set objAttachments = emailObj.Attachments
    attachmentCount = objAttachments.Count
    
    'Start a loop that is the same amount as the attachmentCount'
    For i = attachmentCount To 1 Step -1
        attachmentName = matterNo & " - " & objAttachments.Item(i).FileName
        
        If DoesDriveExist("U") Then
            'Use the integer counter for the item index'
            upload = "N:\" & attachmentName
            If Not DoesFileExist(upload, "") Then
                objAttachments.Item(i).SaveAsFile (upload)
            Else
                DebugPrintOut ("[" & upload & "] Attachment already exists.")
            End If

            DebugPrintOut (attachmentName & " saved. [" & i & "/" & attachmentCount & "]")
        End If
    Next

    DebugPrintOut attachmentCount & " attachment(s) saved."
End Function

'   DoesMatterExist (matter[string])                                    '
'       Function that is used to check if the matter exists             '
Function DoesMatterExist(ByVal matter As String) As Boolean
    Dim location As String, fsoObj As Object
    Set fsoObj = CreateObject("Scripting.FileSystemObject")
    
    location = "Y:\" & matter
    DoesMatterExist = fsoObj.folderexists(location)
End Function

'   DoesFileExist (location[string], extension[string])                 '
'       Function that is used to check if there are any existing files  '
Function DoesFileExist(ByVal location As String, ByVal extension As String) As Boolean
    Dim fsoObj As Object
    Set fsoObj = CreateObject("Scripting.FileSystemObject")
    DoesFileExist = fsoObj.FileExists(location & extension)
End Function

'   StripIllegalCharacters [search(string)]                                '
'        Function that uses REGEX to strip a string of invalid windows     '
'        file name                                                         '
Function StripIllegalCharacters(search As String) As String
    'The regex pattern to find special characters
    Dim strPattern As String: strPattern = "[^\w\s\.@-]"
    'The replacement for the special characters
    Dim strReplace As String: strReplace = ""
    Dim regEx As Object
    Set regEx = CreateObject("vbscript.regexp")

    ' Configure the regex object
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = strPattern
    End With

    ' Perform the regex replacement
    StripIllegalCharacters = regEx.Replace(search, strReplace)
End Function

'   CreateInboxFolder (inboxFolder[string])                                 '
'       Checks if the folder you are search is in the users Inbox if not    '
'       then create the folder                                              '
Function CreateInboxFolder(ByVal inboxFolder As String)
    Dim nameSpaceObj As Outlook.NameSpace, folderObj As Outlook.Folder, newFolderObj As Outlook.Folder
    Set nameSpaceObj = Application.GetNamespace("MAPI")
    Set folderObj = nameSpaceObj.GetDefaultFolder(olFolderInbox)
    
On Error GoTo CreateFileSkip
    'Create a new inbox folder if one is not present'
    Set newFolderObj = folderObj.Folders.Add(inboxFolder)
    DebugPrintOut ("Created new inbox folder [" & inboxFolder & "].")
CreateFileSkip:
End Function

'   IsOtherSideMatter [matter(string)]                                           '
'       Function that checks if the matterNo is a case where we are acting       '
'       on both sides                                                            '
Function IsOtherSideMatter(ByVal matter As String) As Boolean
    Dim location As String, fsoObj As Object
    Set fsoObj = CreateObject("Scripting.FileSystemObject")
    
    location = "Y:\" & matter & "\otherside.txt"
    IsOtherSideMatter = fsoObj.FileExists(location)
    DebugPrintOut ("Other side status [" & fsoObj.FileExists(location) & ")")
End Function

'   DebugPrintOut (output[string])                                          '
'       Prints the output with the prefixed time for easier debugging       '
Function DebugPrintOut(ByVal output As String) As String
    Debug.Print "[Filr(" & Now & ")]: " & output
End Function

'   PingMessage (message[string])                                           '
'       Pings a message to the user if something has gone wrong or more     '
'       information is need to be inputted - Similar to DebugPrintOut but   '
'       has less information and is for a quick feedback only               '
Function PingMessage(ByVal message As String)
    DebugPrintOut (message)
    MsgBox message, vbOKOnly, "Filing Operation"
End Function

'   DoesDriveExist (driveLetter[string])                                    '
'       Checks to make sure the user drive exists if so then return true    '
Function DoesDriveExist(ByVal driveLetter As String) As Boolean
    Dim fsoObj As Object
    Set fsoObj = CreateObject("Scripting.FileSystemObject")

    'Check if the FSObject can see the drive'
    DoesDriveExist = IIf(fsoObj.DriveExists(driveLetter), True, False)
End Function

'   clearDebugConsole ()                                                    '
'       Clears the debug console by just placing in a bunch of spaces       '
'       hey - if it works it works..                                        '
Function clearDebugConsole()
    For i = 0 To 20
        Debug.Print ""
    Next i
    Debug.Print ("---------- Cleared ----------")
End Function
