Attribute VB_Name = "SaveEmail"

Option Explicit
'======================================================================================='

' Declare ShellExecute API for opening Obsidian links
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub ExtractEmail_MarkDown()

    ' Dim ObsidianFolder As String
    ' Dim baseFolder As String
    Dim vaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String
    Dim optionValue As Integer
    Dim mailItem As Outlook.mailItem
    Dim htmlContent As String
    Dim processedHtml As String
        
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink

    '================================================'
    ' Save as plain text
    Const OLTXT = 0
    ' Object holding variable
    Dim obj As Object
    ' Instantiate an Outlook Email Object
    Dim oMail As Outlook.mailItem
    ' If something breaks, skip to the end, tidy up and shut the door
    On Error GoTo EndClean:
    
    ' Establish the environment and selected items (emails)
    ' NOTE: selecting a conversation-view stack wont work
    '       you'll need to select one of the emails
    Dim fileName As String, mName As String
    Dim temporarySubjectLineString As String
    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection
    ' For each email in the Selection
    ' Assigning email item to the `obj` holding variable
    For Each obj In Selection
        ' set the oMail object equal to that mail item
        Set oMail = obj
        ' Is it an Email?
        If oMail.Class <> 43 Then
          MsgBox "This code only works with Emails."
          GoTo EndClean: ' you broke it
        End If

        ' Yank the mail items subject line to `temporarySubjectLineString`
        temporarySubjectLineString = oMail.subject
        ' function call the name cleaner to remove any
        '    illegal characters from the subject line
        ReplaceCharsForFileName temporarySubjectLineString, ""
        ' Yank the received date-time to a holding variable

        ' Build Recipient string based on receipient collection
        Dim recips As Outlook.Recipients
            Set recips = oMail.Recipients
        Dim recip As Outlook.Recipient
        Dim result As String
        Dim recipString As String
            recipString = ""

        For Each recip In recips
            recipString = recipString & vbTab
            recipString = recipString & "- "
            recipString = recipString & formatName(recip.Name, personNameStartChar)
            recipString = recipString & vbCrLf
        Next
        
        ' ===============================================================
        
        ' (1) 아웃룩 파일 html으로 저장
        Dim objItem As mailItem, htmlpath As String
        Set objItem = Application.ActiveExplorer.Selection(1)
        htmlpath = vaultPathToSaveFileTo & temporarySubjectLineString & ".html"
        ' Debug.Print htmlpath
        objItem.SaveAs htmlpath, 5
        
        
        ' (2) 첨부파일 저장
        Dim attachments As Outlook.attachments
        Dim Attachment As Outlook.Attachment
        Set attachments = objItem.attachments
        Dim i As Long
        For i = 1 To attachments.Count
            Set Attachment = attachments(i)
            Dim attachFileName As String
            attachFileName = vaultPathToSaveFileTo & temporarySubjectLineString & ".files\" & Attachment.fileName
            Attachment.SaveAsFile attachFileName
        Next i
        
        
        ' Build the result file content to be sent to the mail item body
        ' Then save that mail item same as the meeting extractor
        Dim sender As String
            sender = formatName(oMail.sender, personNameStartChar)
        Dim dtDate As Date
            dtDate = oMail.ReceivedTime
        Dim resultString As String

        ' resultString = ""
        ' resultString = resultString & "# [[" & emailFileNameStartChr & Format(oMail.ReceivedTime, "yyyy-mm-dd hhnn") & " " & temporarySubjectLineString & "|" & temporarySubjectLineString & "]]"
        ' resultString = resultString & vbCrLf & vbCrLf & vbCrLf
        
        ' (3) 프로퍼티 표시(YAML frontmatter)
        ' ----------------------------------------------- Properties star -----------------------------------------------------------
        resultString = "---" & vbCrLf
        
        ' Add tags
        resultString = resultString & "tags:  " & """SOURCE/MAIL today""" & vbCrLf
        resultString = resultString & "Index: " & vbCrLf
        resultString = resultString & "title: """ & temporarySubjectLineString & """" & vbCrLf
        resultString = resultString & "aliases:" & vbCrLf
        resultString = resultString & "create: " & Format(Now, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "수신일시: " & Format(oMail.SentOn, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "요청일자: " & Format(oMail.SentOn, "yyyy-MM-dd") & vbCrLf
        resultString = resultString & "요청자: """ & sender & """" & vbCrLf
        resultString = resultString & "진행상태: " & "대기" & vbCrLf
        resultString = resultString & "D-day: """ & vbCrLf
        resultString = resultString & "완료일: """ & vbCrLf
        resultString = resultString & "ITSM: """ & vbCrLf
        resultString = resultString & "ITSM_URL: """ & vbCrLf
        
        ' Convert recipients to YAML list
        ' resultString = resultString & "to:" & vbCrLf
        ' For Each recip In recips
        '     resultString = resultString & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
        ' Next

        ' End YAML block
        resultString = resultString & "---" & vbCrLf
        ' ----------------------------------------------- Properties end ------------------------------------------------------------
        
        ' Now we create the file name
        ' mName = emailFileNameStartChr
        mName = mName & Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, vbUseSystem)
        ' mName = mName & Format(dtDate, " hhMM", vbUseSystemDayOfWeek, vbUseSystem)
        mName = mName & " " & temporarySubjectLineString
 
 
        ' (4) 첨부파일 표시
        Dim propertyAccessor As Outlook.propertyAccessor
        Dim propHidden As String
        Dim isHidden As Variant
        propHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
        For i = 1 To attachments.Count
            Dim fNm As String
            Dim fLnk As String
            Dim fLst As String
       
            Set Attachment = attachments(i)
            Set propertyAccessor = Attachment.propertyAccessor
            isHidden = propertyAccessor.GetProperty(propHidden)
       
            ' PR_ATTACHMENT_HIDDEN 속성이 True인 경우에만 보여줌
            If Not isHidden Then
            fNm = Attachment.fileName
                fLnk = Replace("![[" & baseFolder & temporarySubjectLineString & ".files\" & fNm & "]]", "\", "/")
                resultString = resultString & fLnk & vbCrLf
            End If
        Next i
        
        
        ' Add a horizontal rule for separation
        resultString = resultString & vbCrLf & "---" & vbCrLf
        
        
        ' (5) 메일 본문 표시
        ' ------------------------------------------------- Body Start --------------------------------------------------------------
        ' (6) 소스 표시
        resultString = resultString & "![[" & temporarySubjectLineString & ".html]]" & vbCrLf
        resultString = resultString & "[[" & mName & "]]" & vbCrLf
        
        ' (7) 영역 포시
        resultString = resultString & "# Note" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        resultString = resultString & "# Email" & vbCrLf
                
        ' html 가져오기
        htmlContent = ReadFileContent(htmlpath)
        processedHtml = htmlContent
        
        
        ' 표시 옵션에 따른 분개 처리
        optionValue = 2 ' 1:html그대로, 2:markdown 프로세싱
        Select Case optionValue
            Case 1
                ' html 그대로 출력
                
            Case 2
                ' 주석,태그 정리, table 변환
                processedHtml = ConvertHTMLToMarkdown(processedHtml, temporarySubjectLineString)
                
                ' ## 헤더로 회신메일 구분
                processedHtml = InsertReplyHeading(processedHtml)
                
                ' 하단 이미지, 표 이상 개선
                
            Case Else
                ' html 그대로 출력
        End Select
        
        ' (8) 메일 표시
        resultString = resultString & processedHtml
        
        ' ------------------------------------------------- Body End ----------------------------------------------------------------

        
        ' ===============================================================
        
        ' Make a dummy email to hold the details we're saving
        ' This way we dont get junk in the message header when saving
        Dim outputItem As mailItem
        Set outputItem = Application.CreateItem(olMailItem)
        outputItem.body = resultString

        fileName = mName & ".md"

        Debug.Print ObsidianFolder

        ' Save the result
        SaveAsUTF8 ObsidianFolder & fileName, resultString

        ' Fully encode the file path for Obsidian URI
        Dim obsidianURI As String
        obsidianURI = "obsidian://open?path=" & UrlEncodeUtf8NoBom(ObsidianFolder & fileName)

        ' Use ShellExecute to open the note in Obsidian
        ShellExecute 0, "open", obsidianURI, vbNullString, vbNullString, 1


    Next
EndClean:
    Set obj = Nothing
    Set oMail = Nothing
    Set outputItem = Nothing
End Sub


' ==================================================================================================
Sub ExtractEmail_html()

    ' Dim ObsidianFolder As String
    ' Dim baseFolder As String
    Dim vaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String
    Dim optionValue As Integer
    Dim mailItem As Outlook.mailItem
    Dim htmlContent As String
    Dim processedHtml As String
        
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink

    '================================================'
    ' Save as plain text
    Const OLTXT = 0
    ' Object holding variable
    Dim obj As Object
    ' Instantiate an Outlook Email Object
    Dim oMail As Outlook.mailItem
    ' If something breaks, skip to the end, tidy up and shut the door
    On Error GoTo EndClean:
    
    ' Establish the environment and selected items (emails)
    ' NOTE: selecting a conversation-view stack wont work
    '       you'll need to select one of the emails
    Dim fileName As String, mName As String
    Dim temporarySubjectLineString As String
    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection
    ' For each email in the Selection
    ' Assigning email item to the `obj` holding variable
    For Each obj In Selection
        ' set the oMail object equal to that mail item
        Set oMail = obj
        ' Is it an Email?
        If oMail.Class <> 43 Then
          MsgBox "This code only works with Emails."
          GoTo EndClean: ' you broke it
        End If

        ' Yank the mail items subject line to `temporarySubjectLineString`
        temporarySubjectLineString = oMail.subject
        ' function call the name cleaner to remove any
        '    illegal characters from the subject line
        ReplaceCharsForFileName temporarySubjectLineString, ""
        ' Yank the received date-time to a holding variable

        ' Build Recipient string based on receipient collection
        Dim recips As Outlook.Recipients
            Set recips = oMail.Recipients
        Dim recip As Outlook.Recipient
        Dim result As String
        Dim recipString As String
            recipString = ""

        For Each recip In recips
            recipString = recipString & vbTab
            recipString = recipString & "- "
            recipString = recipString & formatName(recip.Name, personNameStartChar)
            recipString = recipString & vbCrLf
        Next
        
        ' ===============================================================
        
        ' (1) 아웃룩 파일 html으로 저장
        Dim objItem As mailItem, htmlpath As String
        Set objItem = Application.ActiveExplorer.Selection(1)
        htmlpath = vaultPathToSaveFileTo & temporarySubjectLineString & ".html"
        ' Debug.Print htmlpath
        objItem.SaveAs htmlpath, 5
        
        
        ' (2) 첨부파일 저장
        Dim attachments As Outlook.attachments
        Dim Attachment As Outlook.Attachment
        Set attachments = objItem.attachments
        Dim i As Long
        For i = 1 To attachments.Count
            Set Attachment = attachments(i)
            Dim attachFileName As String
            attachFileName = vaultPathToSaveFileTo & temporarySubjectLineString & ".files\" & Attachment.fileName
            Attachment.SaveAsFile attachFileName
        Next i
        
        
        ' Build the result file content to be sent to the mail item body
        ' Then save that mail item same as the meeting extractor
        Dim sender As String
            sender = formatName(oMail.sender, personNameStartChar)
        Dim dtDate As Date
            dtDate = oMail.ReceivedTime
        Dim resultString As String

        ' resultString = ""
        ' resultString = resultString & "# [[" & emailFileNameStartChr & Format(oMail.ReceivedTime, "yyyy-mm-dd hhnn") & " " & temporarySubjectLineString & "|" & temporarySubjectLineString & "]]"
        ' resultString = resultString & vbCrLf & vbCrLf & vbCrLf
        
        ' (3) 프로퍼티 표시(YAML frontmatter)
        ' ----------------------------------------------- Properties star -----------------------------------------------------------
        resultString = "---" & vbCrLf
        
        ' Add tags
        resultString = resultString & "tags:" & vbCrLf
        
        ' Add classification and optional properties
        resultString = resultString & "tags:  " & """SOURCE/MAIL today""" & vbCrLf
        resultString = resultString & "Index: " & vbCrLf
        resultString = resultString & "title: """ & temporarySubjectLineString & """" & vbCrLf
        resultString = resultString & "aliases:" & vbCrLf
        resultString = resultString & "create: " & Format(Now, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "수신일시: " & Format(oMail.SentOn, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "요청일자: " & Format(oMail.SentOn, "yyyy-MM-dd") & vbCrLf
        resultString = resultString & "요청자: """ & sender & """" & vbCrLf
        resultString = resultString & "진행상태: " & "대기" & vbCrLf
        resultString = resultString & "D-day: """"" & vbCrLf
        resultString = resultString & "완료일: """"" & vbCrLf
        resultString = resultString & "ITSM: """"" & vbCrLf
        resultString = resultString & "ITSM_URL: """"" & vbCrLf
        ' Convert recipients to YAML list
        ' resultString = resultString & "to:" & vbCrLf
        ' For Each recip In recips
        '     resultString = resultString & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
        ' Next

        ' End YAML block
        resultString = resultString & "---" & vbCrLf
        ' ----------------------------------------------- Properties end ------------------------------------------------------------
        
        ' Now we create the file name
        ' mName = emailFileNameStartChr
        mName = mName & Format(dtDate, "yyyymmdd", vbUseSystemDayOfWeek, vbUseSystem)
        ' mName = mName & Format(dtDate, " hhMM", vbUseSystemDayOfWeek, vbUseSystem)
        mName = mName & " " & temporarySubjectLineString
 
 
        ' (4) 첨부파일 표시
        Dim propertyAccessor As Outlook.propertyAccessor
        Dim propHidden As String
        Dim isHidden As Variant
        propHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
        For i = 1 To attachments.Count
            Dim fNm As String
            Dim fLnk As String
            Dim fLst As String
       
            Set Attachment = attachments(i)
            Set propertyAccessor = Attachment.propertyAccessor
            isHidden = propertyAccessor.GetProperty(propHidden)
       
            ' PR_ATTACHMENT_HIDDEN 속성이 True인 경우에만 보여줌
            If Not isHidden Then
            fNm = Attachment.fileName
                fLnk = Replace("![[" & baseFolder & temporarySubjectLineString & ".files\" & fNm & "]]", "\", "/")
                resultString = resultString & fLnk & vbCrLf
            End If
        Next i
        
        
        ' Add a horizontal rule for separation
        resultString = resultString & vbCrLf & "---" & vbCrLf
        
        
        ' (5) 메일 본문 표시
        ' ------------------------------------------------- Body Start --------------------------------------------------------------
        ' (6) 소스 표시
        resultString = resultString & "![[" & temporarySubjectLineString & ".html]]" & vbCrLf
        resultString = resultString & "[[" & mName & "]]" & vbCrLf
        
        ' (7) 영역 포시
        resultString = resultString & "# Note" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        resultString = resultString & "# Email" & vbCrLf
                
        ' html 가져오기
        htmlContent = ReadFileContent(htmlpath)
        processedHtml = htmlContent
        
        
        ' 표시 옵션에 따른 분개 처리
        optionValue = 1 ' 1:html그대로, 2:markdown 프로세싱
        Select Case optionValue
            Case 1
                ' html 그대로 출력
                
            Case 2
                ' 주석,태그 정리, table 변환
                processedHtml = ConvertHTMLToMarkdown(processedHtml, temporarySubjectLineString)
                
                ' ## 헤더로 회신메일 구분
                processedHtml = InsertReplyHeading(processedHtml)
                
                ' 하단 이미지, 표 이상 개선
                
            Case Else
                ' html 그대로 출력
        End Select
        
        ' (8) 메일 표시
        resultString = resultString & processedHtml
        
        ' ------------------------------------------------- Body End ----------------------------------------------------------------

        
        ' ===============================================================
        
        ' Make a dummy email to hold the details we're saving
        ' This way we dont get junk in the message header when saving
        Dim outputItem As mailItem
        Set outputItem = Application.CreateItem(olMailItem)
        outputItem.body = resultString

        fileName = mName & ".md"

        Debug.Print ObsidianFolder

        ' Save the result
        SaveAsUTF8 ObsidianFolder & fileName, resultString

        ' Fully encode the file path for Obsidian URI
        Dim obsidianURI As String
        obsidianURI = "obsidian://open?path=" & UrlEncodeUtf8NoBom(ObsidianFolder & fileName)

        ' Use ShellExecute to open the note in Obsidian
        ShellExecute 0, "open", obsidianURI, vbNullString, vbNullString, 1


    Next
EndClean:
    Set obj = Nothing
    Set oMail = Nothing
    Set outputItem = Nothing
End Sub

