Attribute VB_Name = "SaveUtilities"
Option Explicit
'======================================================================================='
Public Function GetCurrentItem() As Object
    ' Instantiate an Outlook application instance
    Dim objApp As Outlook.Application
        Set objApp = Application

    ' Don't nuke the process if something breaks
    On Error Resume Next

    ' Depending on Which type of active window is active in Outlook
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            ' If explorer than grab the current active selection
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.item(1)
        Case "Inspector"
            ' If Inspector than grab the current item
            Set GetCurrentItem = objApp.ActiveInspector.CurrentItem
    End Select
    ' Tidy up and de-allocate the Outlook instance
    Set objApp = Nothing
End Function

Function URLEncode(str As String) As String
    Dim i As Integer
    Dim c As String
    Dim result As String
    result = ""

    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        Select Case Asc(c)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                ' Keep alphanumeric characters and safe symbols (-, ., _, ~)
                result = result & c
            Case 32
                ' Encode space as %20
                result = result & "%20"
            Case Else
                ' Encode all other characters
                result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i

    URLEncode = result
End Function

Public Function UrlEncodeUtf8NoBom(ByVal sText As String) As String
    Dim oStream As Object     ' ADODB.Stream
    Dim byteArray() As Byte
    Dim i As Long
    Dim sEncoded As String
    Dim startIndex As Long
    
    ' --- Step 1: Write text as UTF-8 into a stream ---
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2            ' adTypeText
    oStream.Mode = 3            ' adModeReadWrite
    oStream.Charset = "UTF-8"
    oStream.Open
    oStream.WriteText sText
    oStream.Position = 0
    oStream.Type = 1            ' Switch to binary to read raw bytes
    byteArray = oStream.Read
    oStream.Close
    Set oStream = Nothing
    
    ' --- Step 2: Detect & skip BOM if present (EF BB BF) ---
    startIndex = LBound(byteArray)
    If (UBound(byteArray) - LBound(byteArray) >= 2) Then
        If byteArray(0) = &HEF And byteArray(1) = &HBB And byteArray(2) = &HBF Then
            startIndex = 3 ' jump past the BOM
        End If
    End If
    
    ' --- Step 3: Percent-encode each remaining byte ---
    For i = startIndex To UBound(byteArray)
        sEncoded = sEncoded & "%" & Right("0" & Hex(byteArray(i)), 2)
    Next i
    
    UrlEncodeUtf8NoBom = sEncoded
End Function


'======================================================================================='
' STRING CLEANING SUBROUTINE
Public Sub ReplaceCharsForFileName(temporarySubjectLineString As String, sChr As String)
    ' This just cleans the Email subject line of invalid characters
    temporarySubjectLineString = Replace(temporarySubjectLineString, "/", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "\", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, ":", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "?", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, Chr(34), sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "<", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, ">", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "|", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "[", sChr)
    temporarySubjectLineString = Replace(temporarySubjectLineString, "]", sChr)
End Sub
'======================================================================================='
Public Function formatName(str As String, personNameStartChar As String) As String
    ' Meeting attendee names are formatted strangely
    ' This function parses the attendees and formats them to:
    ' [[@Bryan Jenks]]
    Dim typeOfNameToClean As Integer
    
    ' If attendee is an outside Active Directory individual
    ' like a gmail account or external person then the display
    ' is just first and last names. these are perfect to easily format
    ' ex: `Bryan Jenks`
    Dim regexJustFirstNameAndLastName As Object
        Set regexJustFirstNameAndLastName = New RegExp
        regexJustFirstNameAndLastName.Pattern = "^\w+\s\w+$"
    If regexJustFirstNameAndLastName.Test(str) = True Then typeOfNameToClean = 1
    Set regexJustFirstNameAndLastName = Nothing
    
    ' This finds emails that are `last name, first name@domain.com`
    ' not just `.com` will also pick up multiples like
    ' ex: `@domain.or.gov` etc
    Dim regexFirstNameLastNameAndFullDomain As Object
        Set regexFirstNameLastNameAndFullDomain = New RegExp
        regexFirstNameLastNameAndFullDomain.Pattern = "^\w+,\s\w+@\w+(\.\w+)+"
    If regexFirstNameLastNameAndFullDomain.Test(str) = True Then typeOfNameToClean = 2
    Set regexFirstNameLastNameAndFullDomain = Nothing
    
    ' A full and normal email only as the invited person
    ' ex: johndoe@domain.com
    Dim regexPlainEmailAddress As Object
        Set regexPlainEmailAddress = New RegExp
        regexPlainEmailAddress.Pattern = "^\w+@\w+\.\w+"
    If regexPlainEmailAddress.Test(str) = True Then typeOfNameToClean = 3
    Set regexPlainEmailAddress = Nothing
    
    ' Active Directory people that display like:
    'ex: `Last Name, First Name (AGENCY)`
    Dim regexLastNameFirstNameAndAgency As Object
        Set regexLastNameFirstNameAndAgency = New RegExp
        regexLastNameFirstNameAndAgency.Pattern = "^[a-zA-Z_\-]+,\s[a-zA-Z_\-]+\s\(\w+\)$"
    If regexLastNameFirstNameAndAgency.Test(str) = True Then typeOfNameToClean = 4
    Set regexLastNameFirstNameAndAgency = Nothing
    
    ' Single name entity such as a distribution list or 1 word titled entity
    ' ex: `payroll (Agency)`
    Dim regexSingleNameAndDomain As Object
        Set regexSingleNameAndDomain = New RegExp
        regexSingleNameAndDomain.Pattern = "^([a-zA-Z_\-]+\s)+\([a-zA-Z_\-]+\)"
    If regexSingleNameAndDomain.Test(str) = True Then typeOfNameToClean = 5
    Set regexSingleNameAndDomain = Nothing
    
    Select Case typeOfNameToClean
        Case 1 ' John Doe
            formatName = "[[" & personNameStartChar & str & "]]"
        Case 2 ' Doe, John@domain.or.gov
            Dim fName As String, lname As String
            fName = Mid(str, InStr(str, ", ") + 2, InStr(str, "@") - (InStr(str, ", ") + 2))
            lname = Mid(str, 1, InStr(str, ",") - 1)
            ' Assemble the building blocks and assigning to return value
            formatName = "[[" & personNameStartChar & fName & " " & lname & "]]"
        Case 3 ' JohnDoe@gmail.com
            formatName = "[[" & personNameStartChar & Left(str, InStr(str, "@") - 1) & "]]"
        Case 4 ' Doe, John (Agency)
            Dim fname1 As String, lname1 As String
            fname1 = Mid(str, InStr(str, ", ") + 2, InStr(str, " (") - (InStr(str, ", ") + 2))
            lname1 = Mid(str, 1, InStr(str, ",") - 1)
            ' Assemble the building blocks and assigning to return value
            formatName = "[[" & personNameStartChar & fname1 & " " & lname1 & "]]"
        Case 5 ' Payroll (Agency)
            formatName = "[[" & Left(str, InStr(str, " (") - 1) & "]]"
        Case Else ' Anything else
            formatName = "[[" & str & "]]"
    End Select

End Function

'======================================================================================='
Public Sub SaveAsUTF8(filePath As String, content As String)
    ' Late-binding ADODB, no extra references needed
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    ' We are writing text
    stm.Type = 2 'adTypeText
    ' Set text mode read/write
    stm.Mode = 3 'adModeReadWrite
    ' Use UTF-8 to preserve accents and special chars
    stm.Charset = "UTF-8"
    
    ' Open the stream, write content, save to file
    stm.Open
    stm.WriteText content
    stm.SaveToFile filePath, 2 ' adSaveCreateOverWrite = 2
    stm.Close
    
    Set stm = Nothing

End Sub

'============================================
' HTML ������ �о� ���ڿ��� ��ȯ
Public Function ReadFileContent(ByVal filePath As String) As String
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        ReadFileContent = ""
        Exit Function
    End If
    Set f = fso.OpenTextFile(filePath, 1) ' ForReading = 1
    ReadFileContent = f.ReadAll
    f.Close
End Function


'============================================
' HTML �� <img> �±׸� Obsidian �̹��� ��ũ�� ��ȯ (��δ� baseName.files\ ���)
Function ReplaceImageTagsWithMarkdownLinks(ByVal html As String, ByVal baseName As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<img[^>]*src\s*=\s*[""']([^""']+)[""'][^>]*>"
    
    Dim matches As Object, m As Object
    Dim result As String, newSrc As String, imgFileName As String
    result = html
    Set matches = regEx.Execute(html)
    Dim i As Long
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        newSrc = m.SubMatches(0)
        imgFileName = GetFileNameFromPath(newSrc)
        newSrc = baseName & ".files\" & imgFileName
        Dim mdLink As String
        mdLink = Replace("![[" & baseFolder & newSrc & "]]", "\", "/")
        result = Left(result, m.FirstIndex) & mdLink & Mid(result, m.FirstIndex + m.Length + 1)
    Next i
    ReplaceImageTagsWithMarkdownLinks = result
End Function

'============================================
' ��ο��� ���ϸ� ���� (��: "folder/image001.png" �� "image001.png")
Function GetFileNameFromPath(ByVal path As String) As String
    Dim pos As Long
    pos = InStrRev(path, "/")
    If pos = 0 Then pos = InStrRev(path, "\")
    If pos > 0 Then
        GetFileNameFromPath = Mid(path, pos + 1)
    Else
        GetFileNameFromPath = path
    End If
End Function


'============================================
' HTML �� Markdown ��ȯ (ǥ�� �̹����� ���� ó��)
Public Function ConvertHTMLToMarkdown(ByVal html As String, ByVal baseName As String) As String
    Dim processedHtml As String
    ' 1. ���ʿ��� �±� ���� (���Ǻ� �ּ�, VML ��)
    processedHtml = clearHTML(html)
    
    ' 2. �̹��� �±� ��ȯ
    processedHtml = ReplaceImageTagsWithMarkdownLinks(processedHtml, baseName)
    
    ' 3. ���̺� ó��: HTML ���� <table> �±׺��� �и��ؼ� mdǥ�� ��ȯ�� ��,
    '    �ش� �κ��� placeholder�� ġȯ�ϰ� ���߿� �ٽ� ����
    Dim tableRegEx As Object, tableMatches As Object, match As Object
    Dim placeholders() As String, tableMarkdown() As String
    Dim idx As Long
    idx = 0
    Set tableRegEx = CreateObject("VBScript.RegExp")
    tableRegEx.Pattern = "<table[\s\S]*?</table>"
    tableRegEx.Global = True
    tableRegEx.IgnoreCase = True
    Set tableMatches = tableRegEx.Execute(processedHtml)
    
    For Each match In tableMatches
        ReDim Preserve placeholders(idx)
        ReDim Preserve tableMarkdown(idx)
        placeholders(idx) = "%%TABLE_" & idx & "%%"
        tableMarkdown(idx) = ConvertHTMLTableToMarkdown(match.Value)
        processedHtml = Replace(processedHtml, match.Value, placeholders(idx))
        idx = idx + 1
    Next match
    
    ' 4. ���� HTML �±� ����
    processedHtml = StripHtmlTags(processedHtml)
    
    ' 5. placeholder���� mdǥ�� ġȯ
    Dim i As Long
    For i = 0 To idx - 1
        processedHtml = Replace(processedHtml, placeholders(i), vbCrLf & tableMarkdown(i) & vbCrLf)
    Next i
    
    ConvertHTMLToMarkdown = processedHtml
End Function

'============================================
' html ���� ����(��ũ�ٿ� ��ȯ)
Public Function clearHTML(ByVal html As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' 1. ���Ǻ� �ּ� ����: <!--[if ...]> ... <![endif]-->
    regEx.Pattern = "<!--\[if[\s\S]*?<!\[endif\]-->"
    html = regEx.Replace(html, "")
    
    ' 2. VML �±� ����: <v:...> �±�
    regEx.Pattern = "</?v:[^>]+>"
    html = regEx.Replace(html, "")
    
    ' 3. &nbsp; ���� (�������� ġȯ)
    regEx.Pattern = "&nbsp;"
    html = regEx.Replace(html, " ")
    
    ' 4. �±� ���� �����ݷ�(;) �����
    Dim tagMatches As Object, tagMatch As Object
    Dim originalTag As String, newTag As String
    regEx.Pattern = "<[^>]+>"
    Set tagMatches = regEx.Execute(html)
    For Each tagMatch In tagMatches
        originalTag = tagMatch.Value
        newTag = Replace(originalTag, ";", "")
        html = Replace(html, originalTag, newTag)
    Next tagMatch
    
    ' 5. <br> �±׸� �ٹٲ�(vbCrLf)���� ��ȯ
    regEx.Pattern = "<br\s*/?>"
    html = regEx.Replace(html, vbCrLf)
    
    ' 6. <p> �±�(����/�ݴ� �±�)�� �ٹٲ����� ��ȯ
    regEx.Pattern = "</?p\s*[^>]*>"
    html = regEx.Replace(html, vbCrLf)
    
    ' 7. ���ʿ��� ����(�ٹٲ�) ����
    ' 7-1. �� �� �̻��� ���ӵ� �ٹٲ��� �ܶ� ���п� ��Ŀ�� ����
    regEx.Pattern = "(\r\n){2,}"
    html = regEx.Replace(html, "<<<PARA>>>")
    ' 7-2. ���� ���� �ٹٲ��� �������� ��ȯ (��, �� �� �� ����)
    regEx.Pattern = "\r\n"
    html = regEx.Replace(html, " ")
    ' 7-3. �ܶ� ���� ��Ŀ�� �ٽ� �ٹٲ����� ����
    html = Replace(html, "<<<PARA>>>", vbCrLf)
    
    ' 8. �յ� ���� ����
    html = Trim(html)
    
    clearHTML = html
End Function

'============================================
' HTML ���̺��� Markdown ǥ�� ��ȯ
Public Function ConvertHTMLTableToMarkdown(tblHtml As String) As String
    On Error GoTo ErrorHandler
    Dim md As String
    Dim htmlDoc As Object, tbl As Object, rows As Object, r As Long, c As Long
    Dim cellText As String, colCount As Long
    
    Set htmlDoc = CreateObject("htmlfile")
    htmlDoc.Open
    htmlDoc.Write tblHtml
    htmlDoc.Close
    Set tbl = htmlDoc.getElementsByTagName("table")(0)
    Set rows = tbl.getElementsByTagName("tr")
    
    If rows.Length = 0 Then
        ConvertHTMLTableToMarkdown = ""
        Exit Function
    End If
    
    ' ù ���� ����� ��� (th �Ǵ� td)
    Dim headerCells As Object
    Set headerCells = rows(0).getElementsByTagName("th")
    If headerCells.Length = 0 Then
        Set headerCells = rows(0).getElementsByTagName("td")
    End If
    colCount = headerCells.Length
    If colCount = 0 Then
        ConvertHTMLTableToMarkdown = ""
        Exit Function
    End If
    
    md = "|"
    For c = 0 To colCount - 1
        cellText = CleanText(headerCells(c).innerText)
        md = md & " " & cellText & " |"
    Next c
    md = md & vbCrLf & "|"
    For c = 0 To colCount - 1
        md = md & " --- |"
    Next c
    md = md & vbCrLf
    
    Dim currentCells As Object
    For r = 1 To rows.Length - 1
        Set currentCells = rows(r).getElementsByTagName("td")
        If currentCells.Length = 0 Then
            Set currentCells = rows(r).getElementsByTagName("th")
        End If
        If currentCells.Length > 0 Then
            md = md & "|"
            For c = 0 To currentCells.Length - 1
                cellText = CleanText(currentCells(c).innerText)
                md = md & " " & cellText & " |"
            Next c
            md = md & vbCrLf
        End If
    Next r
    
    ConvertHTMLTableToMarkdown = md
    Exit Function
ErrorHandler:
    ConvertHTMLTableToMarkdown = ""
End Function

'============================================
' �� �ؽ�Ʈ���� �ٹٲ� ���� �� �յ� ���� ����
Public Function CleanText(ByVal text As String) As String
    CleanText = Trim(Replace(text, vbCrLf, " "))
End Function

'============================================
' ���̺� ��ȯ�� ���� ���ʿ��� HTML �±� ����
Function StripHtmlTags(ByVal html As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<[^>]+>"
    StripHtmlTags = regEx.Replace(html, "")
End Function


'============================================
' "���� ���:" ������ ��Ÿ����, ���� �ִ� 10�� ������ "����:"�� ã��
' "# ȸ�� ���� - [����]" �������� ��� ����
Function InsertReplyHeading(ByVal text As String) As String
    Dim lines() As String, i As Long, j As Long, subject As String
    Dim newLines As Collection
    Set newLines = New Collection
    
    ' �ؽ�Ʈ �ٹٲ� ó��
    If InStr(text, vbCrLf) > 0 Then
        lines = Split(text, vbCrLf)
    Else
        lines = Split(text, vbLf) ' �Ϻ� ȯ�濡�� vbLf�� ������ ���ɼ� ó��
    End If

    subject = "" ' �⺻�� �ʱ�ȭ

    ' ����� �޽��� ���
    ' Debug.Print "----- ���� -----"
    
    For i = 0 To UBound(lines)
        ' Debug.Print "���� ��(" & i & "): " & lines(i)

        ' "���� ���:" ã��
        If InStr(LCase(Trim(lines(i))), "���� ���:") > 0 Then
            ' Debug.Print "�� '���� ���:' �߰� (��: " & i & ")"

            ' "���� ���:" ���� �ִ� 10�� ������ "����:" ã��
            Dim maxSearchIndex As Long
            maxSearchIndex = IIf(i + 10 < UBound(lines), i + 10, UBound(lines))
            
            For j = i + 1 To maxSearchIndex
                ' Debug.Print "  - Ž�� ��(" & j & "): " & lines(j)
                
                If InStr(LCase(Trim(lines(j))), "����:") > 0 Then
                    subject = Trim(Mid(lines(j), InStr(lines(j), "����:") + 4)) ' "����:" ���� �� ����
                    ' Debug.Print "  ?? '����:' �߰� �� " & subject
                    Exit For ' ���� ã���� ����
                End If
            Next j

            ' "����:"�� ã�Ҵ��� Ȯ�� �� ��� ����
            If subject <> "" Then
                newLines.Add "## " & subject
            Else
                newLines.Add "## ȸ�� ����"
            End If
        End If

        ' ���� �� �߰�
        newLines.Add lines(i)
    Next i

    ' ��� ���ڿ� ��ȯ
    ' Debug.Print "----- �Ϸ� -----"
    InsertReplyHeading = Join(CollectionToArray(newLines), vbCrLf)
End Function

' Collection�� �迭�� ��ȯ�ϴ� �Լ�
Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String, i As Long
    ReDim arr(0 To col.Count - 1)
    
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    
    CollectionToArray = arr
End Function
