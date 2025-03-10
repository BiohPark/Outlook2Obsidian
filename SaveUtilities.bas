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
' HTML 파일을 읽어 문자열로 반환
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
' HTML 내 <img> 태그를 Obsidian 이미지 링크로 변환 (경로는 baseName.files\ 사용)
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
' 경로에서 파일명만 추출 (예: "folder/image001.png" → "image001.png")
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
' HTML → Markdown 변환 (표와 이미지를 별도 처리)
Public Function ConvertHTMLToMarkdown(ByVal html As String, ByVal baseName As String) As String
    Dim processedHtml As String
    ' 1. 불필요한 태그 제거 (조건부 주석, VML 등)
    processedHtml = clearHTML(html)
    
    ' 2. 이미지 태그 변환
    processedHtml = ReplaceImageTagsWithMarkdownLinks(processedHtml, baseName)
    
    ' 3. 테이블 처리: HTML 내의 <table> 태그별로 분리해서 md표로 변환한 후,
    '    해당 부분을 placeholder로 치환하고 나중에 다시 삽입
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
    
    ' 4. 남은 HTML 태그 제거
    processedHtml = StripHtmlTags(processedHtml)
    
    ' 5. placeholder들을 md표로 치환
    Dim i As Long
    For i = 0 To idx - 1
        processedHtml = Replace(processedHtml, placeholders(i), vbCrLf & tableMarkdown(i) & vbCrLf)
    Next i
    
    ConvertHTMLToMarkdown = processedHtml
End Function

'============================================
' html 사전 정리(마크다운 변환)
Public Function clearHTML(ByVal html As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' 1. 조건부 주석 제거: <!--[if ...]> ... <![endif]-->
    regEx.Pattern = "<!--\[if[\s\S]*?<!\[endif\]-->"
    html = regEx.Replace(html, "")
    
    ' 2. VML 태그 제거: <v:...> 태그
    regEx.Pattern = "</?v:[^>]+>"
    html = regEx.Replace(html, "")
    
    ' 3. &nbsp; 제거 (공백으로 치환)
    regEx.Pattern = "&nbsp;"
    html = regEx.Replace(html, " ")
    
    ' 4. 태그 안의 세미콜론(;) 지우기
    Dim tagMatches As Object, tagMatch As Object
    Dim originalTag As String, newTag As String
    regEx.Pattern = "<[^>]+>"
    Set tagMatches = regEx.Execute(html)
    For Each tagMatch In tagMatches
        originalTag = tagMatch.Value
        newTag = Replace(originalTag, ";", "")
        html = Replace(html, originalTag, newTag)
    Next tagMatch
    
    ' 5. <br> 태그를 줄바꿈(vbCrLf)으로 변환
    regEx.Pattern = "<br\s*/?>"
    html = regEx.Replace(html, vbCrLf)
    
    ' 6. <p> 태그(여는/닫는 태그)를 줄바꿈으로 변환
    regEx.Pattern = "</?p\s*[^>]*>"
    html = regEx.Replace(html, vbCrLf)
    
    ' 7. 불필요한 엔터(줄바꿈) 정리
    ' 7-1. 두 개 이상의 연속된 줄바꿈은 단락 구분용 마커로 변경
    regEx.Pattern = "(\r\n){2,}"
    html = regEx.Replace(html, "<<<PARA>>>")
    ' 7-2. 남은 단일 줄바꿈은 공백으로 변환 (즉, 한 줄 내 연결)
    regEx.Pattern = "\r\n"
    html = regEx.Replace(html, " ")
    ' 7-3. 단락 구분 마커를 다시 줄바꿈으로 복원
    html = Replace(html, "<<<PARA>>>", vbCrLf)
    
    ' 8. 앞뒤 공백 제거
    html = Trim(html)
    
    clearHTML = html
End Function

'============================================
' HTML 테이블을 Markdown 표로 변환
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
    
    ' 첫 행을 헤더로 사용 (th 또는 td)
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
' 셀 텍스트에서 줄바꿈 제거 및 앞뒤 공백 삭제
Public Function CleanText(ByVal text As String) As String
    CleanText = Trim(Replace(text, vbCrLf, " "))
End Function

'============================================
' 테이블 변환을 위한 불필요한 HTML 태그 제거
Function StripHtmlTags(ByVal html As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<[^>]+>"
    StripHtmlTags = regEx.Replace(html, "")
End Function


'============================================
' "보낸 사람:" 문구가 나타나면, 이후 최대 10줄 내에서 "제목:"을 찾아
' "# 회신 메일 - [제목]" 형식으로 헤더 삽입
Function InsertReplyHeading(ByVal text As String) As String
    Dim lines() As String, i As Long, j As Long, subject As String
    Dim newLines As Collection
    Set newLines = New Collection
    
    ' 텍스트 줄바꿈 처리
    If InStr(text, vbCrLf) > 0 Then
        lines = Split(text, vbCrLf)
    Else
        lines = Split(text, vbLf) ' 일부 환경에서 vbLf만 존재할 가능성 처리
    End If

    subject = "" ' 기본값 초기화

    ' 디버깅 메시지 출력
    ' Debug.Print "----- 시작 -----"
    
    For i = 0 To UBound(lines)
        ' Debug.Print "현재 줄(" & i & "): " & lines(i)

        ' "보낸 사람:" 찾기
        If InStr(LCase(Trim(lines(i))), "보낸 사람:") > 0 Then
            ' Debug.Print "▶ '보낸 사람:' 발견 (줄: " & i & ")"

            ' "보낸 사람:" 이후 최대 10줄 내에서 "제목:" 찾기
            Dim maxSearchIndex As Long
            maxSearchIndex = IIf(i + 10 < UBound(lines), i + 10, UBound(lines))
            
            For j = i + 1 To maxSearchIndex
                ' Debug.Print "  - 탐색 줄(" & j & "): " & lines(j)
                
                If InStr(LCase(Trim(lines(j))), "제목:") > 0 Then
                    subject = Trim(Mid(lines(j), InStr(lines(j), "제목:") + 4)) ' "제목:" 이후 값 추출
                    ' Debug.Print "  ?? '제목:' 발견 → " & subject
                    Exit For ' 제목 찾으면 종료
                End If
            Next j

            ' "제목:"을 찾았는지 확인 후 헤더 삽입
            If subject <> "" Then
                newLines.Add "## " & subject
            Else
                newLines.Add "## 회신 메일"
            End If
        End If

        ' 현재 줄 추가
        newLines.Add lines(i)
    Next i

    ' 결과 문자열 반환
    ' Debug.Print "----- 완료 -----"
    InsertReplyHeading = Join(CollectionToArray(newLines), vbCrLf)
End Function

' Collection을 배열로 변환하는 함수
Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String, i As Long
    ReDim arr(0 To col.Count - 1)
    
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    
    CollectionToArray = arr
End Function
