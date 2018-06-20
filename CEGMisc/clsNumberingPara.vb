Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports NCalc
Imports CEGINI
Public Class clsNumPara
    Public pNum As String
    Public pModNum As String
    Public pContent As String
    Public pError As String
End Class
Public Class clsNumberingPara
    Public dictCheckParaNum As Dictionary(Of String, List(Of clsNumPara))
    Public lstClass As List(Of clsNumPara)
    Public oWrdApp As Word.Application
    Public dChpNumber As String
    Public dChpTitle As String
    Public bPresedingZero As Boolean
    Private Shared _romanMap As New Dictionary(Of Char, Integer)() From {{"I", 1}, {"V", 5}, {"X", 10}, {"L", 50}, {"C", 100}, {"D", 500}, {"M", 1000}}
    Public Function OUPNumberingParaFormatQC(oWordApp As Word.Application, wListFiles As String, wFilePath As String)
        Dim miscINI As String
        Dim wDoc As Word.Document
        Dim strHtml As String
        Dim ranDoc As Word.Range
        Try
            oWrdApp = oWordApp
            miscINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New CEGINI.clsINI(miscINI)
            Dim nStyleName = oReadINI.IniReadValue("NumberingParaQc", "Style")
            dictCheckParaNum = New Dictionary(Of String, List(Of clsNumPara))
            If nStyleName <> String.Empty Then
                For Each wordFileName In wListFiles.Split("||")
                    If wordFileName <> String.Empty Then
                        oWrdApp.Documents(Path.Combine(wFilePath, wordFileName)).Activate() : wDoc = oWrdApp.ActiveDocument
                        ''''''''''''''''''''''''''
                        dChpNumber = "" : dChpTitle = ""
                        ranDoc = wDoc.Content
                        ranDoc.Find.ClearFormatting()
                        ranDoc.Find.Text = ""
                        ranDoc.Find.Style = oReadINI.IniReadValue("NumberingParaQc", "NumStyle")
                        If ranDoc.Find.Execute Then
                            dChpNumber = ranDoc.Text
                        End If
                        ranDoc.Find.ClearFormatting()
                        ranDoc = wDoc.Content
                        ranDoc.Find.Text = ""
                        ranDoc.Find.Style = oReadINI.IniReadValue("NumberingParaQc", "TitleStyle")
                        If ranDoc.Find.Execute Then
                            dChpTitle = ranDoc.Text
                        End If
                        ''''''''''''''''''''''''''
                        CollectNumberedParaContent(wDoc, nStyleName)
                        CheckOutofSequenceNumber()
                        dictCheckParaNum.Add(wordFileName & "|" & dChpNumber & "|" & dChpTitle, lstClass)
                    End If
                Next
                strHtml = GetLogFileInfo()
                If strHtml <> String.Empty Then
                    modDeclaration.WriteHtmlFile(Path.Combine(wFilePath, "NumberingParaLog_NP.html"), strHtml)
                End If

                For Each xVrnt In Split(wListFiles, "||")
                    If xVrnt <> String.Empty Then
                        oWrdApp.Documents(xVrnt).Activate()
                        wDoc = oWordApp.ActiveDocument
                        AddQCIteminCollection("NumberingParaQc", wDoc)
                    End If
                Next

            Else
                MessageBox.Show("Style not defined in ini : NumberingParaQc", sMsgTitle)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle)
        End Try
    End Function
    Public Function CheckOutofSequenceNumber()
        Dim CurNumber As String
        Dim PreNumber As String
        Try
            For Each cPara As clsNumPara In lstClass
                If cPara.pError = Nothing Then
                    CurNumber = cPara.pModNum
                    If PreNumber <> String.Empty Then
                        If CompareNumberSequence(CurNumber, PreNumber) = False Then
                            cPara.pError = "Please check the numbering sequence"
                        End If
                    End If
                    PreNumber = CurNumber
                End If
            Next
        Catch ex As Exception

        End Try
    End Function
    Public Function CompareNumberSequence(curNum As String, preNum As String) As Boolean
        Dim notSeq As Boolean
        Dim i As Integer
        Try
            Dim curArr() = curNum.Split(".")
            Dim PreArr() = preNum.Split(".")
            Dim incVal As String
            Dim decVal As String
            For i = UBound(PreArr) To 1 Step -1
                incVal = IncrementValue(PreArr, i)
                decVal = DecrementValue(PreArr, i)
                If incVal = curNum Then
                    CompareNumberSequence = True
                    Exit Function
                ElseIf decVal = curNum Then
                    CompareNumberSequence = True
                    Exit Function
                ElseIf Val(PreArr(LBound(PreArr)) + 1) & ".1" = curNum Then
                    CompareNumberSequence = True
                    Exit Function
                ElseIf Val(PreArr(LBound(PreArr)) + 1) & ".0" = curNum Then
                    CompareNumberSequence = True
                    Exit Function
                End If
            Next


            'If Val(curArr(UBound(curArr))) = Val(PreArr(UBound(PreArr)) + 1) Then
            '    If Val(curArr(LBound(curArr))) <> Val(PreArr(LBound(PreArr))) Then
            '        notSeq = True
            '    End If
            'ElseIf Val(curArr(UBound(curArr))) <> Val(PreArr(UBound(PreArr)) + 1) Then
            '    If Val(curArr(LBound(curArr))) < Val(PreArr(LBound(PreArr))) Then
            '        notSeq = True
            '    End If
            'End If
        Catch ex As Exception

        End Try
        CompareNumberSequence = notSeq
    End Function
    Public Function DecrementValue(curArr() As String, cPos As Integer) As String
        Dim sTemp As String
        Dim i As Integer
        Try
            For i = LBound(curArr) To UBound(curArr)
                If UBound(curArr) <> LBound(curArr) Then
                    If cPos <> i Then
                        If i = cPos - 1 Then
                            sTemp = sTemp & Val(curArr(i) + 1) & "."
                        Else
                            sTemp = sTemp & curArr(i) & "."
                        End If
                    End If
                End If
            Next
            sTemp = Regex.Replace(sTemp, "[\.]$", "")
        Catch ex As Exception

        End Try
        DecrementValue = sTemp
    End Function
    Public Function IncrementValue(curArr() As String, cPos As Integer) As String
        Dim sTemp As String
        Dim i As Integer
        Try
            For i = LBound(curArr) To UBound(curArr)
                If cPos = UBound(curArr) Then
                    If i = cPos Then
                        sTemp = sTemp & Val(curArr(i) + 1) & "."
                    Else
                        sTemp = sTemp & curArr(i) & "."
                    End If

                ElseIf cPos = LBound(curArr) Then
                    If i = cPos Then
                        sTemp = sTemp & Val(curArr(i) + 1) & "."
                    Else
                        sTemp = sTemp & "1."
                    End If
                Else
                    If i = cPos Then
                        sTemp = sTemp & Val(curArr(i) + 1) & "."
                    ElseIf i > cPos Then
                        sTemp = sTemp & "1."
                    Else
                        sTemp = sTemp & curArr(i) & "."
                    End If
                End If
            Next
            sTemp = Regex.Replace(sTemp, "[\.]$", "")
        Catch ex As Exception

        End Try
        IncrementValue = sTemp
    End Function
    Public Function CollectNumberedParaContent(wDoc As Word.Document, nStyleName As String)

        Dim wordFileName As String
        lstClass = New List(Of clsNumPara)
        Dim ranDoc As Word.Range
        'Dim regEx As New Regex("^\s*([a-zA-Z0-9.]+)")
        Dim regEx As New Regex("(^\s*)([A-z]+\.?)*(?<NSpace>\s*)(?<NPnum>([0-9]+\.?)*)")

        Try
            Dim I As Integer
            For I = 1 To 3
                Select Case (I)
                    Case 1
                        ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                    Case 2
                        If wDoc.Footnotes.Count > 0 Then
                            ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory)
                        Else
                            ranDoc = Nothing
                        End If
                    Case 3
                        If wDoc.Endnotes.Count > 0 Then
                            ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory)
                        Else
                            ranDoc = Nothing
                        End If
                End Select
                If Not ranDoc Is Nothing Then
                    Dim ranText As String
                    'If ranDoc.Text.Contains("22") Then
                    '    MsgBox("found")
                    'End If
                    bPresedingZero = False
                    ranDoc.Find.Text = String.Empty
                    ranDoc.Find.Style = nStyleName
                    Do While ranDoc.Find.Execute
                        ranText = ranDoc.Text
                        Dim objNumPara As New clsNumPara
                        If regEx.IsMatch(ranText) Then
                            objNumPara.pNum = regEx.Match(ranText).ToString()
                            If regEx.Match(ranText).Groups("NSpace").ToString() <> String.Empty And regEx.Match(ranText).Groups("NPnum").ToString().Replace(".", "") <> String.Empty Then
                                objNumPara.pError = "Please check the space"
                            End If
                            If objNumPara.pNum <> objNumPara.pNum.ToUpper Then
                                objNumPara.pError = "Numbering content should be uppercase."
                            End If
                            Dim sTemp
                            sTemp = ChangeTexttoNumber(objNumPara.pNum)
                            If bPresedingZero = True Then 'Srilakshmi request as per client
                                objNumPara.pError = "Please check preseding zero"
                                bPresedingZero = False
                            End If
                            If sTemp <> Nothing Then
                                objNumPara.pModNum = sTemp
                            Else
                                objNumPara.pModNum = "0"
                                objNumPara.pError = "Numbering content not found"
                            End If
                        Else
                            objNumPara.pNum = ""
                            objNumPara.pError = "Numbering content not found"
                            objNumPara.pModNum = "0"
                        End If

                        objNumPara.pContent = ranDoc.Text
                        lstClass.Add(objNumPara)
                        ranDoc = wDoc.Range(ranDoc.End + 1, wDoc.Range.End)
                        ranDoc.Find.Text = String.Empty
                        ranDoc.Find.Style = nStyleName
                        bPresedingZero = False
                    Loop
                End If
            Next 'I

        Catch ex As Exception

        End Try
    End Function
    Public Function ChangeTexttoNumber(strNumber As String) As String
        Dim rText As String
        Try
            strNumber = Regex.Replace(strNumber, "[\.]$", "")
            For Each eDot As String In strNumber.Split(".")
                Dim n As Integer
                If eDot.Length > 2 Then
                    If eDot(0).ToString() = "0" Then
                        bPresedingZero = True
                    End If
                End If
                Dim isNumeric As Boolean = Integer.TryParse(eDot, n)
                If Not isNumeric Then
                    If Regex.IsMatch(eDot, "^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$") Then
                        eDot = ConvertRomanToNumber(eDot).ToString()
                    Else
                        Dim nNum As String
                        For Each eChar As Char In eDot
                            If Char.IsDigit(eChar) Then
                                nNum = nNum & Char.GetNumericValue(eChar) & "."
                            Else
                                'nNum = nNum & ConvertChartoNumber(eChar) & "."
                                nNum = nNum & Asc(eChar) & "."
                            End If
                        Next
                        nNum = Regex.Replace(nNum, "[\.]$", "")
                        eDot = nNum
                    End If
                Else
                    eDot = Trim(Val(eDot))
                End If
                rText = rText & Trim(eDot) & "."
            Next
            rText = Regex.Replace(rText, "[\.]$", "")
        Catch ex As Exception

        End Try

        ChangeTexttoNumber = rText
    End Function
    Public Shared Function ConvertRomanToNumber(text As String) As Integer
        Dim totalValue As Integer = 0, prevValue As Integer = 0
        For Each c In text
            If Not _romanMap.ContainsKey(c) Then
                Return 0
            End If
            Dim crtValue = _romanMap(c)
            totalValue += crtValue
            If prevValue <> 0 AndAlso prevValue < crtValue Then
                If prevValue = 1 AndAlso (crtValue = 5 OrElse crtValue = 10) OrElse prevValue = 10 AndAlso (crtValue = 50 OrElse crtValue = 100) OrElse prevValue = 100 AndAlso (crtValue = 500 OrElse crtValue = 1000) Then
                    totalValue -= 2 * prevValue
                Else
                    Return 0
                End If
            End If
            prevValue = crtValue
        Next
        Return totalValue
    End Function
    Public Function GetLogFileInfo() As String
        Dim htmContent As String
        Dim HtmlRoot As String
        HtmlRoot = "<HTML><head><META http-equiv='Content-Type' content='text/html; charset=utf-8'>" & _
                   "</head><body style='font-family:Times New Roman'><H4 align=""center"" style='background-color:FF99CC;font-family:Verdana'>CE Genius Numbering Para Report</H4><body bgcolor=""#FFFFFF"" style=""font-family:Verdana""><table border=""0"" align=""center""><tbody>" + _
                   "<tr><td><b>Date and time</b></td><td><b>: " + DateTime.Now + "</b></td></tr><tr>" & _
                   "<td><b> User name</b></td><td><b>: " + Environment.UserName + "</b></td></tr><tr>" & _
                   "</tbody></table><hr color='#FF8C00'/>"

        Try
            Dim fileName As String
            dChpNumber = "" : dChpTitle = "" : lstClass = Nothing
            For Each pair As KeyValuePair(Of String, List(Of clsNumPara)) In dictCheckParaNum
                fileName = pair.Key.Split("|")(0)
                dChpNumber = pair.Key.Split("|")(1)
                dChpTitle = pair.Key.Split("|")(2)
                htmContent = htmContent & "<h2 style='background-color:#00CFFF'>File Name: " + fileName + "</h2>"
                If dChpTitle <> String.Empty And dChpNumber <> String.Empty Then
                    htmContent = htmContent & "<h4 style='background-color:#87CEEB'>" + dChpNumber + " : " + dChpTitle + "</h4><br/>"
                ElseIf dChpTitle <> String.Empty Then
                    htmContent = htmContent & "<h4 style='background-color:#87CEEB'>" + dChpTitle + "</h4><br/>"
                End If
                lstClass = pair.Value
                htmContent = htmContent & "<table style='border-spacing: 20px 5px;'>"

                If lstClass.Count = 0 Then
                    htmContent = htmContent & "<tr><td style='color:#D2691E'>None</td><td></td><td></td></tr>"
                Else
                    For Each cPara As clsNumPara In lstClass
                        'htmContent = htmContent & "<tr><td>" & cPara.pNum & "</td><tr>"
                        If cPara.pError <> Nothing Then
                            htmContent = htmContent & "<tr><td style='background-color:#DC143C;'>" & cPara.pNum & "</td><td style='background-color:#87CEEB;'>" & cPara.pError & "</td><td>" & cPara.pContent & "</td><tr>"
                        Else
                            htmContent = htmContent & "<tr><td>" & cPara.pNum & "</td><td></td><td></td></tr>"
                        End If
                    Next
                End If

                htmContent = htmContent & "</table><hr color='#FF8C00'/>"
            Next
            htmContent = HtmlRoot & htmContent & "</body></html>"

            'Dim categories = lstClass
            'categories.Sort(New CategoryComparer())
            'For Each category In categories
            '    'htmContent = htmContent & "<p>" & category.pModNum & "</p>"
            'Next

        Catch ex As Exception

        End Try

        GetLogFileInfo = htmContent
    End Function
End Class
