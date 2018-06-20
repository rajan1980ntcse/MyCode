Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports System.Text.RegularExpressions
Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Public Class clsRefStyleInfo
    Public sParaStyleName As String
    Public dCharStyleInfo As Dictionary(Of Word.Range, String)
    Public rContent As Word.Range
    Public sMessage As String
    Public sValid As Boolean
    Public sPattern As String
    Public sResultPattern As String
End Class
Module RefStyleModifier
    Public dictReferInfo As Dictionary(Of Word.Range, clsRefStyleInfo)
    Public wAppp As Word.Application
    Public Function ReferenceStyleModifierMain(wDoc As Word.Document, WAPP As Word.Application, StyleName As String)
        Try
            Dim pStyleStructureINI As String
            wAppp = WAPP
            pStyleStructureINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGRefStyleStructure.ini")
            If Not File.Exists(pStyleStructureINI) Then
                MessageBox.Show("Unable to process due to configuration file missing in CEG", "CE Genius")
            End If
            dictReferInfo = New Dictionary(Of Word.Range, clsRefStyleInfo)
            CollectReferenceInformation(wDoc, pStyleStructureINI, StyleName)
            If dictReferInfo.Count > 0 Then
                ConvertStyledReftoStructuredRef(wDoc, pStyleStructureINI, StyleName)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Public Function ConvertStyledReftoStructuredRef(aDoc As Word.Document, sConfigPath As String, sRefStyle As String)
        Try
            Dim ranDoc As Word.Range
            Dim dictResult = dictReferInfo.Where(Function(x) x.Value.sPattern <> Nothing And x.Value.sMessage = Nothing And x.Value.sValid = True).ToList
            ''[Author]([pubdateYear]). '[articleTitle]'. [journalTitle] [volume]([issueNumber]):[page].
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sSkipList() = oReadINI.INIReadValue("RefStyleStructure", "Skip_StyleList").Split("|")
            Dim tmpDoc As Word.Document = wAppp.Documents.Add(Visible:=True)
            tmpDoc.Activate()
            Dim oCmt As Word.Comment
            For Each dictVal In dictResult
                tmpDoc.Activate()
                Dim ranRef As Word.Range
                ranRef = tmpDoc.Content
                ''dictVal.Key.Select()
                tmpDoc.Range.Text = dictVal.Value.sPattern
                For Each pair As KeyValuePair(Of Word.Range, String) In dictVal.Value.dCharStyleInfo
                    If Not sSkipList.Contains(pair.Value) Then
                        ranRef = tmpDoc.Content
                        With ranRef.Find
                            .ClearFormatting() : .Replacement.ClearFormatting()
                            .Text = "[" & pair.Value.Replace("ref_", "") & "]"
                        End With
                        If ranRef.Find.Execute Then
                            ranRef.Select()
                            Dim ranSel As Word.Range = wAppp.Selection.Range
                            ranSel.FormattedText = pair.Key.FormattedText
                        End If
                    End If
                Next
                Dim dicTemp As New Dictionary(Of Word.Range, String)
                For Each pair As KeyValuePair(Of Word.Range, String) In dictVal.Value.dCharStyleInfo
                    dicTemp.Add(pair.Key, pair.Value)
                Next
                Dim sortedDict = (From entry In dicTemp Order By entry.Key.Start Ascending).ToDictionary(Function(pair) pair.Key, Function(pair) pair.Value)

                Dim sRemoveList() = {"[Author]", "[Editor]", "[Translator]"}

                Dim authorCount = dicTemp.Values.Where(Function(x) x.Contains("auSurName")).Count
                Dim editorCount = dicTemp.Values.Where(Function(x) x.Contains("edSurName")).Count
                Dim transCount = dicTemp.Values.Where(Function(x) x.Contains("trSurName")).Count

                Dim CurAuthorCnt As Integer = authorCount
                Dim CurEditorCnt As Integer = editorCount
                Dim CurTranslatorCnt As Integer = transCount
                Dim nameFound As Boolean

                For Each sStr In sRemoveList
                    Dim AuthorDoc As Word.Document = wAppp.Documents.Add(Visible:=True)
                    AuthorDoc.Activate()
                    AuthorDoc.Range.Select()
                    wAppp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                    nameFound = False
                    Dim ranSubsidiary As Word.Range
                    For Each pair As KeyValuePair(Of Word.Range, String) In sortedDict
                        If pair.Value.Contains("ref_foreTitle") Then '''for Prefix
                            ranSubsidiary = pair.Key
                        ElseIf pair.Value.Contains("ref_subsidiaryName") Then '''
                            ranSubsidiary = pair.Key
                        End If
                        If sStr.Contains("Author") Then
                            If Not IsNothing(ranSubsidiary) Then
                                wAppp.Selection.Range.FormattedText = ranSubsidiary.FormattedText
                                wAppp.Selection.Text = " "
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                ranSubsidiary = Nothing
                            End If
                            If pair.Value.Contains("ref_au") Then
                                nameFound = True
                                If authorCount >= 2 And CurAuthorCnt = 1 And pair.Value.Contains("auSurName") And Not sortedDict.Values.Contains("ref_etal") Then
                                    wAppp.Selection.Range.Text = GetAuthorPattern("Author", pair.Value, sRefStyle, dictVal.Value.sParaStyleName, sConfigPath, "AND")
                                    wAppp.Selection.Range.Style = "Normal"
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                    wAppp.Selection.Text = " "
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                End If

                                If pair.Value.Contains("auSurName") Then
                                    wAppp.Selection.Range.FormattedText = pair.Key.FormattedText
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                    If wAppp.Selection.Previous.Text = "." Or wAppp.Selection.Previous.Text = "," Or wAppp.Selection.Previous.Text = " " Then
                                        wAppp.Selection.Previous.Delete()
                                    End If
                                    wAppp.Selection.Range.Text = GetAuthorPattern("Author", pair.Value, sRefStyle, dictVal.Value.sParaStyleName, sConfigPath, "NAME")
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                Else   '''''''''''''GivenName
                                    wAppp.Selection.Range.FormattedText = pair.Key.FormattedText
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                    If wAppp.Selection.Previous.Text = "." Or wAppp.Selection.Previous.Text = "," Or wAppp.Selection.Previous.Text = " " Then
                                        wAppp.Selection.Previous.Delete()
                                    End If
                                    If (CurAuthorCnt = 0 Or (CurAuthorCnt = 1 And authorCount = 1)) And Not sortedDict.Values.Contains("ref_etal") Then
                                        wAppp.Selection.Range.Text = ". "
                                    ElseIf (CurAuthorCnt = 0 Or (CurAuthorCnt = 1 And authorCount = 1)) And Not sortedDict.Values.Contains("ref_etal") Then
                                        '''''' neeed to write et al 
                                        wAppp.Selection.Text = "et al."
                                        wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                    Else
                                        wAppp.Selection.Range.Text = GetAuthorPattern("Author", pair.Value, sRefStyle, dictVal.Value.sParaStyleName, sConfigPath, "NAME")
                                    End If
                                    wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                End If

                                If pair.Value.Contains("auSurName") Then
                                    CurAuthorCnt = CurAuthorCnt - 1
                                End If
                            Else
                                ''''''''''''''''
                            End If
                        ElseIf sStr.Contains("Editor") Then
                            If Not IsNothing(ranSubsidiary) Then
                                wAppp.Selection.Range.FormattedText = ranSubsidiary.FormattedText
                                wAppp.Selection.Text = " "
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                ranSubsidiary = Nothing
                            End If
                            If pair.Value.Contains("ref_ed") Then
                                nameFound = True
                                wAppp.Selection.Range.FormattedText = pair.Key.FormattedText
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                If wAppp.Selection.Previous.Text = "." Or wAppp.Selection.Previous.Text = "," Or wAppp.Selection.Previous.Text = " " Then
                                    wAppp.Selection.Previous.Delete()
                                End If
                                wAppp.Selection.Range.Text = GetAuthorPattern("Editor", pair.Value, sRefStyle, dictVal.Value.sParaStyleName, sConfigPath, "NAME")
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                If pair.Value.Contains("edSurName") Then
                                    editorCount = editorCount - 1
                                End If
                            Else
                                ''''''''''''''
                            End If
                        ElseIf sStr.Contains("Translator") Then
                            If Not IsNothing(ranSubsidiary) Then
                                wAppp.Selection.Range.FormattedText = ranSubsidiary.FormattedText
                                wAppp.Selection.Text = " "
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                ranSubsidiary = Nothing
                            End If
                            If pair.Value.Contains("ref_tr") Then
                                nameFound = True
                                wAppp.Selection.Range.FormattedText = pair.Key.FormattedText
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                If wAppp.Selection.Previous.Text = "." Or wAppp.Selection.Previous.Text = "," Or wAppp.Selection.Previous.Text = " " Then
                                    wAppp.Selection.Previous.Delete()
                                End If
                                wAppp.Selection.Range.Text = GetAuthorPattern("Translator", pair.Value, sRefStyle, dictVal.Value.sParaStyleName, sConfigPath, "NAME")
                                wAppp.Selection.Move(Word.WdUnits.wdParagraph, 1)
                                If pair.Value.Contains("trSurName") Then
                                    transCount = transCount - 1
                                End If
                            Else
                                ''''''''''
                            End If
                        End If
                    Next
                    If nameFound = True Then
                        ranDoc = AuthorDoc.Range
                        If ranDoc.Characters.Last.Text = Chr(13) Then ranDoc.SetRange(ranDoc.Start, ranDoc.End - 1)
                        ranDoc.Copy()
                        AuthorDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)

                        tmpDoc.Activate()
                        ranRef = tmpDoc.Content
                        With ranRef.Find
                            .ClearFormatting() : .Replacement.ClearFormatting()
                            .Text = sStr
                        End With
                        If ranRef.Find.Execute Then
                            ranRef.Select()
                            nameFound = True
                            Dim ranSel As Word.Range = wAppp.Selection.Range
                            ranSel.PasteAndFormat(Word.WdPasteOptions.wdKeepSourceFormatting)
                        End If
                    Else
                        AuthorDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                    End If
                Next

                ranDoc = tmpDoc.Range
                If ranDoc.Characters.Last.Text = Chr(13) Then ranDoc.SetRange(tmpDoc.Range.Start, tmpDoc.Range.End - 1)
                ranDoc.Copy()
                aDoc.Activate()

                Dim hnryRng As Word.Range
                hnryRng = dictVal.Key
                oCmt = aDoc.Comments.Add(dictVal.Key)
                Do While oCmt.Range.Paragraphs(1).Range.Fields.Count > 0
                    oCmt.Range.Paragraphs(1).Range.Fields(1).Delete()
                Loop
                oCmt.Author = "StyleRef"
                oCmt.Range.FormattedText = dictVal.Key.FormattedText
                dictVal.Key.PasteAndFormat(Word.WdPasteOptions.wdKeepSourceFormatting)
                tmpDoc.Range.Delete()
                aDoc.UndoClear()
            Next
            tmpDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            dictResult = dictReferInfo.Where(Function(x) x.Value.sPattern = Nothing Or x.Value.sMessage <> Nothing Or x.Value.sValid = False).ToList
            If dictResult.Count > 0 Then
                For Each dictVal In dictResult
                    oCmt = aDoc.Comments.Add(dictVal.Key)
                    Do While oCmt.Range.Paragraphs(1).Range.Fields.Count > 0
                        oCmt.Range.Paragraphs(1).Range.Fields(1).Delete()
                    Loop
                    oCmt.Author = "StyleRef"
                    If String.IsNullOrEmpty(dictVal.Value.sMessage) Then
                        If String.IsNullOrEmpty(dictVal.Value.sPattern) Then
                            oCmt.Range.Text = "Reference Pattern not found in the INI file."
                        End If
                    Else
                        oCmt.Range.Text = dictVal.Value.sMessage
                    End If

                Next
            End If
        Catch ex As Exception

        End Try
    End Function
    ''Regex.Match(sAuPattern, "@(.*?)$").Groups(1).Value)
    Public Function GetAuthorPattern(iniInput As String, inputVal As String, sRefSTyle As String, srefParaStyle As String, sConfigPath As String, sSpl As String) As String
        Try
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sAuPattern = oReadINI.INIReadValue(sRefSTyle & "_" & srefParaStyle, iniInput)
            If Not String.IsNullOrEmpty(sAuPattern) Then
                If sSpl = "NAME" Then
                    GetAuthorPattern = Regex.Match(sAuPattern, "\[" & inputVal.Replace("ref_", "") & "\]\<(.*?)\>").Groups(1).Value
                ElseIf sSpl = "AND" Then
                    GetAuthorPattern = Regex.Match(sAuPattern, "@(.*?)$").Groups(1).Value
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Public Function GetPatternfromINI(ByVal dictChar As Dictionary(Of Word.Range, String), sParaStyle As String, sRefStyle As String, sConfigPath As String) As String
        Try
            Dim I As Integer
            Dim possibleStylePattern As New List(Of String)
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sSkipList() = oReadINI.INIReadValue("RefStyleStructure", "Skip_StyleList").Split("|")
            Dim dictDup As New Dictionary(Of Word.Range, String)
            For Each pair As KeyValuePair(Of Word.Range, String) In dictChar
                dictDup.Add(pair.Key, pair.Value)
            Next

            I = 1
            Do While (True)
                Dim sRule = oReadINI.INIReadValue(sRefStyle & "_" & sParaStyle, "Pattern" & I)
                If sRule = String.Empty Then Exit Do

                Dim sRemoveList() = {"[Author]", "[Editor]", "[Translator]"}
                Dim sAutherType As String = sRule
                Dim SAuthorTypeList As New List(Of String)
                Dim sPatternwithoutAuthor As String = sRule

                For Each sStr In sRemoveList
                    sPatternwithoutAuthor = sPatternwithoutAuthor.Replace(sStr, "")
                Next
                'Dim sPatternwithoutAuthor As String = Regex.Replace("\#(.*?)\#", sRule, "")
                For K = dictChar.Count - 1 To 0 Step -1
                    If sSkipList.Contains(dictChar.ElementAt(K).Value) Then
                        dictChar.Remove(dictChar.ElementAt(K).Key)
                    End If
                Next


                Dim rMatches As MatchCollection
                Dim rMatch As Match
                Dim patternStyleList As New List(Of String)
                rMatches = Regex.Matches(sPatternwithoutAuthor, "\[(.*?)\]")
                For Each rMatch In rMatches
                    patternStyleList.Add("ref_" & rMatch.Groups(1).Value)
                Next
                rMatches = Nothing

                For Each st In sRemoveList
                    If sRule.Contains(st) Then
                        SAuthorTypeList.Add(st)
                    End If
                Next

                'rMatches = Regex.Matches(sAutherType, "\[(.*?)\]")
                'For Each rMatch In rMatches
                '    If sRemoveList.Contains(rMatch.Value) Then
                '        SAuthorTypeList.Add(rMatch.Value)
                '    End If
                'Next

                Dim flagName As Boolean
                If patternStyleList.Count = dictChar.Count Then
                    possibleStylePattern.Add(sRule)
                    For Each pair As KeyValuePair(Of Word.Range, String) In dictChar
                        If Not patternStyleList.Contains(pair.Value) Then
                            flagName = True
                        End If
                    Next
                    Dim dictList = dictChar.Values.ToList()
                    For Each st In patternStyleList
                        If Not dictList.Contains(st) Then
                            flagName = True
                        End If
                    Next

                    For Each st In SAuthorTypeList
                        Dim flag As Boolean
                        If st.Contains("Author") Then
                            For Each dst In dictDup.Values
                                If dst.Contains("ref_au") Then flag = True
                            Next
                        ElseIf st.Contains("Editor") Then
                            For Each dst In dictDup.Values
                                If dst.Contains("ref_ed") Then flag = True
                            Next
                        ElseIf st.Contains("Translator") Then
                            For Each dst In dictDup.Values
                                If dst.Contains("ref_tr") Then flag = True
                            Next
                        End If
                        If flag = False Then
                            flagName = True
                        End If
                    Next
                    If flagName = False Then
                        GetPatternfromINI = sRule
                        Exit Do
                    End If
                End If
                I = I + 1
            Loop


        Catch ex As Exception

        End Try
    End Function
    Public Function CollectReferenceInformation(wdoc As Word.Document, sConfigPath As String, sRefStyle As String)
        Try
            Dim ranDoc As Word.Range
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sTemp = oReadINI.INIReadValue("RefStyleStructure", "StyleName")
            ''Dim CEGRefStyle() = {"REF:CONFERENCE", "REF:PERIODICAL", "REF:WEBLINK", "REF:BK", "REF:JART", "REF:BKCH", "REF"}
            Dim I As Integer
            Dim CEGRefStyle() = sTemp.Split("|")
            For I = LBound(CEGRefStyle) To UBound(CEGRefStyle)
                If CEGRefStyle(I) <> String.Empty Then
                    ranDoc = wdoc.Range.Duplicate
                    If AutoStyleExists(CEGRefStyle(I), wdoc) = True Then
                        With ranDoc.Find
                            .ClearFormatting() : .Replacement.ClearFormatting()
                            .Text = "" : .Style = CEGRefStyle(I).Trim()
                        End With
                        Do While ranDoc.Find.Execute
                            ranDoc.Select()
                            Dim ranSel = wAppp.Selection.Range
                            Dim dCharInfo As New Dictionary(Of Word.Range, String)
                            Dim sErrorMessage As String
                            dCharInfo = CollectCharacterStyleInfo(ranSel, sConfigPath, sRefStyle, CEGRefStyle(I), sErrorMessage)
                            ranDoc.Select()
                            Dim fValid = CheckUnstyledContent(ranSel)
                            Dim objCls = New clsRefStyleInfo
                            If Not dCharInfo Is Nothing Then
                                Dim dTempChar As New Dictionary(Of Word.Range, String)
                                For Each pair As KeyValuePair(Of Word.Range, String) In dCharInfo
                                    dTempChar.Add(pair.Key, pair.Value)
                                Next

                                objCls.dCharStyleInfo = dTempChar
                                Dim sPattern = GetPatternfromINI(dCharInfo, CEGRefStyle(I), sRefStyle, sConfigPath)
                                objCls.sParaStyleName = CEGRefStyle(I)
                                objCls.sValid = fValid
                                If fValid = False Then
                                    objCls.sMessage = "Unstyled content found in the reference."
                                End If
                                If Not String.IsNullOrEmpty(sPattern) Then
                                    objCls.sPattern = sPattern
                                End If
                                If ranSel.Characters.Last.Text = Chr(13) Then ranSel.SetRange(ranSel.Start, ranSel.End - 1)
                                dictReferInfo.Add(ranSel, objCls)
                            Else
                                If sErrorMessage <> String.Empty Then
                                    objCls.sMessage = "The Character style not found in the document."
                                Else
                                    objCls.sMessage = sErrorMessage
                                End If
                                dictReferInfo.Add(ranSel, objCls)
                                End If
                                If ranDoc.End > wdoc.Range.End Then Exit Do
                            ranDoc = wdoc.Range(ranDoc.End + 1, wdoc.Range.End)
                            ranDoc.Find.Text = ""
                            ranDoc.Find.Style = CEGRefStyle(I)
                        Loop
                    Else
                        MessageBox.Show(CEGRefStyle(I) & " Style not found in the document. Please check the configuration.")
                    End If

                End If
            Next
        Catch ex As Exception

        End Try
    End Function

    Public Function CollectCharacterStyleInfo(ranRef As Word.Range, sConfigPath As String, srefStyle As String, sStyleName As String, ByRef sErrorMessage As String) As Dictionary(Of Word.Range, String)
        Try
            Dim drefSt As New Dictionary(Of Word.Range, String)
            Dim ranDup As Word.Range
            ranDup = ranRef.Duplicate
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sTemp = oReadINI.INIReadValue("RefStyleStructure", sStyleName)
            Dim StyleNameList() = sTemp.Split("|")
            For i = LBound(StyleNameList) To UBound(StyleNameList)
                If StyleNameList(i) <> String.Empty Then
                    ranRef = ranDup.Duplicate
                    If modCEGUtility.AutoStyleExists(StyleNameList(i), wAppp.ActiveDocument) = True Then
                        With ranRef.Find
                            .ClearFormatting() : .Replacement.ClearFormatting()
                            .Text = "" : .Style = StyleNameList(i)
                        End With
                        Do While ranRef.Find.Execute
                            ranRef.Select()
                            Dim ranSel As Word.Range
                            ranSel = wAppp.Selection.Range
                            If ranSel.Characters.Last.Text = Chr(13) Then ranSel.SetRange(ranSel.Start, ranSel.End - 1)
                            If ranSel.Text <> "" And ranSel.InRange(ranDup) Then
                                drefSt.Add(ranSel, StyleNameList(i))
                            Else
                                '' MessageBox.Show("fsdfsfsf")
                            End If
                            If ranRef.End >= ranDup.End Then Exit Do
                            If Not ranRef.InRange(ranDup) Then
                                Exit Do
                            End If
                            ranRef.SetRange(ranSel.End + 1, ranDup.End)
                            ranRef.Find.ClearFormatting()
                            ranRef.Find.Text = ""
                            ranRef.Find.Style = StyleNameList(i)
                        Loop
                    Else
                        If sErrorMessage = String.Empty Then
                            sErrorMessage = sErrorMessage & "" & StyleNameList(i)
                        Else
                            sErrorMessage = sErrorMessage & ", " & StyleNameList(i)
                        End If

                    End If
                End If
            Next
            CollectCharacterStyleInfo = drefSt
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Public Function CheckUnstyledContent(ranRef As Word.Range) As Boolean
        Try
            Dim CharStyle As Word.Style
            Dim ParaStyle As Word.Style
            Dim bTemp As Boolean = True
            Dim ranSel As Word.Range
            ranSel = ranRef.Duplicate
            For I = 1 To ranSel.Words.Count
                CharStyle = ranSel.Words(I).Style
                ParaStyle = ranSel.Words(I).ParagraphStyle
                If CharStyle.NameLocal = ParaStyle.NameLocal Then
                    Dim sTemp = Regex.Replace(ranSel.Words(I).Text.Replace(vbCr, ""), "[?.&,;!¡¿。、. ·'"":()\[\]]+", "")
                    sTemp = sTemp.Replace("and", "")
                    sTemp = sTemp.Replace("vol", "")
                    sTemp = sTemp.Replace("no", "")
                    sTemp = sTemp.Replace("pp", "")
                    sTemp = sTemp.Replace("in", "")
                    sTemp = sTemp.Replace("eds", "")
                    sTemp = sTemp.Replace("ed", "")
                    sTemp = sTemp.Replace(vbTab, "")
                    If Not String.IsNullOrEmpty(sTemp) Then
                        bTemp = False
                        ''MessageBox.Show("Word: " & sTemp)
                    End If
                End If
            Next
            CheckUnstyledContent = bTemp
        Catch ex As Exception

        End Try
    End Function

End Module
