
Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports CEGINI

Public Class frmCitationStyle
    Public oWordApp As Word.Application
    Public PubName As String
    Public JrnName As String
    Public sConfigPath As String
    Public curRefIndex As Integer
    Public objCitationStyle As clsCitationStyle
    Public sFirstCitation As String
    Public sRemainCitation As String
    Public ImpStyleList As ArrayList
    Public lstCitationInDoc As ArrayList
    Public lstCitationtoBe As ArrayList
    Public lstWrongCitation As ArrayList

    Public Sub New(oWApp As Word.Application, pName As String, JName As String, sConPath As String)
        oWordApp = oWApp
        PubName = pName
        JrnName = JName
        sConfigPath = sConPath
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Private Sub frmCitationStyle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sAutPattern As String
        Dim sWrongAuthorPattern As String
        Try
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sCitStyle = oReadINI.INIReadValue(PubName & "@" & JrnName, "CitationStyle")
            objCitationStyle = New clsCitationStyle()
            objCitationStyle.sCitStyle = sCitStyle.Split("|")(0)
            objCitationStyle.sCitTextType = sCitStyle.Split("|")(1)
            objCitationStyle.sCitSep = sCitStyle.Split("|")(2)
            'If sCitStyle.Split("|")(0).ToUpper() = "CMS" Then
            '    sCitationStyle = 3
            'ElseIf sCitStyle.Split("|")(0).ToUpper() = "APA" Then
            '    cCitationAuthorCount = 5
            'Else
            '    MessageBox.Show("Citation style not defined " & sCitStyle.Split("|")(0))
            '    Return
            'End If

            curRefIndex = 1
            dictRefInfo = New Dictionary(Of Integer, clsRefInfo)
            CollectRefInfo1(oWordApp.ActiveDocument, "†Reference")
            If ModRefUtility.dictRefInfo.Count > 0 Then
                rtbReference.Text = dictRefInfo(curRefIndex).oRefRng.Text
                sAutPattern = GetPatternCitationRule(curRefIndex)
                sWrongAuthorPattern = GetPatternWrongCitationRule(curRefIndex)
                If sAutPattern <> String.Empty Then
                    lstCitationInDoc = New ArrayList
                    lstCitationtoBe = New ArrayList
                    lstWrongCitation = New ArrayList
                    GetCitationAsperRule(curRefIndex)
                    DisplayCitationInformation(sAutPattern)
                    If sWrongAuthorPattern <> String.Empty Then
                        DisplayWrongCitationInformation(sWrongAuthorPattern)
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Function DisplayCitationInformation(sPattern As String)
        Dim wDoc As Word.Document
        Dim rMatches As MatchCollection
        wDoc = oWordApp.ActiveDocument
        Dim I As Integer
        Dim oRng As Word.Range
        Try
            For I = 1 To 3
                Select Case I
                    Case 1 : oRng = wDoc.Content.Duplicate
                    Case 2 : If wDoc.Footnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory).Duplicate Else oRng = Nothing
                    Case 3 : If wDoc.Endnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory).Duplicate Else oRng = Nothing
                End Select
                If Not IsNothing(oRng) Then
                    Dim ranDoc As Word.Range
                    ranDoc = oRng.Duplicate
                    rMatches = Regex.Matches(oRng.Text, sPattern)

                    For Each mat As Match In rMatches
                        ranDoc = oRng.Duplicate
                        ranDoc.Find.ClearFormatting()
                        ranDoc.Find.Text = mat.Value
                        Do While ranDoc.Find.Execute
                            ranDoc.Select()
                            lstCitationInDoc.Add(oWordApp.Selection.Range)
                        Loop
                    Next

                    For Each ranCit As Word.Range In lstCitationInDoc
                        Dim n As Integer = dgvCitationResult.Rows.Add()
                        dgvCitationResult.Rows.Item(n).Cells(1).Value = ranCit.Text
                    Next
                End If
            Next
            ''''''''''''''''''''''''''''''''''''''''
            For I = 0 To dgvCitationResult.RowCount - 2
                If objCitationStyle.sCitTextType.ToUpper() = "FIRST" And I = 0 Then
                    dgvCitationResult.Rows(I).Cells(2).Value = sFirstCitation
                ElseIf objCitationStyle.sCitTextType.ToUpper() = "FIRST" And I <> 0 Then
                    dgvCitationResult.Rows(I).Cells(2).Value = sRemainCitation
                ElseIf objCitationStyle.sCitTextType.ToUpper() = "ALL" Then
                    dgvCitationResult.Rows(I).Cells(2).Value = sRemainCitation
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function DisplayWrongCitationInformation(sPattern As String)
        Dim wDoc As Word.Document
        Dim rMatches As MatchCollection
        wDoc = oWordApp.ActiveDocument
        Dim I As Integer
        Dim oRng As Word.Range
        Try
            For I = 1 To 3
                Select Case I
                    Case 1 : oRng = wDoc.Content.Duplicate
                    Case 2 : If wDoc.Footnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory).Duplicate Else oRng = Nothing
                    Case 3 : If wDoc.Endnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory).Duplicate Else oRng = Nothing
                End Select
                If Not IsNothing(oRng) Then
                    Dim ranDoc As Word.Range
                    ranDoc = oRng.Duplicate
                    rMatches = Regex.Matches(oRng.Text, sPattern)

                    For Each mat As Match In rMatches
                        If CheckCitationInDocList(mat.Value) = False Then
                            ranDoc = oRng.Duplicate
                            ranDoc.Find.ClearFormatting()
                            ranDoc.Find.Text = mat.Value
                            Do While ranDoc.Find.Execute
                                ranDoc.Select()
                                lstWrongCitation.Add(oWordApp.Selection.Range)
                            Loop
                        End If
                    Next
                    For Each ranCit As Word.Range In lstWrongCitation
                        lbWrongCitation.Items.Add(ranCit.Text)
                    Next
                End If

            Next
        Catch ex As Exception

        End Try
    End Function

    Function CheckCitationInDocList(sString As String) As Boolean
        Try
            For Each ranCit As Word.Range In lstCitationInDoc
                If (sString = ranCit.Text) Then
                    CheckCitationInDocList = True : Exit Function
                End If
            Next
            CheckCitationInDocList = False : Exit Function
        Catch ex As Exception

        End Try
    End Function
    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        Dim sAutPattern As String
        Dim sWrongAuthorPattern As String
        Try
            If dgvCitationResult.RowCount > 0 Then dgvCitationResult.Rows.Clear()
            If lbWrongCitation.Items.Count > 0 Then lbWrongCitation.Items.Clear()

            curRefIndex = curRefIndex + 1
            If dictRefInfo.Count < curRefIndex Then
            Else
                rtbReference.Text = dictRefInfo(curRefIndex).oRefRng.Text
                sAutPattern = GetPatternCitationRule(curRefIndex)
                sWrongAuthorPattern = GetPatternWrongCitationRule(curRefIndex)
                If sAutPattern <> String.Empty Then
                    lstCitationInDoc = New ArrayList
                    lstCitationtoBe = New ArrayList
                    lstWrongCitation = New ArrayList
                    GetCitationAsperRule(curRefIndex)
                    DisplayCitationInformation(sAutPattern)
                    DisplayWrongCitationInformation(sWrongAuthorPattern)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        Dim sAutPattern As String
        Dim sWrongAuthorPattern As String
        Try
            If dgvCitationResult.RowCount > 0 Then dgvCitationResult.Rows.Clear()
            If lbWrongCitation.Items.Count > 0 Then lbWrongCitation.Items.Clear()
            curRefIndex = curRefIndex - 1
            If dictRefInfo.Count < curRefIndex Then
            Else
                rtbReference.Text = dictRefInfo(curRefIndex).oRefRng.Text
                sAutPattern = GetPatternCitationRule(curRefIndex)
                sWrongAuthorPattern = GetPatternWrongCitationRule(curRefIndex)
                If sAutPattern <> String.Empty Then
                    lstCitationInDoc = New ArrayList
                    lstCitationtoBe = New ArrayList
                    lstWrongCitation = New ArrayList
                    GetCitationAsperRule(curRefIndex)
                    DisplayCitationInformation(sAutPattern)
                    DisplayWrongCitationInformation(sWrongAuthorPattern)
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Function CollectRefInfo1(wDoc As Word.Document, refStyleName As String) As Boolean
        Try
            Const sTrimPattern As String = "([\[\]\(\)\t\s\.\:]+)"
            Dim fStyle() = {"‡ref_auSurname", "‡ref_auCollab"}
            Dim ranDoc As Word.Range : Dim rCount As Integer
            ranDoc = wDoc.Content
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = refStyleName
            End With
            Do While ranDoc.Find.Execute = True

                Dim AuthorList As New List(Of String)
                Dim dictAuthor As New Dictionary(Of Integer, String)
                Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
                ranDoc.Select()
                ranDupRef = oWordApp.Selection.Range
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''Validation
                Dim ranVal As Word.Range
                ranVal = ranDoc.Duplicate
                If Regex.IsMatch(ranVal.Text, "^(—)+") Then
                    MessageBox.Show(" Em dashes found instead of author names! Please change em dashes into author names and try again!!", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
                'Dim skipStyle() = {"‡ref_auPrefix", "‡ref_edPrefix", "‡ref_transPrefix", "‡ref_transedPrefix", "‡ref_assigneePrefix", "‡ref_compilerPrefix", "‡ref_directorPrefix", "‡ref_guestedPrefix", "‡ref_inventorPrefix"}
                'For i = LBound(skipStyle) To UBound(skipStyle)
                '    If modCEGUtility.AutoStyleExists(skipStyle(i), wDoc) = True Then
                '        ranVal = ranDoc.Duplicate
                '        With ranVal.Find
                '            .ClearFormatting() : .Replacement.ClearFormatting()
                '            .Text = "" : .Style = skipStyle(i)
                '        End With
                '        If ranVal.Find.Execute = True Then
                '            MessageBox.Show(skipStyle(i) & " style found! Please remove author prefix style and try again!!", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                '            Return False
                '        End If
                '    End If
                'Next
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                For i = LBound(fStyle) To UBound(fStyle)
                    ranAuthor = ranDupRef.Duplicate
                    With ranAuthor.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With
                    Do While ranAuthor.Find.Execute
                        isFoundSurname = True
                        'If ranAuthor.Previous.Words.Count > 0 Then
                        '    Dim ranSuffix As Word.Range
                        '    ranAuthor.Previous.Words(1).Select()
                        '    ranSuffix = oActApp.Selection.Range
                        '    With ranSuffix.Find
                        '        .ClearFormatting() : .Replacement.ClearFormatting()
                        '        .Text = "" : .Style = "‡ref_auPrefix"
                        '    End With
                        '    If ranSuffix.Find.Execute Then
                        '        ranSuffix.Select()
                        '    End If
                        'End If
                        ranAuthor.Select()
                        If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
                            dictAuthor.Add(ranAuthor.Start, ranAuthor.Text)
                        End If
                        If ranAuthor.End > ranDupRef.End Then Exit Do
                        ranAuthor = wDoc.Range(ranAuthor.End + 1, ranDupRef.End)
                        ranAuthor.Find.Text = ""
                        ranAuthor.Find.Style = fStyle(i)
                    Loop
                Next
                If CheckCitationStyleAuthorCount(dictAuthor.Count) = True Then
                    '' If dictAuthor.Count = cCitationAuthorCount Then
                    rCount = rCount + 1
                    Dim ranRef As Word.Range
                    Dim sYear As String = ""
                    Dim objRefInfo As New clsRefInfo
                    ranRef = ranDoc.Duplicate
                    With ranRef.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = "‡ref_pubdateYear"
                    End With
                    If ranRef.Find.Execute = True Then
                        sYear = ranRef.Text
                    End If
                    ranRef = ranDoc.Duplicate
                    With ranRef.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = "‡ref_etal"
                    End With
                    If ranRef.Find.Execute = True Then
                        objRefInfo.sRefEtal = True
                    Else
                        objRefInfo.sRefEtal = False
                    End If

                    If refStyleName = "‡ref_auCollab" Then
                        objRefInfo.sRefCollab = True
                    Else
                        objRefInfo.sRefCollab = False
                    End If

                    If isFoundSurname = True Then
                        dictAuthor.OrderBy(Function(item) item.Key)
                        Dim val As String
                        Dim sCondText As String = ""
                        For Each val In dictAuthor.Values
                            AuthorList.Add(val)
                            sCondText = sCondText & Trim(val)
                        Next
                        objRefInfo.rCondText = sCondText & sYear
                    Else
                        objRefInfo.rCondText = ""
                    End If
                    objRefInfo.oRefRng = ranDoc.Duplicate
                    objRefInfo.cText = Regex.Replace(ranDoc.Text, "\n[\n\s\r\t]*", String.Empty, RegexOptions.IgnoreCase) ''ranDoc.Text.Replace(ChrW(13), "")
                    objRefInfo.sRefYear = sYear
                    objRefInfo.olRefAuthors = AuthorList
                    If (dictRefInfo.ContainsKey(rCount)) = False Then
                        dictRefInfo.Add(rCount, objRefInfo)
                    End If
                End If
                If ranDoc.End >= wDoc.Range.End Then Exit Do
                ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
                ranDoc.Find.Text = ""
                ranDoc.Find.Style = refStyleName
                dictAuthor = Nothing
            Loop
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function
    Function CheckCitationStyleAuthorCount(aCount As Integer) As Boolean
        Try
            If objCitationStyle.sCitStyle.ToUpper() = "CMS" Then
                If aCount = 3 Then Return True
            ElseIf objCitationStyle.sCitStyle.ToUpper() = "APA" Then
                If aCount > 2 And aCount < 6 Then Return True
            ElseIf objCitationStyle.sCitStyle.ToUpper() = "HARVARD" Then
                If aCount > 2 Then Return True
            End If
            Return False
        Catch ex As Exception

        End Try
    End Function
    Function GetCitationAsperRule(cIndex As Integer)
        Try
            ''Public sFirstCitation As String
            ''Public sRemainCitation As String
            Dim strTemp As String
            ''First
            If dictRefInfo(cIndex).olRefAuthors.Count > 0 And dictRefInfo(cIndex).sRefYear <> String.Empty Then
                If objCitationStyle.sCitStyle.ToUpper() = "CMS" Then
                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        If dictRefInfo(cIndex).olRefAuthors.Count = dictRefInfo(cIndex).olRefAuthors.IndexOf(strAut) Then
                            strTemp = strTemp & strAut & "###"
                        Else
                            strTemp = strTemp & strAut & ", "
                        End If
                    Next
                    sFirstCitation = strTemp & " " & dictRefInfo(cIndex).sRefYear
                ElseIf objCitationStyle.sCitStyle.ToUpper() = "APA" Then
                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        If dictRefInfo(cIndex).olRefAuthors.Count = dictRefInfo(cIndex).olRefAuthors.IndexOf(strAut) Then
                            strTemp = strTemp & strAut & "###"
                        Else
                            strTemp = strTemp & strAut & ", "
                        End If
                    Next
                    sFirstCitation = strTemp & " " & dictRefInfo(cIndex).sRefYear
                ElseIf objCitationStyle.sCitStyle.ToUpper() = "HARVARD" Then
                    sFirstCitation = dictRefInfo(cIndex).olRefAuthors(1).ToString() & "  EETTAALL., " & dictRefInfo(cIndex).sRefYear
                End If
            End If
            '######################'ALL or Remaining Author #######################################################################
            If dictRefInfo(cIndex).olRefAuthors.Count > 0 And dictRefInfo(cIndex).sRefYear <> String.Empty Then
                Dim cc As Integer : cc = 0 : strTemp = ""
                If objCitationStyle.sCitStyle.ToUpper() = "CMS" Or objCitationStyle.sCitStyle.ToUpper() = "APA" Then

                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        cc = cc + 1
                        If cc = 1 Then
                            strTemp = strTemp & strAut : Exit For
                        End If
                    Next
                    sRemainCitation = sRemainCitation & strTemp & " et al.,"

                ElseIf objCitationStyle.sCitStyle.ToUpper() = "HARVARD" Then

                End If
            End If
            ''#######################################################################################################################
        Catch ex As Exception

        End Try
    End Function
    Function GetPatternWrongCitationRule(cIndex As Integer) As String
        Try
            Dim sAuthorPattern As String
            Dim K As Integer
            Try
                Dim srtTemp As String
                If dictRefInfo(cIndex).olRefAuthors.Count > 0 And dictRefInfo(cIndex).sRefYear <> String.Empty Then
                    ''''Total number of author
                    If dictRefInfo(cIndex).olRefAuthors.Count > 2 Then
                        srtTemp = ""
                        For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                            srtTemp = srtTemp & strAut & "(,)?(\s)?(\s|&|and|)?(\s)+(\w+)?(\s)?"
                        Next
                        If srtTemp <> String.Empty Then
                            sAuthorPattern = sAuthorPattern & srtTemp & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"
                        End If
                    End If
                    ''''Two author
                    srtTemp = ""
                    If dictRefInfo(cIndex).olRefAuthors.Count > 1 Then
                        sAuthorPattern = sAuthorPattern & "|"
                        For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                            srtTemp = srtTemp & strAut & "(,)?(\s)?(\s|&|and|)?(\s)+(\w+)?(\s)?"
                        Next
                        If srtTemp <> String.Empty Then
                            sAuthorPattern = sAuthorPattern & srtTemp & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"
                        End If
                    End If
                    '''''' Single author et al
                    sAuthorPattern = sAuthorPattern & "|"
                    srtTemp = ""
                    Dim cc As Integer
                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        cc = cc + 1
                        If cc = 1 Then
                            srtTemp = srtTemp & strAut & "(\s|,)?" : Exit For
                        End If

                    Next
                    sAuthorPattern = sAuthorPattern & srtTemp & " et al(\.,|\.|,)?(\s)?" & "(\()?" & dictRefInfo(cIndex).sRefYear & "((,)?([a-z]+)?(\s)?([0-9]?)(&)?)+" & "(\))?"

                    ''''Single author
                    sAuthorPattern = sAuthorPattern & "|"
                    sAuthorPattern = sAuthorPattern & srtTemp & " " & "(\()?" & dictRefInfo(cIndex).sRefYear & "((,)?([a-z]+)?(\s)?([0-9]?)(&)?)+" & "(\))?"
                    sAuthorPattern = Replace(sAuthorPattern, "||", "|")
                    If InStr(1, sAuthorPattern, "|") = 1 Then
                        sAuthorPattern = Mid(sAuthorPattern, 2)
                    End If

                ElseIf dictRefInfo(cIndex).sRefCollab = True And dictRefInfo(cIndex).sRefYear <> String.Empty Then

                Else
                    MsgBox("1111111111111111111")
                End If
                GetPatternWrongCitationRule = sAuthorPattern
            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try
    End Function
    Function GetPatternCitationRule(cIndex As Integer) As String
        Dim sAuthorPattern As String
        Dim K As Integer
        Try
            Dim srtTemp As String
            If dictRefInfo(cIndex).olRefAuthors.Count > 0 And dictRefInfo(cIndex).sRefYear <> String.Empty Then
                ''''Total number of author
                If dictRefInfo(cIndex).olRefAuthors.Count > 2 Then
                    srtTemp = ""
                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        srtTemp = srtTemp & strAut & "(,)?(\s)?(\s|&|and|)?(\s)+"
                    Next
                    If srtTemp <> String.Empty Then
                        sAuthorPattern = sAuthorPattern & srtTemp & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"
                    End If

                End If
                ''''Two author
                srtTemp = ""
                If dictRefInfo(cIndex).olRefAuthors.Count > 1 Then
                    sAuthorPattern = sAuthorPattern & "|"
                    For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                        srtTemp = srtTemp & strAut & "(,)?(\s)?(\s|&|and|)?(\s)+"
                    Next
                    If srtTemp <> String.Empty Then
                        sAuthorPattern = sAuthorPattern & srtTemp & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"
                    End If

                End If
                '''''' Single author et al
                sAuthorPattern = sAuthorPattern & "|"
                srtTemp = ""
                Dim cc As Integer
                For Each strAut As String In dictRefInfo(cIndex).olRefAuthors
                    cc = cc + 1
                    If cc = 1 Then
                        srtTemp = srtTemp & strAut & "(\s|,)?" : Exit For
                    End If

                Next
                sAuthorPattern = sAuthorPattern & srtTemp & " et al(\.|,|\.,)? " & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"

                ''''Single author
                sAuthorPattern = sAuthorPattern & "|"
                sAuthorPattern = sAuthorPattern & srtTemp & " " & "(\()?" & dictRefInfo(cIndex).sRefYear & "(\))?"
                sAuthorPattern = Replace(sAuthorPattern, "||", "|")
                If InStr(1, sAuthorPattern, "|") = 1 Then
                    sAuthorPattern = Mid(sAuthorPattern, 2)
                End If

            ElseIf dictRefInfo(cIndex).sRefCollab = True And dictRefInfo(cIndex).sRefYear <> String.Empty Then

            Else
                MsgBox("1111111111111111111")
            End If
            GetPatternCitationRule = sAuthorPattern
        Catch ex As Exception

        End Try
    End Function

    Private Sub dgvCitationResult_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvCitationResult.CellClick
        Try
            Dim ranCitation As Word.Range = lstCitationInDoc.Item(e.RowIndex)
            ranCitation.Select()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub lbWrongCitation_Click(sender As Object, e As EventArgs) Handles lbWrongCitation.Click
        Try
            Dim ranCitation As Word.Range = lstWrongCitation.Item(lbWrongCitation.SelectedIndex)
            ranCitation.Select()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub lbWrongCitation_MouseMove(sender As Object, e As MouseEventArgs)
        ''

    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click

    End Sub

    Private Sub lbWrongCitation_MouseClick(sender As Object, e As MouseEventArgs) Handles lbWrongCitation.MouseClick
        Dim ranCitation As Word.Range = lstWrongCitation.Item(lbWrongCitation.SelectedIndex)
    End Sub
End Class
Public Class clsCitationStyle
    Public sCitStyle As String
    Public sCitTextType As String
    Public sCitSep As String
    Public sCitEtal As String
    Public sCitBraceTypeText As String
End Class