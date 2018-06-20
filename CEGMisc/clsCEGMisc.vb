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
Imports oWrd = Microsoft.Office.Interop.Word
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class clsCEGMisc
#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class and its COM interfaces. If you change them, existing clients will no longer be able to access the class.
    Public Const ClassId As String = "dcc69d57-3cfa-40c1-b58e-829f397b11bc"
    Public Const InterfaceId As String = "EDED909C-9271-4670-BA32-109AE917B1D6"
    Public Const EventsId As String = "17C731B8-CE61-4B5F-B114-10F3E46153AD"
#End Region

    ''

    Public Function ToCallShadingFirstwordOfFootnoteText(wDoc As oWrd.Document, WordApp As oWrd.Application) As Boolean
        Try
            If modCEGUtility.ToShadingFootnoteText(wDoc, WordApp) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallRemoveAllBookmarks(oWordDoc As oWrd.Document, oWordApp As oWrd.Application) As Boolean
        Try
            If modCEGUtility.ToRemoveAllBookmarks(oWordDoc, oWordApp) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function ToCallUsedFontReport(WordApp As oWrd.Application, wLstFiles As String, wFPath As String) As Boolean
        Try
            If modCEGUtility.ReportUsedFontFromListOfDocument(WordApp, wLstFiles, wFPath) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallCheckFMAuthorwithDiscloserAuthor(wDoc As oWrd.Document) As Boolean
        Try
            If modCEGUtility.CheckFMAuthorWithDiscloserAuthor(wDoc) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    ''*********0040684: [IOP Journals] Add publisher's note in comment************
    Public Function ToCallAddPublisherNoteInComment(PubName As String, JName As String, wDoc As oWrd.Document, JConfigPath As String) As Boolean
        Try
            If modCEGUtility.InsertPublisherNoteasComment(PubName, JName, wDoc, JConfigPath) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    ''*********0040684: [IOP Journals] Add publisher's note in comment************
    Public Function ToCallReferenceFormatingMain(wDoc As oWrd.Document, pName As String, jName As String, JConfig As String)
        Try
            Call ModRefUtility.ReferenceGranularFormating(wDoc, wDoc.Application, pName, jName, JConfig)

        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallRefStyleTypeMain(oWordApp As oWrd.Application, wLstFiles As String, wFPath As String, sRefStyle As String, pName As String, sCEGPath As String) As Boolean
        Try
            Dim tempRefLinkType As String = ""
            Dim curdocPath As String
            Dim curDoc As oWrd.Document

            Dim objFrm As New frmRefStyleType(oWordApp, wLstFiles, wFPath, sRefStyle, pName, sCEGPath)
            If vbOK = objFrm.ShowDialog() Then

            End If

            '''''''''''''''''''''''''''''''''
            'If sRefStyle.ToUpper() = "BOOK" Then
            '    For Each fDoc As String In wLstFiles.Split("||")
            '        If fDoc <> String.Empty Then
            '            curdocPath = Path.Combine(wFPath, fDoc)
            '            oWordApp.Documents(curdocPath).Activate()
            '            curDoc = oWordApp.ActiveDocument
            '            Dim objFrm As New frmRefStyleType(oWordApp, wLstFiles, wFPath, sRefStyle, pName, sCEGPath)
            '            If vbOK = objFrm.ShowDialog() Then
            '                tempRefLinkType = objFrm.gRefStyleType
            '                If tempRefLinkType <> "" Then Exit For
            '            Else
            '                Exit Function
            '            End If
            '        End If
            '    Next
            '    For Each fDoc As String In wLstFiles.Split("||")
            '        If fDoc <> String.Empty Then
            '            curdocPath = Path.Combine(wFPath, fDoc)
            '            oWordApp.Documents(curdocPath).Activate()
            '            curDoc = oWordApp.ActiveDocument
            '            If VariableExists(curDoc, "CEGRefStyleType") = False Then
            '                curDoc.Variables.Add("CEGRefStyleType", tempRefLinkType)
            '            Else
            '                curDoc.Variables("CEGRefStyleType").Value = tempRefLinkType
            '            End If
            '        End If
            '    Next
            'ElseIf sRefStyle.ToUpper() = "CHAPTER" Then
            '    For Each fDoc As String In wLstFiles.Split("||")
            '        If fDoc <> String.Empty Then
            '            curdocPath = Path.Combine(wFPath, fDoc)
            '            oWordApp.Documents(curdocPath).Activate()
            '            curDoc = oWordApp.ActiveDocument
            '            Dim objFrm As New frmRefStyleType(curDoc, "")
            '            If vbOK = objFrm.ShowDialog() Then
            '                tempRefLinkType = objFrm.gRefStyleType
            '                If VariableExists(curDoc, "CEGRefStyleType") = False Then
            '                    curDoc.Variables.Add("CEGRefStyleType", tempRefLinkType)
            '                Else
            '                    curDoc.Variables("CEGRefStyleType").Value = tempRefLinkType
            '                End If
            '            Else
            '                MessageBox.Show("Reference Style information variable not set in the document.." + curDoc.Name)
            '            End If
            '        End If
            '    Next
            'Else
            '    MessageBox.Show("Ref style should be book/Chapter.. Please check..")
            'End If

            ''''''''''''''''''''''''''''''''

        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallRefLinkTypeForm(oWordApp As oWrd.Application, wLstFiles As String, wFPath As String) As Boolean
        Try
            Dim objfrm As New frmRefLinkType(oWordApp, wLstFiles, wFPath)
            If vbOK = objfrm.ShowDialog() Then
                Dim curdocPath As String
                Dim curDoc As oWrd.Document
                If objfrm.srefLinkType <> "" Then
                    For Each fDoc As String In wLstFiles.Split("||")
                        If fDoc <> String.Empty Then
                            '' MsgBox(wFPath)
                            curdocPath = Path.Combine(wFPath, fDoc)
                            oWordApp.Documents(curdocPath).Activate()
                            curDoc = oWordApp.ActiveDocument
                            If VariableExists(curDoc, "CEGRefLinkType") = False Then
                                curDoc.Variables.Add("CEGRefLinkType", objfrm.srefLinkType)
                            Else
                                curDoc.Variables("CEGRefLinkType").Value = objfrm.srefLinkType
                            End If
                        End If
                    Next
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallSetLanguageforBook(oWordApp As oWrd.Application, wLstFiles As String, wFPath As String)
        Try
            Dim objfrm As New frmLanguageSelect(oWordApp, wLstFiles, wFPath)
            objfrm.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallReferenceStyler(wDoc As oWrd.Document, oWordApp As oWrd.Application)
        Try
            Dim objfrm As New frmReferenceStyler(wDoc, oWordApp)
            objfrm.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function CEGWebRequest(sURL As String, sSearch As String) As String
        Return modCEGUtility.CEGWebQuery(sURL, sSearch)
    End Function
    Public Function ToCallFormatChangeAfterRSTinReference(wDoc As oWrd.Document, oWordApp As oWrd.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Call modCEGUtility.FormatChangeAfterRSTinReference(jName, pName, oWordApp, wDoc, sConfigPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallVancouver2HavardCitation(wDoc As oWrd.Document, oWordApp As oWrd.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Call ModRefUtility.Vancouver2HarvardCitationMain(wDoc, oWordApp, pName, jName, sConfigPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function ToCallSortReferenceCitation(wDoc As oWrd.Document, oWordApp As oWrd.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Call ModRefUtility.ReferenceCitationSort(wDoc, oWordApp, pName, jName, sConfigPath)
            ''MessageBox.Show("Completed....")
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallRefArrageAlphabetical(wDoc As oWrd.Document, oWordApp As oWrd.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Call ModRefUtility.RefArrageAlphabetical(wDoc, oWordApp, pName, jName, sConfigPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    ''ReferenceCitationStyleChecking
    Public Function ToCallReferenceCitationStyleChecking(wDoc As oWrd.Document, oWordApp As oWrd.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Call modCitationStyleChecker.ReferenceCitationStyleChecking(wDoc, oWordApp, pName, jName, sConfigPath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallOUPFigcaptionlogCreation(oWordApp As oWrd.Application, wlistFiles As String, wFilePath As String)
        Try
            modCEGUtility.FigCaptionLogCreation(oWordApp, wlistFiles, wFilePath)
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function ToCallOUPRemoveShadingAndStyles(sCEGPath As String, oActDoc As oWrd.Document)
        Try
            Dim oclsRem As New RemoveStyles.clsRemoveStyles
            If oclsRem.RemoveStyles(sCEGPath, oActDoc) = True Then
                Call AddQCIteminCollection("RemoveStyles", oActDoc)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function ToCallOUPAbstractKeywordMain(oWordApp As oWrd.Application, wListFiles As String, wFilePath As String, sXMLPath As String)
        Dim oclsAbs As New clsAbsractKeyword
        Call oclsAbs.OUPAbstractKeywordMain(oWordApp, wListFiles, wFilePath, sXMLPath)
    End Function

    Public Function ToCallOUPLawStyleOrganizerMain(oWordApp As oWrd.Application, wListFiles As String, wFilePath As String, styleINI As String)
        'Dim styleINI As String = "C:\Program Files\newgen\CEGenius\Main\Styles\OUP Law Pilot preediting.ini"
        Try
            If File.Exists(styleINI) Then
                Dim ofrmStyle = New frmStyleImport(oWordApp, wListFiles, wFilePath, styleINI)
                ofrmStyle.Show()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function


    Public Function ToCallCheckOMath(oActApp As oWrd.Application, oActDoc As oWrd.Document) As Boolean
        Dim sclsMsgTitle As String = sMsgTitle & " - Mathtype Warning"
        Try
            If oActDoc.Range.XML.Contains("m:oMath") = True OrElse oActDoc.Range.XML.Contains("equationxml=") Then
                MessageBox.Show("Non-MathType equation(s) found in file." & Environment.NewLine & "Please convert to MathType format before proceeding.", sclsMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf oActDoc.Range.XML.Contains("equationxml") = True Then
                MessageBox.Show("Non-MathType equation(s) is present as an image." & Environment.NewLine & "Please convert to MathType format before proceeding.", sclsMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            'If CInt(oActApp.Version) > 11 AndAlso oActDoc.OMaths.Count > 0 Then
            '    For Each oM As oWrd.OMath In oActDoc.OMaths
            '        oM.Range.HighlightColorIndex = oWrd.WdColorIndex.wdRed
            '    Next
            'End If
            oActDoc.UndoClear()
            Call AddQCIteminCollection("ToCheckOMath", oActDoc)
            ToCallCheckOMath = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, sclsMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>
    ''' This process for merge all chapter documents with one version for CUP final process
    ''' </summary>
    ''' <param name="oActApp"></param>
    ''' <param name="oActDoc"></param>
    ''' <param name="sDocNameList"></param>
    ''' <param name="sDocDirPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ToMergeDocument(oActApp As oWrd.Application, sDocNameList As String, sDocDirPath As String, sMergedDocName As String, Optional bIsOpenDocProcess As Boolean = True) As Boolean
        Dim oActDoc As oWrd.Document = Nothing
        Dim osdFiles As New SortedDictionary(Of Integer, clsFile)
        Dim oDocList As String() = sDocNameList.Replace("||", "|").Split(New Char() {"|"}).ToArray()
        Try
            Dim iTtlCnt As Integer = 0 : Dim sMissingDocName As String = String.Empty
            If CInt(oActApp.Version) <= 11 Then
                sMergedDocName = Regex.Replace(sMergedDocName, "\.docx", ".doc", RegexOptions.IgnoreCase)
            End If
            Dim sMergedDocFullName = Path.Combine(sDocDirPath, sMergedDocName)
            Dim sTempFileFullName As String = Path.Combine(Path.GetTempPath(), sMergedDocName)

            For Each sDocName As String In oDocList
                If String.IsNullOrWhiteSpace(sDocName) = False Then
                    Dim oclsFile As New clsFile
                    Dim sFullName As String = Path.Combine(sDocDirPath, sDocName)
                    oclsFile.sFileName = sDocName
                    oclsFile.sFullName = sFullName
                    oclsFile.sDirFullName = sDocDirPath
                    If Not oActApp Is Nothing AndAlso bIsOpenDocProcess = True Then
                        For Each oDoc As oWrd.Document In oActApp.Documents
                            If sFullName.ToLower = oDoc.FullName.ToLower Then
                                oDoc.Save() : oclsFile.oWrdDoc = oDoc 'oActApp.Documents(sFullName) : 'oActApp.Documents(sDocName).Save()
                            End If
                            If sMergedDocName.ToLower = oDoc.FullName.ToLower Then oDoc.Close(oWrd.WdSaveOptions.wdSaveChanges)
                        Next
                        If oclsFile.oWrdDoc Is Nothing Then
                            sMissingDocName = sMissingDocName + vbCrLf + sDocName
                        End If
                    End If
                    osdFiles.Add(iTtlCnt, oclsFile) : iTtlCnt += 1
                End If
            Next
            If oActApp Is Nothing = False AndAlso String.IsNullOrWhiteSpace(sMissingDocName.Trim()) = False AndAlso bIsOpenDocProcess = True Then
                MessageBox.Show("The following document is missing in the active application." & vbCrLf & "Name : " & sMissingDocName, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return False
                Exit Function
            End If
            '''''''' Set application settings ''''''''''''''
            If bIsOpenDocProcess = False Then
                oActApp = New oWrd.Application() : System.Threading.Thread.Sleep(750)
                oActDoc = oActApp.Documents.Add()
                oActApp.Visible = False : oActApp.DisplayAlerts = oWrd.WdAlertLevel.wdAlertsNone
            Else
                oActDoc = oActApp.Documents.Add()
                oActApp.Visible = True : oActApp.DisplayAlerts = oWrd.WdAlertLevel.wdAlertsNone
                oActDoc.ActiveWindow.Visible = True
            End If

            If Regex.IsMatch(sMergedDocName, "\.docx$", RegexOptions.IgnoreCase) = True Then
                oActDoc.SaveAs(sTempFileFullName, oWrd.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles:=False)
            Else
                oActDoc.SaveAs(sTempFileFullName, oWrd.WdSaveFormat.wdFormatDocument, AddToRecentFiles:=False)
            End If
            '''For Each oAddIns As oWrd.AddIn In oActApp.AddIns : oAddIns.Installed = False : Next
            oActDoc.Content.FootnoteOptions.NumberingRule = oWrd.WdNumberingRule.wdRestartSection
            oActDoc.TrackRevisions = False
            If bIsOpenDocProcess = True Then
                For Each oKV As KeyValuePair(Of Integer, clsFile) In osdFiles
                    Dim oDoc As oWrd.Document = oKV.Value.oWrdDoc
                    Dim oRng As oWrd.Range = oActDoc.Range(oActDoc.Content.Start, oActDoc.Content.End).Duplicate
                    oRng.SetRange(oRng.End - 1, oRng.End)
                    oRng.FormattedText = oDoc.Range()
                    oRng.InsertParagraphAfter() : oRng.SetRange(oRng.End - 1, oRng.End) : oRng.Font.Reset()
                    oRng.InsertBreak(oWrd.WdBreakType.wdSectionBreakContinuous)
                Next
            Else
                For Each oKV As KeyValuePair(Of Integer, clsFile) In osdFiles
                    Dim oRng As oWrd.Range = oActDoc.Range(oActDoc.Content.Start, oActDoc.Content.End).Duplicate
                    oRng.SetRange(oRng.End - 1, oRng.End)
                    oRng.InsertFile(oKV.Value.sFullName)
                    oRng.InsertParagraphAfter() : oRng.SetRange(oRng.End - 1, oRng.End) : oRng.Font.Reset()
                    oRng.InsertBreak(oWrd.WdBreakType.wdSectionBreakContinuous)
                Next
            End If
            oActDoc.Save()
            If File.Exists(sTempFileFullName) = True Then File.Copy(sTempFileFullName, sMergedDocFullName, True)




        Catch ex As Exception
            MessageBox.Show("Unable to merge the document" + "Error : " + ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If Not oActDoc Is Nothing Then oActDoc.Close(oWrd.WdSaveOptions.wdDoNotSaveChanges)
            If bIsOpenDocProcess = False Then oActApp.Quit()
        End Try
    End Function


    Public Function ToGenerateStyleReportLogWithArguments(sISBN As String, sBKTitle As String, sDocxFullFileName As String, sLevel1HeadingList As String, sLevel2HeadingList As String, sLevel3HeadingList As String, sLevel4HeadingList As String) As Boolean
        Try
            If sDocxFullFileName.ToLower.EndsWith(".doc") Then
                MessageBox.Show("Please run this process with 'DOCX' file format.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
                Exit Function
            End If

            Dim sStyleMsg As String = String.Empty
            Dim oclsCEGWML As New clsCEGWML : Dim oclsDoc As New clsDocInfo
            Dim sDocxStyleReportFullFileName As String = Path.Combine(Path.GetDirectoryName(sDocxFullFileName), sISBN & "_Styles_Report.html")
            Dim sDocxRunningHeadReportFullFileName As String = Path.Combine(Path.GetDirectoryName(sDocxFullFileName), sISBN & "_RH_Report.html")
            If File.Exists(sDocxStyleReportFullFileName) = True Then File.Delete(sDocxStyleReportFullFileName)
            oclsDoc = oclsCEGWML.ToGetDocXUsedStyleList(sDocxFullFileName, sStyleMsg)
            '/////// Collecting Verso and Recto Info //////////
            Dim sRunHeadReportInfo As String = String.Empty
            sLevel1HeadingList = Regex.Escape(Regex.Replace(sLevel1HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel1HeadingList = "(" + sLevel1HeadingList.Replace("\|", "|") + ")\b"
            sLevel2HeadingList = Regex.Escape(Regex.Replace(sLevel2HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel2HeadingList = "(" + sLevel2HeadingList.Replace("\|", "|") + ")\b"
            sLevel3HeadingList = Regex.Escape(Regex.Replace(sLevel3HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel3HeadingList = "(" + sLevel3HeadingList.Replace("\|", "|") + ")\b"
            sLevel4HeadingList = Regex.Escape(Regex.Replace(sLevel4HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel4HeadingList = "(" + sLevel4HeadingList.Replace("\|", "|") + ")\b"
            If oclsDoc Is Nothing Then
                MessageBox.Show("Unable to collect style information from the document", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            For Each oT As Tuple(Of String, String) In oclsDoc.oParaTextWithOrder
                Dim sStyleName As String = oT.Item1
                If Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) Then
                    Dim sTemp As String = Regex.Replace(oT.Item2, "[\s\t\n]+", String.Empty, RegexOptions.IgnoreCase)
                    If sTemp.Length < 50 Then
                        Select Case True
                            Case Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td>" & oT.Item2 & "</td><td/><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Else 'Nothing 
                        End Select
                    Else
                        Select Case True
                            Case Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td>" & oT.Item2 & "</td><td/><td>Verso Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Else 'Nothing 
                        End Select
                    End If
                End If
            Next
            If String.IsNullOrEmpty(sRunHeadReportInfo) = False Then
                sRunHeadReportInfo = "<p/><table width='95%' border='1' align='center'><tr bgcolor='#666666'><td align='center' colspan='3'><b>List of running heads</b></td></tr><tr><td><b>Verso</b></td><td><b>Recto</b></td><td><b>Remarks&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></td></tr>" & sRunHeadReportInfo & "</table>"
                sRunHeadReportInfo = sRunHeadReportInfo.Replace("&#60;", "&nbsp;")
            End If


            '/////// Collecting Style information ////////////
            Dim sStylesReportInfo As String = String.Empty
            For Each oKV As KeyValuePair(Of String, Integer) In oclsDoc.oParaUsedStyleList
                sStylesReportInfo = sStylesReportInfo & "<tr><td>" & oKV.Key & "</td><td>&#160;</td><td/></tr>"
            Next
            Dim oReg As New Regex("(\&\#160;)")
            For Each oKV As KeyValuePair(Of String, Integer) In oclsDoc.oCharUsedStyleList
                If sStylesReportInfo.Contains("&#160;") = True Then
                    sStylesReportInfo = oReg.Replace(sStylesReportInfo, oKV.Key, 1)
                Else
                    sStylesReportInfo = sStylesReportInfo & "<tr><td>&#60;</td><td>" & oKV.Key & "</td><td/></tr>"
                End If
            Next
            If String.IsNullOrEmpty(sStylesReportInfo) = False Then
                sStylesReportInfo = "<p/><table width='95%' border='1' align='center'><tr bgcolor='#666666'><td align='center' colspan='3'><b>Styles used in book</b></td></tr><tr><td width='35%'><b>Paragraph</b></td><td width='35%'><b>Character</b></td><td width='20%'><b>Linked</b></td></tr>" & sStylesReportInfo & "</table>"
                sStylesReportInfo = sStylesReportInfo.Replace("&#60;", "&nbsp;")
            End If

            Dim sBKInfo As String = "<table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!-- ISBN --></td></tr><tr><td><b>Title</b></td><td>: <!-- Btitle --></td></tr><tr><td><b>Publisher</b></td><td>: CUP</td></tr><tr><td><b>Date and time</b></td><td>: " & Now & "</td></tr></tbody></table><br/><hr/></head><br/><body>"
            sBKInfo = sBKInfo.Replace("<!-- ISBN -->", sISBN) : sBKInfo = sBKInfo.Replace("<!-- Btitle -->", sBKTitle)


            Dim sRunningHeadLogInfo As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUP Report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'>CUP Running Head Report</h1>" & sBKInfo
            Dim sStyleLogInfo As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUP Report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'>CUP Styles Report</h1>" & sBKInfo

            sStyleLogInfo = sStyleLogInfo & sStylesReportInfo & "</body></html>"
            sRunningHeadLogInfo = sRunningHeadLogInfo & sRunHeadReportInfo & "</body></html>"

            File.WriteAllText(sDocxStyleReportFullFileName, sStyleLogInfo)
            File.WriteAllText(sDocxRunningHeadReportFullFileName, sRunningHeadLogInfo)
        Catch ex As Exception
            MessageBox.Show("Unable to generate the style report" + "Error : " + ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function ToGenerateStyleReportLog(oclsRepInfo As clsRepInfo) As Boolean
        Try
            Dim sDocxFullFileName As String = oclsRepInfo.sDocxFullFileName
            If sDocxFullFileName.ToLower.EndsWith(".doc") Then
                MessageBox.Show("Please run this process with 'DOCX' file format.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
                Exit Function
            End If

            Dim sStyleMsg As String = String.Empty : Dim oclsCEGWML As New clsCEGWML : Dim oclsDoc As New clsDocInfo
            Dim sDocxStyleReportFullFileName As String = Path.Combine(Path.GetDirectoryName(oclsRepInfo.sDocxFullFileName), oclsRepInfo.sISBN & "_Styles_Report.html")
            Dim sDocxRunningHeadReportFullFileName As String = Path.Combine(Path.GetDirectoryName(oclsRepInfo.sDocxFullFileName), oclsRepInfo.sISBN & "_RH_Report.html")
            If File.Exists(sDocxStyleReportFullFileName) = True Then File.Delete(sDocxStyleReportFullFileName)
            oclsDoc = oclsCEGWML.ToGetDocXUsedStyleList(sDocxFullFileName, sStyleMsg)
            '/////// Collecting Verso and Recto Info //////////
            Dim sRunHeadReportInfo As String = String.Empty
            Dim sLevel1HeadingList As String = Regex.Escape(Regex.Replace(oclsRepInfo.sLevel1HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel1HeadingList = "(" + sLevel1HeadingList.Replace("\|", "|") + ")\b"
            Dim sLevel2HeadingList As String = Regex.Escape(Regex.Replace(oclsRepInfo.sLevel2HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel2HeadingList = "(" + sLevel2HeadingList.Replace("\|", "|") + ")\b"
            Dim sLevel3HeadingList As String = Regex.Escape(Regex.Replace(oclsRepInfo.sLevel3HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel3HeadingList = "(" + sLevel3HeadingList.Replace("\|", "|") + ")\b"
            Dim sLevel4HeadingList As String = Regex.Escape(Regex.Replace(oclsRepInfo.sLevel4HeadingList, "[\(\)]", String.Empty, RegexOptions.IgnoreCase)) : sLevel4HeadingList = "(" + sLevel4HeadingList.Replace("\|", "|") + ")\b"
            If oclsDoc Is Nothing Then
                MessageBox.Show("Unable to collect style information from the document", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If
            For Each oT As Tuple(Of String, String) In oclsDoc.oParaTextWithOrder
                Dim sStyleName As String = oT.Item1
                If Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) OrElse Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) Then
                    Dim sTemp As String = Regex.Replace(oT.Item2, "[\s\t\n]+", String.Empty, RegexOptions.IgnoreCase)
                    If sTemp.Length < 50 Then
                        Select Case True
                            Case Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td>" & oT.Item2 & "</td><td/><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td/></tr>"
                            Case Else 'Nothing 
                        End Select
                    Else
                        Select Case True
                            Case Regex.IsMatch(sStyleName, sLevel1HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td>" & oT.Item2 & "</td><td/><td>Verso Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel2HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel3HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Regex.IsMatch(sStyleName, sLevel4HeadingList, RegexOptions.IgnoreCase) : sRunHeadReportInfo = sRunHeadReportInfo & "<tr><td/><td>" & oT.Item2 & "</td><td>Recto Page: Too Long</td></tr>"
                            Case Else 'Nothing 
                        End Select
                    End If
                End If
            Next
            If String.IsNullOrEmpty(sRunHeadReportInfo) = False Then
                sRunHeadReportInfo = "<p/><table width='95%' border='1' align='center'><tr bgcolor='#666666'><td align='center' colspan='3'><b>List of running heads</b></td></tr><tr><td><b>Verso</b></td><td><b>Recto</b></td><td><b>Remarks&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></td></tr>" & sRunHeadReportInfo & "</table>"
                sRunHeadReportInfo = sRunHeadReportInfo.Replace("&#60;", "&nbsp;")
            End If


            '/////// Collecting Style information ////////////
            Dim sStylesReportInfo As String = String.Empty
            For Each oKV As KeyValuePair(Of String, Integer) In oclsDoc.oParaUsedStyleList
                sStylesReportInfo = sStylesReportInfo & "<tr><td>" & oKV.Key & "</td><td>&#160;</td><td/></tr>"
            Next
            Dim oReg As New Regex("(\&\#160;)")
            For Each oKV As KeyValuePair(Of String, Integer) In oclsDoc.oCharUsedStyleList
                If sStylesReportInfo.Contains("&#160;") = True Then
                    sStylesReportInfo = oReg.Replace(sStylesReportInfo, oKV.Key, 1)
                Else
                    sStylesReportInfo = sStylesReportInfo & "<tr><td>&#60;</td><td>" & oKV.Key & "</td><td/></tr>"
                End If
            Next
            If String.IsNullOrEmpty(sStylesReportInfo) = False Then
                sStylesReportInfo = "<p/><table width='95%' border='1' align='center'><tr bgcolor='#666666'><td align='center' colspan='3'><b>Styles used in book</b></td></tr><tr><td width='35%'><b>Paragraph</b></td><td width='35%'><b>Character</b></td><td width='20%'><b>Linked</b></td></tr>" & sStylesReportInfo & "</table>"
                sStylesReportInfo = sStylesReportInfo.Replace("&#60;", "&nbsp;")
            End If


            Dim sBKInfo As String = "<table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!-- ISBN --></td></tr><tr><td><b>Title</b></td><td>: <!-- Btitle --></td></tr><tr><td><b>Publisher</b></td><td>: CUP</td></tr><tr><td><b>Date and time</b></td><td>: " & Now & "</td></tr></tbody></table><br/><hr/></head><br/><body>"
            sBKInfo = sBKInfo.Replace("<!-- ISBN -->", oclsRepInfo.sISBN) : sBKInfo = sBKInfo.Replace("<!-- Btitle -->", oclsRepInfo.sBKTitle)

            Dim sRunningHeadLogInfo As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUP Report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'>CUP Running Head Report</h1>" & sBKInfo
            Dim sStyleLogInfo As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUP Report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'>CUP Styles Report</h1>" & sBKInfo

            sStyleLogInfo = sStyleLogInfo & sStylesReportInfo & "</body></html>"
            sRunningHeadLogInfo = sRunningHeadLogInfo & sRunHeadReportInfo & "</body></html>"


            File.WriteAllText(sDocxStyleReportFullFileName, sStyleLogInfo)
            File.WriteAllText(sDocxRunningHeadReportFullFileName, sRunningHeadLogInfo)
            Return True
        Catch ex As Exception
            MessageBox.Show("Unable to generate the style report" + "Error : " + ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Public Function ToExportMixedCitation(oActApp As oWrd.Application, sDocNameList As String, sDocDirPath As String)
        Dim oActDoc As oWrd.Document = Nothing
        Dim oDocList As String() = sDocNameList.Replace("||", "|").Split(New Char() {"|"}).ToArray()
        Try

            For Each sDocName As String In oDocList
                Dim strBibText As String = ""
                If String.IsNullOrWhiteSpace(sDocName) = False Then
                    Dim sFullName As String = Path.Combine(sDocDirPath, sDocName)
                    If File.Exists(sFullName.ToLower().Replace(".doc", "_temp.doc")) Then
                        File.Delete(sFullName.ToLower().Replace(".doc", "_temp.doc"))
                    End If
                    File.Copy(sFullName, sFullName.ToLower().Replace(".doc", "_temp.doc"))
                    oActDoc = oActApp.Documents.Open(sFullName.ToLower().Replace(".doc", "_temp.doc"))
                    Dim ranDoc As oWrd.Range = oActDoc.Content
                    ranDoc.Find.ClearFormatting()
                    ranDoc.Find.Style = "†Reference"
                    ranDoc.Find.Text = ""
                    Do While (ranDoc.Find.Execute)
                        Dim ranBib As oWrd.Range
                        Dim ranDup As oWrd.Range
                        ranDoc.Select()
                        ranBib = oActApp.Selection.Range
                        ranDup = ranBib.Duplicate
                        ApplyFormatingTag(ranBib, oActApp, oActDoc)
                        strBibText += ranDup.Text
                    Loop
                    If strBibText <> "" Then
                        Dim sR As New StreamWriter(sFullName.Replace(Path.GetExtension(sFullName), ".BIB"))
                        sR.Write(strBibText)
                        sR.Close()
                    End If
                    oActDoc.Close(oWrd.WdSaveOptions.wdDoNotSaveChanges)
                    If File.Exists(sFullName.ToLower().Replace(".doc", "_temp.doc")) Then
                        File.Delete(sFullName.ToLower().Replace(".doc", "_temp.doc"))
                    End If
                    oActDoc = Nothing
                    oActApp.Documents(sFullName).Activate()
                    oActDoc = oActApp.ActiveDocument
                    oActDoc = oActApp.Documents(sDocName)
                    Call AddQCIteminCollection("ExportBIB", oActDoc)
                End If
            Next
        Catch ex As Exception

        End Try
    End Function
    Public Function ApplyFormatingTag(ranBib As oWrd.Range, oActApp As oWrd.Application, oActDoc As oWrd.Document)
        Try

            Dim ranFormat As oWrd.Range = ranBib.Duplicate
            ranBib.Find.ClearFormatting()
            ranBib.Find.Text = ""
            ranBib.Find.Font.Italic = True
            Do While (ranBib.Find.Execute)
                ranBib.Select()
                oActApp.Selection.Font.Italic = False
                oActApp.Selection.InsertBefore("<italic>")
                oActApp.Selection.InsertAfter("</italic>")
                ranBib.SetRange(oActApp.Selection.End + 1, ranFormat.End)
                ranBib.Find.Font.Italic = True
            Loop
            ranBib = ranFormat.Duplicate

            ranBib.Find.ClearFormatting()
            ranBib.Find.Text = ""
            ranBib.Find.Font.Bold = True
            Do While (ranBib.Find.Execute)
                ranBib.Select()
                oActApp.Selection.Font.Bold = False
                oActApp.Selection.InsertBefore("<bold>")
                oActApp.Selection.InsertAfter("</bold>")
                ranBib.SetRange(oActApp.Selection.End + 1, ranFormat.End)
                ranBib.Find.Font.Bold = True
            Loop
        Catch ex As Exception

        End Try
    End Function
End Class

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class clsRepInfo
    Public Property sISBN As String
    Public Property sBKTitle As String
    Public Property sDocxFullFileName As String
    Public Property sLevel1HeadingList As String
    Public Property sLevel2HeadingList As String
    Public Property sLevel3HeadingList As String
    Public Property sLevel4HeadingList As String
End Class

<ClassInterface(ClassInterfaceType.AutoDual)> _
Public Class clsFile
    Public Property sFullName As String
    Public Property sFileName As String
    Public Property sFilePath As String
    Public Property sDirFullName As String
    Public Property sDirStruct As String
    Public Property oWrdDoc As oWrd.Document
End Class


