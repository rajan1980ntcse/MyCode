Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq
Imports System.Windows
Imports <xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
Imports Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices


Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'Dim sDocxFullName As String = "C:\temp\Check\CARCIN-2012-00096.R1.docx"
        'Using oDocx As WordprocessingDocument = WordprocessingDocument.Open(sDocxFullName, False)
        '    MessageBox.Show(oDocx.MainDocumentPart.Document.Body.InnerXml)
        '    'MessageBox.Show(oDocx.MainDocumentPart.)
        '    ' Insert other code here. 
        'End Using

        Dim ocls As New clsCEGMiscWML
        '||CH2_BM_9710038033123.docx||EM1_9710038033123.docx
        'ocls.ToGetUsedStyle4DocName("CH1_BM_9710038033123.docx||CH2_BM_9710038033123.docx||EM1_9710038033123.docx||EM2_9710038033123.docx||FM1_9710038033123.docx", "C:\AutoProcess\CUPBITS", "CUP", "9710038033123", "An Introduction to the Science of the Mind")


        Dim oclsCEG As New ClsCEGWML
        Dim osDdict As New SortedDictionary(Of String, String)
        Dim smsg As String
        Dim sFN As String = "E:\OUP_Backlist\UPSO\test1\a.docx"
        ''oclsCEG.ToGetDocXUsedStyleList(sFN, smsg)



        ocls.ToGetUsedStyleForEditingFramework("CH1_BM_9781628925807.docx||CH2_BM_9781628925807.docx||CH3_BM_9781628925807.docx||CH4_BM_9781628925807.docx||CH5_BM_9781628925807.docx||CH6_BM_9781628925807.docx||EM1_9781628925807.docx||EM2_9781628925807.docx||FM1_9781628925807.docx||FM2_9781628925807.docx", "D:\AutoProcess\CUP\9781628925807", "CUP", "9710038033123", "An Introduction to the Science of the Mind")
        ''ocls.ToGetUsedStyle("CH1_BM_9781107156999.docx||CH2_BM_9781107156999.docx||CH3_BM_9781107156999.docx", "D:\AutoProcess\CUP\9781107156999", "CUP", "9710038033123", "An Introduction to the Science of the Mind")


        'Dim oL As SortedDictionary(Of String, Integer)
        'Dim sMsg As String = String.Empty
        'Dim fileName = "C:\temp\MS ID.docx" '"C:\temp\Misc\New\TEST.docx"  '"C:\temp\Check\TEST.docx"
        'oL = ocls.ToGetDocXUsedStyleList(fileName, sMsg)
        'MessageBox.Show(sMsg)
        'MessageBox.Show(oL.Join)
        'Dim I As Integer
        'Dim J As Integer = Math.DivRem(80, 8, I).ToString()

        'If I > 0 Then J = J + 1
        'MessageBox.Show(J.ToString)


        MessageBox.Show(Environment.Is64BitOperatingSystem)

    End Sub


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        'Dim myStream As Stream = Nothing
        'Dim openFileDialog1 As New OpenFileDialog()

        ''openFileDialog1.InitialDirectory = "c:\"
        'openFileDialog1.Filter = "Docx files (*.docx)|*.docx"
        'openFileDialog1.FilterIndex = 1
        'openFileDialog1.RestoreDirectory = True

        'If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '    Try
        '        Dim oCls As New clsCEGMiscWML : Dim oL As SortedDictionary(Of String, Integer)
        '        Dim sFileName As String = openFileDialog1.FileName
        '        Dim sMsg As String = String.Empty
        '        'oL = oCls.ToGetDocXUsedStyleList(sFileName, sMsg)
        '        Dim sStyleList As String = "DocName : " & sFileName & vbCrLf
        '        For Each oKV As KeyValuePair(Of String, Integer) In oL
        '            sStyleList = sStyleList & oKV.Key & " (" & oKV.Value & ")" & vbCrLf
        '        Next
        '        MessageBox.Show(sStyleList)
        '    Catch Ex As Exception
        '        MessageBox.Show("Error: " & Ex.Message)
        '    Finally
        '        ' Check this again, since we need to make sure we didn't throw an exception on open. 
        '        If (myStream IsNot Nothing) Then
        '            myStream.Close()
        '        End If
        '    End Try
        'End If
    End Sub

    Private Sub TEST_Click(sender As System.Object, e As System.EventArgs) Handles TEST.Click

        Try


            Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
            'Dim WordDoc As Document = WordApp.ActiveDocument
            Dim objCls As New clsCEGMisc
            ''objCls.ToCallRefStyleTypeMain(WordApp, "", "", "Book", "Springer", "C:\Program Files (x86)\Newgen\CEGenius")
            'objCls.ToCallSortReferenceCitation(WordDoc, WordApp, "IOP", "PMB", "C:\Program Files (x86)\Newgen\CEGenius\Main\Journal_Config.ini")
            ''' objCls.ToCallReferenceStyler(WordDoc, WordApp)'' by jaisoft

            Dim wDoc As Microsoft.Office.Interop.Word.Document
            wDoc = WordApp.ActiveDocument
            ReferencePunctuationMain(WordApp, wDoc)

            ''  modCEGUtility.CheckFMAuthorWithDiscloserAuthor(wDoc)
            '' modCEGUtility.ReportUsedFontFromListOfDocument(WordApp, "Portugal.docx||Mozambique.docx", wDoc.Path)
            ''objCls.ToCallReferenceFormatingMain(wDoc, "LWW", "XAA", "C:\Program Files (x86)\Newgen\CEGenius\Main\Journal_Config.ini")
            ''objCls.ToCallReferenceCitationStyleChecking(wDoc, WordApp, "OUP", "BIOLIN", "C:\Program Files (x86)\Newgen\CEGenius\Main\Journal_Config.ini")
            'Dim oclsab As New clsCEGMisc
            'Dim sfiles As String = "01_9780190657895_Ekelund_Chap 01.doc||13_9780190657895_Ekelund_BM.doc"
            ''sfiles = ""
            ''For Each ft In Directory.GetFiles("E:\temp\2015\Feb\OUP Book\Styled2")
            ''    sfiles = sfiles & Path.GetFileName(ft) & "||"
            ''Next
            'modCEGUtility.FigCaptionLogCreation(WordApp, sfiles, "D:\AutoProcess\OUPBook\Ekelund161216ATUS_MSC\9780190657895")
            'oclsab.ToExportMixedCitation(WordApp, sfiles, "E:\temp\2015\Feb\OUP Book\Styled2")
            ''oclsab.OUPNumberingParaFormatQC(WordApp, sfiles, "c:\temp\test")


            'Dim sDocName As String = "batinic-9781107091078pre.docx||batinic-9781107091078int.docx||batinic-9781107091078c01.docx||batinic-9781107091078c02.docx||batinic-9781107091078c03.docx||batinic-9781107091078c04.docx||batinic-9781107091078c05.docx||batinic-9781107091078con.docx||batinic-9781107091078bib.docx"
            'Dim sDocDirPath As String = "C:\AutoProcess\CUP\TEST"
            'Dim sMergedDocName As String = "342342342342342_Merged.docx"  ''"9781107091078_Merged.docx"

            'Dim oclsMerge As New clsCEGMisc
            ''oclsMerge.ToMergeDocument(WordApp, sDocName, sDocDirPath, sMergedDocName, False)



            'Dim sStyleDocx As String = "C:\AutoProcess\CUP\TEST\342342342342342_Merged.docx"  ''Path.Combine(sDocDirPath, sMergedDocName)
            'Dim oclsRep As New clsRepInfo
            'oclsRep.sDocxFullFileName = sStyleDocx
            'oclsRep.sISBN = "9781107091078"
            'oclsRep.sBKTitle = "Women and Yugoslav Partisans"
            'oclsRep.sLevel1HeadingList = "(§CT|13.02 CT)"
            'oclsRep.sLevel2HeadingList = "(§A|02.01 A)"
            'oclsRep.sLevel3HeadingList = "(§B|02.02 B)"
            'oclsRep.sLevel4HeadingList = "(§C|02.03 C)"
            'MessageBox.Show(oclsMerge.ToGenerate
            '' Log(oclsRep))
            'WordApp = Nothing
            'Dim ocls As New clsQCTool
            'ocls.ToCallJournalBasedQC(WordDoc, "OCCMED")
            Dim oclsab As New clsCEGMisc
            Dim sfiles As String = "CH1_BM_7878787878787.doc||CH2_BM_7878787878787.doc"
            '' oclsab.ToCallRefStyleTypeMain(WordApp, sfiles, "C:\AutoProcess\7878787878787", "Book", "Springer", "C:\Program Files (x86)\Newgen\CEGenius")
            '' oclsab.ToCallRefLinkTypeForm(WordApp, sfiles, "C:\AutoProcess\7878787878787")
            ''RefStyleModifier.ReferenceStyleModifierMain(WordDoc, WordApp)
            'Dim ocls As New ClsCEGWML
            'Dim sDocList As String = "D:\Bugs\TEST_WML\431604.docx"
            'Dim oclsWMLInfo As New clsCEGWMLInfo
            'Dim sMsg As String
            'Call ocls.ToGetDocXParaList(sDocList, oclsWMLInfo, sMsg)
            ''testxml()
            'ToGetWordFiles()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Me.Close()
        End Try
    End Sub


    Private Sub btnIndex_Click(sender As System.Object, e As System.EventArgs) Handles btnIndex.Click
        Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
        Dim oMDoc As Document = WordApp.ActiveDocument
        Dim ocls As New CEGMisc.clsCEGMisc
        Call ocls.ToCallCheckOMath(WordApp, oMDoc)

        'Dim oclsIndex As New clsTransformIndexTerm
        'oclsIndex.ToReviewIndex(oMDoc)
        'oclsIndex.ToTransformIndexTerm()
        'Dim ofrm As New frmIndexTransfer
        'ofrm.ShowDialog()


    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
        Dim oMDoc As Document = WordApp.ActiveDocument
        Dim ocls As New CEGMisc.clsCEGMisc
        ocls.ToCallVancouver2HavardCitation(oMDoc, WordApp, "OUP", "ANNWEH", "C:\Program Files (x86)\Newgen\CEGenius\Main\Journal_Config.ini")
        'Dim ocls As New frmLanguageSelect(WordApp, "", "")
        'ocls.Show()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnFtnFWord_Click(sender As Object, e As EventArgs) Handles btnFtnFWord.Click
        Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
        Dim oMDoc As Document = WordApp.ActiveDocument
        Dim ocls As New CEGMisc.clsCEGMisc
        ocls.ToCallShadingFirstwordOfFootnoteText(oMDoc, WordApp)

    End Sub
End Class
