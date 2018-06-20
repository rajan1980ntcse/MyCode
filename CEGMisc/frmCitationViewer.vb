Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Public Class frmCitationViewer
    Public wDoc As Word.Document
    Public wAPP As Word.Application
    Public htmlText As String
    Public bEtalItalic As Boolean
    Public Sub New(oDoc As Word.Document, oAPP As Word.Application, bEtal As Boolean)
        wDoc = oDoc
        wAPP = oAPP
        bEtalItalic = bEtal
        InitializeComponent()
    End Sub
    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click

        If ModRefUtility.dictConvertedCitation.Count > 0 Then
            For Each pair As KeyValuePair(Of Word.Range, String) In ModRefUtility.dictConvertedCitation
                wDoc.Range(pair.Key.Start, pair.Key.End).Select()
                If wAPP.Selection.Font.Superscript = True Then wAPP.Selection.Font.Superscript = False
                wAPP.Selection.Text = pair.Value ''Jaisoft
                If bEtalItalic = True Then
                    Dim ranEtal As Word.Range
                    ranEtal = wAPP.Selection.Range
                    ranEtal.Find.ClearFormatting()
                    ranEtal.Find.Text = "et al"
                    Do While ranEtal.Find.Execute = True
                        ranEtal.Font.Italic = True
                        ranEtal.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Loop
                End If

            Next
            ''////////////////// Remove ref lable /ref number text ///////////////
            Try
                For Each oKV As KeyValuePair(Of Integer, clsRefInfo) In ModRefUtility.dictRefInfo
                    If wAPP.IsObjectValid(oKV.Value.oRefLabelRng) = True And String.IsNullOrEmpty(oKV.Value.oRefLabelRng.Text.Trim()) = False Then
                        Dim oDRng As Word.Range = oKV.Value.oRefLabelRng : Dim s As Short
                        Do While oDRng.Next.Characters.First.Text = " " Or oDRng.Next.Characters.First.Text = vbTab
                            oDRng.SetRange(oDRng.Start, oDRng.End + 1) : s = s + 1
                            Call wAPP.ActiveWindow.ScrollIntoView(oDRng, True)
                            If s = 5 Then Exit Do
                        Loop
                        oKV.Value.oRefLabelRng.Delete()
                    End If
                Next
            Catch ex As Exception
                ex.Data.Clear()
            End Try


            Dim SWR As New StreamWriter(wDoc.Path & "\Vancouver2HavardCitation.html")
            SWR.Write(htmlText)
            SWR.Close()
        End If
        Me.Close()
    End Sub

    Private Sub frmCitationViewer_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'Public dictRefInfo As Dictionary(Of Integer, clsRefInfo)
        'Public dictCitationInfo As Dictionary(Of Word.Range, String)
        Try
            'For Each pair As KeyValuePair(Of Integer, clsRefInfo) In ModRefUtility.dictRefInfo
            '    Dim n As Integer = dgvRef.Rows.Add()
            '    dgvRef.Rows.Item(n).Cells(0).Value = pair.Key.ToString()
            '    dgvRef.Rows.Item(n).Cells(1).Value = pair.Value.ranRef.Text
            'Next
            htmlText = "<html><head>Vancouver to havard citation log </head><body><table>"
            For Each pair As KeyValuePair(Of Word.Range, String) In ModRefUtility.dictConvertedCitation
                Dim n As Integer = dgvCitation.Rows.Add()
                htmlText = htmlText & "<td>"
                htmlText = htmlText & "<tr>" & pair.Key.Text & "</tr>"
                dgvCitation.Rows.Item(n).Cells(0).Value = pair.Key.Text
                dgvCitation.Rows.Item(n).Cells(1).Value = pair.Value
                htmlText = htmlText & "<tr>" & pair.Key.Text & "</tr>"
                htmlText = htmlText & "</td>"
                dgvCitation.Rows.Item(n).ReadOnly = True
            Next
            htmlText = "</table></body></html>"
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class