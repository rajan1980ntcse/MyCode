Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq


Public Class clsCEGMiscWML


    Public Function ToGetUsedStyleLog(ByVal sDocxFullName As String, sProcessPath As String, sPubName As String, sISBN As String, sBookTitle As String) As Boolean
        Try
            Dim sReportHTML As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUPBITS styles report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'><!-- Report title --></h1><table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!--ISBN--></td></tr><tr><td><b>Title</b></td><td>: <!--Title--></td></tr><tr><td><b>Publisher</b></td><td>: <!--Pub--></td></tr><tr><td><b>Date and time</b></td><td>: <!--Time--></td></tr></tbody></table><br/><hr/></head><br/><body>" &
                                          "<table width='95%' border='1' align='center'>" &
                                           "<tr bgcolor='#666666'><td width='25%'><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Style name</b></font></td><td><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Document name(s)</b></font></td></tr>"
            Dim oDocList As String() = sDocxFullName.Replace("||", "|").Split(New Char() {"|"}).ToArray()
            Call Array.Sort(oDocList)

            Dim oDocNameAndStyle As New SortedDictionary(Of String, SortedDictionary(Of String, Integer))
            Dim oAllStyles As New SortedDictionary(Of String, String)
            Dim oclsCEGWML As New ClsCEGWML


            Dim sMsg As String = String.Empty
            For Each sDocName As String In oDocList
                Dim sDocFullName As String = Path.Combine(sProcessPath, sDocName)
                If File.Exists(sDocFullName) = True AndAlso oDocNameAndStyle.ContainsKey(sDocName) = False Then
                    Dim oclsDoc As clsDocInfo = oclsCEGWML.ToGetDocXUsedStyleList(sDocFullName, sMsg)
                    Call oDocNameAndStyle.Add(sDocName, oclsDoc.oParaUsedStyleList)
                End If
            Next


            sReportHTML = sReportHTML.Replace("<!-- Report title -->", "CUPBITS styles report")
            sReportHTML = sReportHTML.Replace("<!--ISBN-->", sISBN) : sReportHTML = sReportHTML.Replace("<!--Title-->", sBookTitle)
            sReportHTML = sReportHTML.Replace("<!--Pub-->", sPubName) : sReportHTML = sReportHTML.Replace("<!--Time-->", Now)

            Dim sRowHTML As String = String.Empty
            For Each oKV As KeyValuePair(Of String, String) In oAllStyles
                sRowHTML = sRowHTML & "<tr><td>" & oKV.Key & "</td><td>" & oKV.Value & "</td></tr>"
            Next
            MessageBox.Show(sProcessPath + "--" + sISBN)
            Dim oSW As New StreamWriter(Path.Combine(sProcessPath, sISBN & "_Style_Report1.html"))
            oSW.Write(sReportHTML & sRowHTML & "</table></body></html>") : oSW.Flush() : oSW.Close()

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "CEGenius - " & Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Function
    Public Function ToGetUsedStyleForEditingFramework(ByVal sDocxFullName As String, sProcessPath As String, sPubName As String, sISBN As String, sBookTitle As String) As Boolean
        Try
            Dim sReportHTML As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUPBITS styles report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'><!-- Report title --></h1><table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!--ISBN--></td></tr><tr><td><b>Title</b></td><td>: <!--Title--></td></tr><tr><td><b>Publisher</b></td><td>: <!--Pub--></td></tr><tr><td><b>Date and time</b></td><td>: <!--Time--></td></tr></tbody></table><br/><hr/></head><br/><body>" &
                                          "<table width='95%' border='1' align='center'>" &
                                           "<tr bgcolor='#666666'><td width='25%'><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Style name</b></font></td></tr>"
            Dim oDocList As String() = sDocxFullName.Replace("||", "|").Split(New Char() {"|"}).ToArray()
            Call Array.Sort(oDocList)

            Dim oDocNameAndStyle As New SortedDictionary(Of String, SortedDictionary(Of String, Integer))
            Dim dictStyle As New Dictionary(Of String, clsDocInfo)
            Dim oAllStyles As New SortedDictionary(Of String, String)
            Dim oclsCEGWML As New clsCEGWML


            Dim sMsg As String = String.Empty
            For Each sDocName As String In oDocList
                Dim sDocFullName As String = Path.Combine(sProcessPath, sDocName)
                If File.Exists(sDocFullName) = True AndAlso oDocNameAndStyle.ContainsKey(sDocName) = False Then
                    ''Dim oDocUsedStyle As SortedDictionary(Of String, Integer)
                    '' MessageBox.Show(sDocFullName)
                    Dim oclsDoc As clsDocInfo = oclsCEGWML.ToGetDocXUsedStyleList(sDocFullName, sMsg)
                    Call oDocNameAndStyle.Add(sDocName, oclsDoc.oParaUsedStyleList)
                    dictStyle.Add(sDocName, oclsDoc)
                End If
            Next

            Dim miscINI As String : Dim pStyleName As String : Dim cStyleName As String
            miscINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New CEGINI.clsINI(miscINI)

            If (sPubName.ToUpper() = "WKD") Then
                pStyleName = oReadINI.INIReadValue("WKD Journals", "ParagraphStyle")
                cStyleName = oReadINI.INIReadValue("WKD Journals", "CharacterSytle")
            End If

            sReportHTML = sReportHTML.Replace("<!-- Report title -->", sPubName & " styles report")
            sReportHTML = sReportHTML.Replace("<!--ISBN-->", sISBN) : sReportHTML = sReportHTML.Replace("<!--Title-->", sBookTitle)
            sReportHTML = sReportHTML.Replace("<!--Pub-->", sPubName) : sReportHTML = sReportHTML.Replace("<!--Time-->", Now)

            Dim sRowHTML As String = String.Empty
            For Each oKV As KeyValuePair(Of String, clsDocInfo) In dictStyle
                sRowHTML = sRowHTML & "<tr bgcolor='#555555'><td>" & oKV.Key & "</td></tr>"

                sRowHTML = sRowHTML & "<tr><td><b>Paragraphs Style</b></td></tr>"
                For Each sPara As KeyValuePair(Of String, Integer) In oKV.Value.oParaUsedStyleList
                    If Not (InStr(1, pStyleName.ToLower(), "||" & sPara.Key.ToLower() & "||") > 0) Then
                        sRowHTML = sRowHTML & "<tr><td>" + sPara.Key + "</td></tr>"
                    End If
                Next
                sRowHTML = sRowHTML & "<tr><td><b>Character Style</b></td></tr>"
                For Each sPara As KeyValuePair(Of String, Integer) In oKV.Value.oCharUsedStyleList
                    If Not (InStr(1, cStyleName.ToLower(), "||" & sPara.Key.ToLower() & "||") > 0) Then
                        sRowHTML = sRowHTML & "<tr><td>" + sPara.Key + "</td></tr>"
                    End If
                Next
            Next
            Dim oSW As New StreamWriter(Path.Combine(sProcessPath, sISBN & "_Style_Report1.html"))
            oSW.Write(sReportHTML & sRowHTML & "</table></body></html>") : oSW.Flush() : oSW.Close()
            ToGetUsedStyleForEditingFramework = True
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "CEGenius - " & Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Function
    Public Function ToGetUsedStyleForWKD(ByVal sDocxFullName As String, sProcessPath As String, sPubName As String, sISBN As String, sBookTitle As String) As Boolean
        Try
            Dim sReportHTML As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>" + sPubName + " styles report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'><!-- Report title --></h1><table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!--ISBN--></td></tr><tr><td><b>Title</b></td><td>: <!--Title--></td></tr><tr><td><b>Publisher</b></td><td>: <!--Pub--></td></tr><tr><td><b>Date and time</b></td><td>: <!--Time--></td></tr></tbody></table><br/><hr/></head><br/><body>" &
                                          "<table width='95%' border='1' align='center'>" &
                                           "<tr bgcolor='#666666'><td width='25%'><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Style name</b></font></td></tr>"
            Dim oDocList As String() = sDocxFullName.Replace("||", "|").Split(New Char() {"|"}).ToArray()
            Call Array.Sort(oDocList)

            Dim oDocNameAndStyle As New SortedDictionary(Of String, SortedDictionary(Of String, Integer))
            Dim dictStyle As New Dictionary(Of String, clsDocInfo)
            Dim oAllStyles As New SortedDictionary(Of String, String)
            Dim oclsCEGWML As New clsCEGWML


            Dim sMsg As String = String.Empty
            For Each sDocName As String In oDocList
                Dim sDocFullName As String = Path.Combine(sProcessPath, sDocName)
                If File.Exists(sDocFullName) = True AndAlso oDocNameAndStyle.ContainsKey(sDocName) = False Then
                    ''Dim oDocUsedStyle As SortedDictionary(Of String, Integer)
                    '' MessageBox.Show(sDocFullName)
                    Dim oclsDoc As clsDocInfo = oclsCEGWML.ToGetDocXUsedStyleList(sDocFullName, sMsg)
                    Call oDocNameAndStyle.Add(sDocName, oclsDoc.oParaUsedStyleList)
                    dictStyle.Add(sDocName, oclsDoc)
                End If
            Next

            Dim miscINI As String : Dim pStyleName As String : Dim cStyleName As String
            miscINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New CEGINI.clsINI(miscINI)

            If (sPubName.ToUpper() = "WKD") Then
                pStyleName = oReadINI.INIReadValue("WKD Journals", "ParagraphStyle")
                cStyleName = oReadINI.INIReadValue("WKD Journals", "CharacterSytle")
            End If

            sReportHTML = sReportHTML.Replace("<!-- Report title -->", sPubName & " styles report")
            sReportHTML = sReportHTML.Replace("<!--ISBN-->", sISBN) : sReportHTML = sReportHTML.Replace("<!--Title-->", sBookTitle)
            sReportHTML = sReportHTML.Replace("<!--Pub-->", sPubName) : sReportHTML = sReportHTML.Replace("<!--Time-->", Now)

            Dim sRowHTML As String = String.Empty
            For Each oKV As KeyValuePair(Of String, clsDocInfo) In dictStyle
                sRowHTML = sRowHTML & "<tr bgcolor='#555555'><td>" & oKV.Key & "</td></tr>"

                sRowHTML = sRowHTML & "<tr><td><b>Paragraphs Style</b></td></tr>"
                For Each sPara As KeyValuePair(Of String, Integer) In oKV.Value.oParaUsedStyleList
                    If Not (InStr(1, pStyleName.ToLower(), "||" & sPara.Key.ToLower() & "||") > 0) Then
                        sRowHTML = sRowHTML & "<tr><td>" + sPara.Key + "#########" + sPara.Value + "</td></tr>"
                    End If
                Next
                sRowHTML = sRowHTML & "<tr><td><b>Character Style</b></td></tr>"
                For Each sPara As KeyValuePair(Of String, Integer) In oKV.Value.oCharUsedStyleList
                    If Not (InStr(1, cStyleName.ToLower(), "||" & sPara.Key.ToLower() & "||") > 0) Then
                        sRowHTML = sRowHTML & "<tr><td>" + sPara.Key + "#########" + sPara.Value + "</td></tr>"
                    End If
                Next
            Next
            Dim oSW As New StreamWriter(Path.Combine(sProcessPath, sISBN & "_Style_Report1.html"))
            oSW.Write(sReportHTML & sRowHTML & "</table></body></html>") : oSW.Flush() : oSW.Close()
            ToGetUsedStyleForWKD = True
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "CEGenius - " & Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Function

    Public Function ToGetUsedStyle4DocName(ByVal sDocxFullName As String, sProcessPath As String, sPubName As String, sISBN As String, sBookTitle As String) As Boolean
        Try
            Dim sReportHTML As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUPBITS styles report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'><!-- Report title --></h1><table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!--ISBN--></td></tr><tr><td><b>Title</b></td><td>: <!--Title--></td></tr><tr><td><b>Publisher</b></td><td>: <!--Pub--></td></tr><tr><td><b>Date and time</b></td><td>: <!--Time--></td></tr></tbody></table><br/><hr/></head><br/><body>" &
                                          "<table width='95%' border='1' align='center'>" &
                                           "<tr bgcolor='#666666'><td width='25%'><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Style name</b></font></td><td><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Document name(s)</b></font></td></tr>"
            Dim oDocList As String() = sDocxFullName.Replace("||", "|").Split(New Char() {"|"}).ToArray()
            Call Array.Sort(oDocList)

            Dim oDocNameAndStyle As New SortedDictionary(Of String, SortedDictionary(Of String, Integer))
            Dim oAllStyles As New SortedDictionary(Of String, String)
            Dim oclsCEGWML As New ClsCEGWML


            Dim sMsg As String = String.Empty
            For Each sDocName As String In oDocList
                Dim sDocFullName As String = Path.Combine(sProcessPath, sDocName)
                If File.Exists(sDocFullName) = True AndAlso oDocNameAndStyle.ContainsKey(sDocName) = False Then
                    ''Dim oDocUsedStyle As SortedDictionary(Of String, Integer)
                    Dim oclsDoc As clsDocInfo = oclsCEGWML.ToGetDocXUsedStyleList(sDocFullName, sMsg)
                    Call oDocNameAndStyle.Add(sDocName, oclsDoc.oParaUsedStyleList)
                End If
            Next


            sReportHTML = sReportHTML.Replace("<!-- Report title -->", "CUPBITS styles report")
            sReportHTML = sReportHTML.Replace("<!--ISBN-->", sISBN) : sReportHTML = sReportHTML.Replace("<!--Title-->", sBookTitle)
            sReportHTML = sReportHTML.Replace("<!--Pub-->", sPubName) : sReportHTML = sReportHTML.Replace("<!--Time-->", Now)

            Dim sRowHTML As String = String.Empty
            For Each oKV As KeyValuePair(Of String, String) In oAllStyles
                sRowHTML = sRowHTML & "<tr><td>" & oKV.Key & "</td><td>" & oKV.Value & "</td></tr>"
            Next
            Dim oSW As New StreamWriter(Path.Combine(sProcessPath, sISBN & "_Style_Report1.html"))
            oSW.Write(sReportHTML & sRowHTML & "</table></body></html>") : oSW.Flush() : oSW.Close()

        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "CEGenius - " & Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Function



    Public Function ToGetUsedStyle(ByVal sDocxFullName As String, sProcessPath As String, sPubName As String, sISBN As String, sBookTitle As String) As Boolean
        Try
            Dim sReportHTML As String = "<html><META http-equiv='Content-Type' content='text/html; charset=utf-8'/><title>CUP style report</title><style>head, body, table { font-family: Verdana, Helvetica, Arial, sans-serif; } th {background-color: gray;font-size:14;font-weight:bold;color:white}</style><head><br/><h1 align='right'><a href='http://www.newgenediting.com'><img src='http://www.newgenediting.com/images/logo_small.png' width='204' height='25' alt='Newgen Editing'></a></h1><h1 align='center'><!-- Report title --></h1><table border='0' align='center'><tbody><tr><td><b>ISBN</b></td><td>: <!--ISBN--></td></tr><tr><td><b>Title</b></td><td>: <!--Title--></td></tr><tr><td><b>Publisher</b></td><td>: <!--Pub--></td></tr><tr><td><b>Date and time</b></td><td>: <!--Time--></td></tr></tbody></table><br/><hr/></head><br/><body>" &
                                          "<table width='95%' border='1' align='center'><tr bgcolor='#666666'><td align='center' colspan='9'><font style='font-family:Verdana;font-size:14' color='#FFFFFF'><b>Style report</b></font></td></tr><tr><td width='15%'><b>File name</b></td><td colspan='8' align='center'><b>Style name</b></td></tr><tr><th colspan='9'></th></tr>" &
                                          "<tr><td width='15%' rowspan='2'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td></tr><tr><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td><td width='10%'></td></tr>"
            Dim oDocList As String() = sDocxFullName.Replace("||", "|").Split(New Char() {"|"}).ToArray()
            Call Array.Sort(oDocList) : Dim oclsCEGWML As New ClsCEGWML

            Dim oDocNameAndStyle As New SortedDictionary(Of String, SortedDictionary(Of String, Integer))
            Dim oAllStyles As New SortedDictionary(Of String, String)
            Dim sMsg As String = String.Empty
            For Each sDocName As String In oDocList
                Dim sDocFullName As String = Path.Combine(sProcessPath, sDocName)
                If File.Exists(sDocFullName) = True AndAlso oDocNameAndStyle.ContainsKey(sDocName) = False Then
                    'Dim oDocUsedStyle As SortedDictionary(Of String, Integer)
                    Dim oclsDoc As clsDocInfo = oclsCEGWML.ToGetDocXUsedStyleList(sDocFullName, sMsg)
                    Call oDocNameAndStyle.Add(sDocName, oclsDoc.oParaUsedStyleList)
                End If
            Next

            sReportHTML = sReportHTML.Replace("<!-- Report title -->", sISBN & " - Stylereport")
            sReportHTML = sReportHTML.Replace("<!--ISBN-->", sISBN) : sReportHTML = sReportHTML.Replace("<!--Title-->", sBookTitle)
            sReportHTML = sReportHTML.Replace("<!--Pub-->", sPubName) : sReportHTML = sReportHTML.Replace("<!--Time-->", Now)
            Dim sDocHTML As String = String.Empty : Dim sRowHTML As String = String.Empty : Dim sDataHTML As String = String.Empty
            For Each oKV As KeyValuePair(Of String, SortedDictionary(Of String, Integer)) In oDocNameAndStyle
                Dim oStyleUsed As SortedDictionary(Of String, Integer) = oKV.Value
                Dim R As Integer = 0 : Dim X As Integer = 0 : Dim iSpanRow As Integer = Math.DivRem(oStyleUsed.Count, 8, R)
                If R > 0 Then iSpanRow = iSpanRow + 1
                'If iSpanRow > 0 Then iSpanRow = iSpanRow - 1
                sRowHTML = String.Empty : sDataHTML = String.Empty
                For Each oSKV As KeyValuePair(Of String, Integer) In oStyleUsed
                    X = X + 1 : R = 0 : Math.DivRem(X, 8, R)
                    If X = 8 Then
                        If iSpanRow > 1 Then
                            sDataHTML = sDataHTML & "<td>" & oSKV.Key & " (" & oSKV.Value & ")</td>"
                            sRowHTML = sRowHTML & "<tr><td rowspan='" & iSpanRow & "'>" & oKV.Key & "</td>" & sDataHTML & "</tr>"
                            sDataHTML = String.Empty
                        Else
                            sDataHTML = sDataHTML & "<td>" & oSKV.Key & " (" & oSKV.Value & ")</td>"
                            sRowHTML = sRowHTML & "<tr><td>" & oKV.Key & "</td>" & sDataHTML & "</tr>"
                        End If
                    ElseIf X = oStyleUsed.Count Then
                        sDataHTML = sDataHTML & "<td>" & oSKV.Key & " (" & oSKV.Value & ")</td>"
                        For X = R To 7 : sDataHTML = sDataHTML & "<td>&nbsp;</td>" : Next
                        If iSpanRow <= 1 Then
                            sRowHTML = sRowHTML & "<tr><td>" & oKV.Key & "</td>" & sDataHTML & "</tr>" : sDataHTML = String.Empty
                        Else
                            sRowHTML = sRowHTML & "<tr>" & sDataHTML & "</tr>" : sDataHTML = String.Empty
                        End If
                    ElseIf R > 0 Then
                        sDataHTML = sDataHTML & "<td>" & oSKV.Key & " (" & oSKV.Value & ")</td>"
                    Else
                        sDataHTML = "<tr>" & sDataHTML & "<td>" & oSKV.Key & " (" & oSKV.Value & ")</td></tr>"
                        sRowHTML = sRowHTML & sDataHTML : sDataHTML = String.Empty
                    End If
                Next
                sDocHTML = sDocHTML & sRowHTML
            Next
            Dim oSW As New StreamWriter(Path.Combine(sProcessPath, sISBN & "_Style_Report.html"))
            oSW.Write(sReportHTML & sDocHTML & "</table></body></html>") : oSW.Flush() : oSW.Close()


        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "CEGenius - " & Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Function

End Class
