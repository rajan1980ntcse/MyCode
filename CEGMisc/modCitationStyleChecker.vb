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
Module modCitationStyleChecker
    Public oActApp As Word.Application
    Public Function ReferenceCitationStyleChecking(wDoc As Word.Document, WAPP As Word.Application, pName As String, jName As String, sConfigPath As String) As Boolean
        oActApp = WAPP
        Dim objfrmCitationStyle As New frmCitationStyle(WAPP, pName, jName, sConfigPath)
        objfrmCitationStyle.Show()
        'If CollectRefInfo(wDoc, "†Reference") = False Then
        '    MessageBox.Show("ERROR : Unable to collect Reference information ")
        '    Return False
        'End If
    End Function


End Module
