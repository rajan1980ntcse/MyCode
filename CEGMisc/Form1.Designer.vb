﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TEST = New System.Windows.Forms.Button()
        Me.btnIndex = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.btnFtnFWord = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(56, 26)
        Me.Button1.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 28)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(56, 62)
        Me.Button2.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(100, 28)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Browse"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TEST
        '
        Me.TEST.Location = New System.Drawing.Point(56, 97)
        Me.TEST.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.TEST.Name = "TEST"
        Me.TEST.Size = New System.Drawing.Size(100, 28)
        Me.TEST.TabIndex = 2
        Me.TEST.Text = "TEST"
        Me.TEST.UseVisualStyleBackColor = True
        '
        'btnIndex
        '
        Me.btnIndex.Location = New System.Drawing.Point(56, 205)
        Me.btnIndex.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnIndex.Name = "btnIndex"
        Me.btnIndex.Size = New System.Drawing.Size(100, 28)
        Me.btnIndex.TabIndex = 3
        Me.btnIndex.Text = "IndexTest"
        Me.btnIndex.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(56, 148)
        Me.Button3.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 28)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Van2Har"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'btnFtnFWord
        '
        Me.btnFtnFWord.Location = New System.Drawing.Point(199, 45)
        Me.btnFtnFWord.Margin = New System.Windows.Forms.Padding(4)
        Me.btnFtnFWord.Name = "btnFtnFWord"
        Me.btnFtnFWord.Size = New System.Drawing.Size(150, 28)
        Me.btnFtnFWord.TabIndex = 5
        Me.btnFtnFWord.Text = "FootnoteFirstWord"
        Me.btnFtnFWord.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(476, 337)
        Me.Controls.Add(Me.btnFtnFWord)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.btnIndex)
        Me.Controls.Add(Me.TEST)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TEST As System.Windows.Forms.Button
    Friend WithEvents btnIndex As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnFtnFWord As System.Windows.Forms.Button

End Class
