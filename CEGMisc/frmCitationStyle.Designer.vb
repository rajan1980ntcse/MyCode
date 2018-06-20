<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCitationStyle
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.rtbReference = New System.Windows.Forms.RichTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dgvCitationResult = New System.Windows.Forms.DataGridView()
        Me.btnPrevious = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lbWrongCitation = New System.Windows.Forms.ListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CitationCheck = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.CitationDoc = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CitationWill = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvCitationResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'rtbReference
        '
        Me.rtbReference.Location = New System.Drawing.Point(0, 35)
        Me.rtbReference.Name = "rtbReference"
        Me.rtbReference.Size = New System.Drawing.Size(595, 50)
        Me.rtbReference.TabIndex = 0
        Me.rtbReference.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(-3, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Citation rule:"
        '
        'dgvCitationResult
        '
        Me.dgvCitationResult.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
        Me.dgvCitationResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvCitationResult.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CitationCheck, Me.CitationDoc, Me.CitationWill})
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvCitationResult.DefaultCellStyle = DataGridViewCellStyle1
        Me.dgvCitationResult.Location = New System.Drawing.Point(0, 91)
        Me.dgvCitationResult.Name = "dgvCitationResult"
        Me.dgvCitationResult.Size = New System.Drawing.Size(595, 182)
        Me.dgvCitationResult.TabIndex = 2
        '
        'btnPrevious
        '
        Me.btnPrevious.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious.Location = New System.Drawing.Point(277, 339)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(75, 23)
        Me.btnPrevious.TabIndex = 3
        Me.btnPrevious.Text = "<<"
        Me.btnPrevious.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.Location = New System.Drawing.Point(358, 339)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(75, 23)
        Me.btnNext.TabIndex = 4
        Me.btnNext.Text = ">>"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnOk
        '
        Me.btnOk.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOk.Location = New System.Drawing.Point(439, 339)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(75, 23)
        Me.btnOk.TabIndex = 5
        Me.btnOk.Text = "Ok"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(520, 339)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lbWrongCitation
        '
        Me.lbWrongCitation.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbWrongCitation.FormattingEnabled = True
        Me.lbWrongCitation.Location = New System.Drawing.Point(612, 61)
        Me.lbWrongCitation.Name = "lbWrongCitation"
        Me.lbWrongCitation.Size = New System.Drawing.Size(170, 212)
        Me.lbWrongCitation.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(601, 35)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(181, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Wrong citation:"
        '
        'CitationCheck
        '
        Me.CitationCheck.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None
        Me.CitationCheck.HeaderText = "*"
        Me.CitationCheck.MinimumWidth = 2
        Me.CitationCheck.Name = "CitationCheck"
        Me.CitationCheck.Width = 50
        '
        'CitationDoc
        '
        Me.CitationDoc.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.CitationDoc.HeaderText = "Citation in Document"
        Me.CitationDoc.Name = "CitationDoc"
        '
        'CitationWill
        '
        Me.CitationWill.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.CitationWill.HeaderText = "Citation to be"
        Me.CitationWill.Name = "CitationWill"
        '
        'frmCitationStyle
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(794, 374)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lbWrongCitation)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnPrevious)
        Me.Controls.Add(Me.dgvCitationResult)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.rtbReference)
        Me.Name = "frmCitationStyle"
        Me.Text = "frmCitationStyle"
        CType(Me.dgvCitationResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents rtbReference As RichTextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents dgvCitationResult As DataGridView
    Friend WithEvents btnPrevious As Button
    Friend WithEvents btnNext As Button
    Friend WithEvents btnOk As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents lbWrongCitation As ListBox
    Friend WithEvents Label2 As Label
    Friend WithEvents CitationCheck As DataGridViewCheckBoxColumn
    Friend WithEvents CitationDoc As DataGridViewTextBoxColumn
    Friend WithEvents CitationWill As DataGridViewTextBoxColumn
End Class
