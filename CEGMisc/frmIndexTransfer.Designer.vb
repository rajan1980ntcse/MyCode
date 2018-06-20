<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmIndexTransfer
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
        Me.GrpboxMaster = New System.Windows.Forms.GroupBox()
        Me.txtEditeddoc = New System.Windows.Forms.TextBox()
        Me.btnBrowseEditDoc = New System.Windows.Forms.Button()
        Me.GrpboxIndex = New System.Windows.Forms.GroupBox()
        Me.txtIndexdoc = New System.Windows.Forms.TextBox()
        Me.BtnBrowseIndexDoc = New System.Windows.Forms.Button()
        Me.btnProcess = New System.Windows.Forms.Button()
        Me.btncancel = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.tbProgress = New System.Windows.Forms.Label()
        Me.btnReview = New System.Windows.Forms.Button()
        Me.GrpboxMaster.SuspendLayout()
        Me.GrpboxIndex.SuspendLayout()
        Me.SuspendLayout()
        '
        'GrpboxMaster
        '
        Me.GrpboxMaster.Controls.Add(Me.txtEditeddoc)
        Me.GrpboxMaster.Controls.Add(Me.btnBrowseEditDoc)
        Me.GrpboxMaster.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrpboxMaster.Location = New System.Drawing.Point(13, 12)
        Me.GrpboxMaster.Name = "GrpboxMaster"
        Me.GrpboxMaster.Size = New System.Drawing.Size(334, 59)
        Me.GrpboxMaster.TabIndex = 0
        Me.GrpboxMaster.TabStop = False
        Me.GrpboxMaster.Text = "Copy edited manuscript"
        '
        'txtEditeddoc
        '
        Me.txtEditeddoc.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEditeddoc.Location = New System.Drawing.Point(16, 20)
        Me.txtEditeddoc.Name = "txtEditeddoc"
        Me.txtEditeddoc.Size = New System.Drawing.Size(247, 23)
        Me.txtEditeddoc.TabIndex = 1
        '
        'btnBrowseEditDoc
        '
        Me.btnBrowseEditDoc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseEditDoc.Image = Global.CEGMisc.My.Resources.Resources.DirSearch
        Me.btnBrowseEditDoc.Location = New System.Drawing.Point(269, 17)
        Me.btnBrowseEditDoc.Name = "btnBrowseEditDoc"
        Me.btnBrowseEditDoc.Size = New System.Drawing.Size(54, 30)
        Me.btnBrowseEditDoc.TabIndex = 0
        Me.btnBrowseEditDoc.UseVisualStyleBackColor = True
        '
        'GrpboxIndex
        '
        Me.GrpboxIndex.Controls.Add(Me.txtIndexdoc)
        Me.GrpboxIndex.Controls.Add(Me.BtnBrowseIndexDoc)
        Me.GrpboxIndex.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GrpboxIndex.Location = New System.Drawing.Point(353, 12)
        Me.GrpboxIndex.Name = "GrpboxIndex"
        Me.GrpboxIndex.Size = New System.Drawing.Size(343, 59)
        Me.GrpboxIndex.TabIndex = 1
        Me.GrpboxIndex.TabStop = False
        Me.GrpboxIndex.Text = "Indexed manuscript"
        '
        'txtIndexdoc
        '
        Me.txtIndexdoc.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIndexdoc.Location = New System.Drawing.Point(20, 22)
        Me.txtIndexdoc.Name = "txtIndexdoc"
        Me.txtIndexdoc.Size = New System.Drawing.Size(245, 23)
        Me.txtIndexdoc.TabIndex = 3
        '
        'BtnBrowseIndexDoc
        '
        Me.BtnBrowseIndexDoc.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnBrowseIndexDoc.Image = Global.CEGMisc.My.Resources.Resources.DirSearch
        Me.BtnBrowseIndexDoc.Location = New System.Drawing.Point(272, 18)
        Me.BtnBrowseIndexDoc.Name = "BtnBrowseIndexDoc"
        Me.BtnBrowseIndexDoc.Size = New System.Drawing.Size(54, 30)
        Me.BtnBrowseIndexDoc.TabIndex = 2
        Me.BtnBrowseIndexDoc.UseVisualStyleBackColor = True
        '
        'btnProcess
        '
        Me.btnProcess.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnProcess.Location = New System.Drawing.Point(471, 88)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(93, 35)
        Me.btnProcess.TabIndex = 4
        Me.btnProcess.Text = "&Process"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'btncancel
        '
        Me.btncancel.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btncancel.Location = New System.Drawing.Point(586, 88)
        Me.btncancel.Name = "btncancel"
        Me.btncancel.Size = New System.Drawing.Size(93, 35)
        Me.btncancel.TabIndex = 5
        Me.btncancel.Text = "&Exit"
        Me.btncancel.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(29, 98)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(377, 17)
        Me.ProgressBar1.TabIndex = 6
        '
        'tbProgress
        '
        Me.tbProgress.AutoSize = True
        Me.tbProgress.Location = New System.Drawing.Point(412, 102)
        Me.tbProgress.Name = "tbProgress"
        Me.tbProgress.Size = New System.Drawing.Size(39, 13)
        Me.tbProgress.TabIndex = 7
        Me.tbProgress.Text = "Label1"
        '
        'btnReview
        '
        Me.btnReview.Enabled = False
        Me.btnReview.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReview.Location = New System.Drawing.Point(407, 264)
        Me.btnReview.Name = "btnReview"
        Me.btnReview.Size = New System.Drawing.Size(93, 35)
        Me.btnReview.TabIndex = 9
        Me.btnReview.Text = "Review"
        Me.btnReview.UseVisualStyleBackColor = True
        '
        'frmIndexTransfer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(710, 133)
        Me.ControlBox = False
        Me.Controls.Add(Me.btnReview)
        Me.Controls.Add(Me.tbProgress)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.btncancel)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.GrpboxIndex)
        Me.Controls.Add(Me.GrpboxMaster)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "frmIndexTransfer"
        Me.Text = "Index transfer tool v1.2"
        Me.GrpboxMaster.ResumeLayout(False)
        Me.GrpboxMaster.PerformLayout()
        Me.GrpboxIndex.ResumeLayout(False)
        Me.GrpboxIndex.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GrpboxMaster As System.Windows.Forms.GroupBox
    Friend WithEvents txtEditeddoc As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowseEditDoc As System.Windows.Forms.Button
    Friend WithEvents GrpboxIndex As System.Windows.Forms.GroupBox
    Friend WithEvents txtIndexdoc As System.Windows.Forms.TextBox
    Friend WithEvents BtnBrowseIndexDoc As System.Windows.Forms.Button
    Friend WithEvents btnProcess As System.Windows.Forms.Button
    Friend WithEvents btncancel As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents tbProgress As System.Windows.Forms.Label
    Friend WithEvents btnReview As System.Windows.Forms.Button
End Class
