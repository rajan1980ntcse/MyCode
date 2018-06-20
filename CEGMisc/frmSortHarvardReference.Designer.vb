<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSortHarvardReference
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
        Me.dgvRefList = New System.Windows.Forms.DataGridView()
        Me.clmRef = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lvSortedRefList = New System.Windows.Forms.ListView()
        Me.btnSort = New System.Windows.Forms.Button()
        Me.cbSortType = New System.Windows.Forms.ComboBox()
        Me.dgvSortedList = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.dgvRefList, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvSortedList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgvRefList
        '
        Me.dgvRefList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvRefList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvRefList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRefList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.clmRef})
        Me.dgvRefList.Location = New System.Drawing.Point(0, 58)
        Me.dgvRefList.Name = "dgvRefList"
        Me.dgvRefList.Size = New System.Drawing.Size(570, 656)
        Me.dgvRefList.TabIndex = 0
        '
        'clmRef
        '
        Me.clmRef.HeaderText = "Reference List"
        Me.clmRef.Name = "clmRef"
        '
        'lvSortedRefList
        '
        Me.lvSortedRefList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvSortedRefList.GridLines = True
        Me.lvSortedRefList.Location = New System.Drawing.Point(576, 58)
        Me.lvSortedRefList.Name = "lvSortedRefList"
        Me.lvSortedRefList.Size = New System.Drawing.Size(553, 656)
        Me.lvSortedRefList.TabIndex = 1
        Me.lvSortedRefList.UseCompatibleStateImageBehavior = False
        Me.lvSortedRefList.View = System.Windows.Forms.View.List
        '
        'btnSort
        '
        Me.btnSort.Location = New System.Drawing.Point(528, 12)
        Me.btnSort.Name = "btnSort"
        Me.btnSort.Size = New System.Drawing.Size(75, 23)
        Me.btnSort.TabIndex = 2
        Me.btnSort.Text = "Rearrange"
        Me.btnSort.UseVisualStyleBackColor = True
        '
        'cbSortType
        '
        Me.cbSortType.FormattingEnabled = True
        Me.cbSortType.Items.AddRange(New Object() {"order of Surname", "order by character by character", "order by number of authors"})
        Me.cbSortType.Location = New System.Drawing.Point(233, 12)
        Me.cbSortType.Name = "cbSortType"
        Me.cbSortType.Size = New System.Drawing.Size(265, 21)
        Me.cbSortType.TabIndex = 3
        Me.cbSortType.Text = "Select sort type"
        '
        'dgvSortedList
        '
        Me.dgvSortedList.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvSortedList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvSortedList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvSortedList.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1})
        Me.dgvSortedList.Location = New System.Drawing.Point(576, 58)
        Me.dgvSortedList.Name = "dgvSortedList"
        Me.dgvSortedList.Size = New System.Drawing.Size(553, 656)
        Me.dgvSortedList.TabIndex = 4
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.HeaderText = "Sorted Reference List"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        '
        'frmSortHarvardReference
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1128, 717)
        Me.Controls.Add(Me.dgvSortedList)
        Me.Controls.Add(Me.cbSortType)
        Me.Controls.Add(Me.btnSort)
        Me.Controls.Add(Me.lvSortedRefList)
        Me.Controls.Add(Me.dgvRefList)
        Me.Name = "frmSortHarvardReference"
        Me.Text = "Sort Reference"
        CType(Me.dgvRefList, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvSortedList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvRefList As System.Windows.Forms.DataGridView
    Friend WithEvents clmRef As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnSort As System.Windows.Forms.Button
    Friend WithEvents cbSortType As System.Windows.Forms.ComboBox
    Public WithEvents lvSortedRefList As System.Windows.Forms.ListView
    Friend WithEvents dgvSortedList As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
