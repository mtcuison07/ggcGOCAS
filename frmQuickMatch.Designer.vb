<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmQuickMatch
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.dgv1 = New System.Windows.Forms.DataGridView()
        Me.colField00 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colField01 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colField02 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colField03 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colField04 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colField05 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtField00 = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.txtField08 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtField07 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtField06 = New System.Windows.Forms.TextBox()
        Me.Label47 = New System.Windows.Forms.Label()
        Me.Label42 = New System.Windows.Forms.Label()
        Me.txtField05 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.txtField20 = New System.Windows.Forms.TextBox()
        Me.txtField04 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSeeks00 = New System.Windows.Forms.TextBox()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.cmdButton00 = New System.Windows.Forms.Button()
        Me.txtField09 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.dgv1)
        Me.Panel1.Controls.Add(Me.txtField00)
        Me.Panel1.Controls.Add(Me.Label40)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Controls.Add(Me.txtSeeks00)
        Me.Panel1.Location = New System.Drawing.Point(2, 3)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(605, 357)
        Me.Panel1.TabIndex = 0
        '
        'dgv1
        '
        Me.dgv1.AllowUserToAddRows = False
        Me.dgv1.AllowUserToDeleteRows = False
        Me.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgv1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colField00, Me.colField01, Me.colField02, Me.colField03, Me.colField04, Me.colField05})
        Me.dgv1.Location = New System.Drawing.Point(3, 197)
        Me.dgv1.Name = "dgv1"
        Me.dgv1.ReadOnly = True
        Me.dgv1.Size = New System.Drawing.Size(597, 158)
        Me.dgv1.TabIndex = 110
        '
        'colField00
        '
        Me.colField00.HeaderText = "No"
        Me.colField00.Name = "colField00"
        Me.colField00.ReadOnly = True
        Me.colField00.Width = 50
        '
        'colField01
        '
        Me.colField01.HeaderText = "Name"
        Me.colField01.Name = "colField01"
        Me.colField01.ReadOnly = True
        Me.colField01.Width = 103
        '
        'colField02
        '
        Me.colField02.HeaderText = "QM Result"
        Me.colField02.Name = "colField02"
        Me.colField02.ReadOnly = True
        '
        'colField03
        '
        Me.colField03.HeaderText = "Acccount#"
        Me.colField03.Name = "colField03"
        Me.colField03.ReadOnly = True
        '
        'colField04
        '
        Me.colField04.HeaderText = "MCSO Ref#"
        Me.colField04.Name = "colField04"
        Me.colField04.ReadOnly = True
        '
        'colField05
        '
        Me.colField05.HeaderText = "Appl Ref#"
        Me.colField05.Name = "colField05"
        Me.colField05.ReadOnly = True
        '
        'txtField00
        '
        Me.txtField00.BackColor = System.Drawing.SystemColors.Window
        Me.txtField00.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField00.Location = New System.Drawing.Point(97, 4)
        Me.txtField00.Name = "txtField00"
        Me.txtField00.ReadOnly = True
        Me.txtField00.Size = New System.Drawing.Size(126, 20)
        Me.txtField00.TabIndex = 106
        Me.txtField00.TabStop = False
        Me.txtField00.Text = "M00000000001"
        '
        'Label40
        '
        Me.Label40.AutoSize = True
        Me.Label40.Location = New System.Drawing.Point(5, 6)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(83, 13)
        Me.Label40.TabIndex = 108
        Me.Label40.Text = "Transaction No."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtField09)
        Me.GroupBox1.Controls.Add(Me.Panel3)
        Me.GroupBox1.Controls.Add(Me.txtField07)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtField06)
        Me.GroupBox1.Controls.Add(Me.Label47)
        Me.GroupBox1.Controls.Add(Me.Label42)
        Me.GroupBox1.Controls.Add(Me.txtField05)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.txtField20)
        Me.GroupBox1.Controls.Add(Me.txtField04)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Location = New System.Drawing.Point(5, 36)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(595, 154)
        Me.GroupBox1.TabIndex = 107
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Applicant Info"
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.txtField08)
        Me.Panel3.Controls.Add(Me.Label3)
        Me.Panel3.Location = New System.Drawing.Point(343, 62)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(246, 84)
        Me.Panel3.TabIndex = 117
        '
        'txtField08
        '
        Me.txtField08.BackColor = System.Drawing.SystemColors.Window
        Me.txtField08.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField08.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtField08.Location = New System.Drawing.Point(5, 42)
        Me.txtField08.Multiline = True
        Me.txtField08.Name = "txtField08"
        Me.txtField08.ReadOnly = True
        Me.txtField08.Size = New System.Drawing.Size(236, 38)
        Me.txtField08.TabIndex = 115
        Me.txtField08.TabStop = False
        Me.txtField08.Text = "AP36-Ex-07-P"
        Me.txtField08.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(39, 5)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(151, 31)
        Me.Label3.TabIndex = 116
        Me.Label3.Text = "QM Result"
        '
        'txtField07
        '
        Me.txtField07.BackColor = System.Drawing.SystemColors.Window
        Me.txtField07.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField07.Location = New System.Drawing.Point(91, 104)
        Me.txtField07.Multiline = True
        Me.txtField07.Name = "txtField07"
        Me.txtField07.ReadOnly = True
        Me.txtField07.Size = New System.Drawing.Size(243, 41)
        Me.txtField07.TabIndex = 116
        Me.txtField07.TabStop = False
        Me.txtField07.Text = "Alaminos"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 13)
        Me.Label2.TabIndex = 115
        Me.Label2.Text = "Address:"
        '
        'txtField06
        '
        Me.txtField06.BackColor = System.Drawing.SystemColors.Window
        Me.txtField06.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField06.Location = New System.Drawing.Point(91, 81)
        Me.txtField06.Name = "txtField06"
        Me.txtField06.ReadOnly = True
        Me.txtField06.Size = New System.Drawing.Size(243, 20)
        Me.txtField06.TabIndex = 114
        Me.txtField06.TabStop = False
        Me.txtField06.Text = "Baracao, Kenneth"
        '
        'Label47
        '
        Me.Label47.AutoSize = True
        Me.Label47.Location = New System.Drawing.Point(14, 84)
        Me.Label47.Name = "Label47"
        Me.Label47.Size = New System.Drawing.Size(46, 13)
        Me.Label47.TabIndex = 113
        Me.Label47.Text = "Spouse:"
        '
        'Label42
        '
        Me.Label42.AutoSize = True
        Me.Label42.Location = New System.Drawing.Point(14, 39)
        Me.Label42.Name = "Label42"
        Me.Label42.Size = New System.Drawing.Size(45, 13)
        Me.Label42.TabIndex = 109
        Me.Label42.Text = "Address"
        '
        'txtField05
        '
        Me.txtField05.BackColor = System.Drawing.SystemColors.Window
        Me.txtField05.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField05.Location = New System.Drawing.Point(431, 15)
        Me.txtField05.Name = "txtField05"
        Me.txtField05.ReadOnly = True
        Me.txtField05.Size = New System.Drawing.Size(158, 20)
        Me.txtField05.TabIndex = 16
        Me.txtField05.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(340, 17)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(88, 13)
        Me.Label11.TabIndex = 64
        Me.Label11.Text = "Application Date:"
        '
        'txtField20
        '
        Me.txtField20.BackColor = System.Drawing.SystemColors.Window
        Me.txtField20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField20.Location = New System.Drawing.Point(92, 37)
        Me.txtField20.Multiline = True
        Me.txtField20.Name = "txtField20"
        Me.txtField20.ReadOnly = True
        Me.txtField20.Size = New System.Drawing.Size(243, 41)
        Me.txtField20.TabIndex = 8
        Me.txtField20.TabStop = False
        Me.txtField20.Text = "Alaminos"
        '
        'txtField04
        '
        Me.txtField04.BackColor = System.Drawing.SystemColors.Window
        Me.txtField04.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField04.Location = New System.Drawing.Point(92, 14)
        Me.txtField04.Name = "txtField04"
        Me.txtField04.ReadOnly = True
        Me.txtField04.Size = New System.Drawing.Size(243, 20)
        Me.txtField04.TabIndex = 7
        Me.txtField04.TabStop = False
        Me.txtField04.Text = "Baracao, Kenneth"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(14, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 92
        Me.Label1.Text = "Client Name"
        '
        'txtSeeks00
        '
        Me.txtSeeks00.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.txtSeeks00.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSeeks00.Location = New System.Drawing.Point(104, 6)
        Me.txtSeeks00.Name = "txtSeeks00"
        Me.txtSeeks00.ReadOnly = True
        Me.txtSeeks00.Size = New System.Drawing.Size(126, 20)
        Me.txtSeeks00.TabIndex = 109
        Me.txtSeeks00.TabStop = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.cmdButton00)
        Me.Panel2.Location = New System.Drawing.Point(611, 3)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(94, 357)
        Me.Panel2.TabIndex = 1
        '
        'cmdButton00
        '
        Me.cmdButton00.Location = New System.Drawing.Point(1, 3)
        Me.cmdButton00.Name = "cmdButton00"
        Me.cmdButton00.Size = New System.Drawing.Size(90, 27)
        Me.cmdButton00.TabIndex = 2
        Me.cmdButton00.Text = "Ok"
        Me.cmdButton00.UseVisualStyleBackColor = True
        '
        'txtField09
        '
        Me.txtField09.BackColor = System.Drawing.SystemColors.Window
        Me.txtField09.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtField09.Location = New System.Drawing.Point(431, 36)
        Me.txtField09.Name = "txtField09"
        Me.txtField09.ReadOnly = True
        Me.txtField09.Size = New System.Drawing.Size(158, 20)
        Me.txtField09.TabIndex = 118
        Me.txtField09.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(341, 39)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(79, 13)
        Me.Label4.TabIndex = 119
        Me.Label4.Text = "Application No:"
        '
        'frmQuickMatch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(708, 361)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmQuickMatch"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "QM Result"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.dgv1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtField00 As System.Windows.Forms.TextBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtField06 As System.Windows.Forms.TextBox
    Friend WithEvents Label47 As System.Windows.Forms.Label
    Friend WithEvents Label42 As System.Windows.Forms.Label
    Friend WithEvents txtField05 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtField20 As System.Windows.Forms.TextBox
    Friend WithEvents txtField04 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSeeks00 As System.Windows.Forms.TextBox
    Friend WithEvents txtField07 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dgv1 As System.Windows.Forms.DataGridView
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents cmdButton00 As System.Windows.Forms.Button
    Friend WithEvents colField00 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colField01 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colField02 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colField03 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colField04 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colField05 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtField08 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtField09 As System.Windows.Forms.TextBox
End Class
