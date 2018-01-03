<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExcelMatcher
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
        Me.ExcelFile1Button = New System.Windows.Forms.Button()
        Me.ExcelFile2Button = New System.Windows.Forms.Button()
        Me.ExcelFile1TextBox = New System.Windows.Forms.TextBox()
        Me.ExcelFile2TextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.MergeButton = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.OpenFileDialog2 = New System.Windows.Forms.OpenFileDialog()
        Me.SheetsExcel1ComboBox = New System.Windows.Forms.ComboBox()
        Me.SheetsExcel2ComboBox = New System.Windows.Forms.ComboBox()
        Me.LookupColumnExcel2ComboBox = New System.Windows.Forms.ComboBox()
        Me.LookupColumnExcel1ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TargetColumnExcel1ComboBox = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.OutputTextBox = New System.Windows.Forms.TextBox()
        Me.OutputButton = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.AddButton = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.RemoveButton = New System.Windows.Forms.Button()
        Me.PrecisionComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ExcelFile1Button
        '
        Me.ExcelFile1Button.Location = New System.Drawing.Point(11, 46)
        Me.ExcelFile1Button.Name = "ExcelFile1Button"
        Me.ExcelFile1Button.Size = New System.Drawing.Size(75, 21)
        Me.ExcelFile1Button.TabIndex = 0
        Me.ExcelFile1Button.Text = "Browse"
        Me.ExcelFile1Button.UseVisualStyleBackColor = True
        '
        'ExcelFile2Button
        '
        Me.ExcelFile2Button.Location = New System.Drawing.Point(318, 46)
        Me.ExcelFile2Button.Name = "ExcelFile2Button"
        Me.ExcelFile2Button.Size = New System.Drawing.Size(75, 21)
        Me.ExcelFile2Button.TabIndex = 1
        Me.ExcelFile2Button.Text = "Browse"
        Me.ExcelFile2Button.UseVisualStyleBackColor = True
        '
        'ExcelFile1TextBox
        '
        Me.ExcelFile1TextBox.Location = New System.Drawing.Point(11, 75)
        Me.ExcelFile1TextBox.Name = "ExcelFile1TextBox"
        Me.ExcelFile1TextBox.ReadOnly = True
        Me.ExcelFile1TextBox.Size = New System.Drawing.Size(289, 20)
        Me.ExcelFile1TextBox.TabIndex = 2
        '
        'ExcelFile2TextBox
        '
        Me.ExcelFile2TextBox.Location = New System.Drawing.Point(318, 75)
        Me.ExcelFile2TextBox.Name = "ExcelFile2TextBox"
        Me.ExcelFile2TextBox.ReadOnly = True
        Me.ExcelFile2TextBox.Size = New System.Drawing.Size(289, 20)
        Me.ExcelFile2TextBox.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 108)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Excel Sheets"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(315, 108)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Excel Sheets"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 160)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(81, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Lookup Column"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(315, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Lookup Column"
        '
        'MergeButton
        '
        Me.MergeButton.Enabled = False
        Me.MergeButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MergeButton.Location = New System.Drawing.Point(222, 379)
        Me.MergeButton.Name = "MergeButton"
        Me.MergeButton.Size = New System.Drawing.Size(191, 59)
        Me.MergeButton.TabIndex = 6
        Me.MergeButton.Text = "Merge"
        Me.MergeButton.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'OpenFileDialog2
        '
        Me.OpenFileDialog2.FileName = "OpenFileDialog2"
        '
        'SheetsExcel1ComboBox
        '
        Me.SheetsExcel1ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.SheetsExcel1ComboBox.FormattingEnabled = True
        Me.SheetsExcel1ComboBox.Location = New System.Drawing.Point(11, 127)
        Me.SheetsExcel1ComboBox.Name = "SheetsExcel1ComboBox"
        Me.SheetsExcel1ComboBox.Size = New System.Drawing.Size(289, 21)
        Me.SheetsExcel1ComboBox.TabIndex = 8
        '
        'SheetsExcel2ComboBox
        '
        Me.SheetsExcel2ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.SheetsExcel2ComboBox.FormattingEnabled = True
        Me.SheetsExcel2ComboBox.Location = New System.Drawing.Point(318, 127)
        Me.SheetsExcel2ComboBox.Name = "SheetsExcel2ComboBox"
        Me.SheetsExcel2ComboBox.Size = New System.Drawing.Size(289, 21)
        Me.SheetsExcel2ComboBox.TabIndex = 8
        '
        'LookupColumnExcel2ComboBox
        '
        Me.LookupColumnExcel2ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LookupColumnExcel2ComboBox.FormattingEnabled = True
        Me.LookupColumnExcel2ComboBox.Location = New System.Drawing.Point(318, 175)
        Me.LookupColumnExcel2ComboBox.Name = "LookupColumnExcel2ComboBox"
        Me.LookupColumnExcel2ComboBox.Size = New System.Drawing.Size(289, 21)
        Me.LookupColumnExcel2ComboBox.TabIndex = 8
        '
        'LookupColumnExcel1ComboBox
        '
        Me.LookupColumnExcel1ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.LookupColumnExcel1ComboBox.FormattingEnabled = True
        Me.LookupColumnExcel1ComboBox.Location = New System.Drawing.Point(11, 176)
        Me.LookupColumnExcel1ComboBox.Name = "LookupColumnExcel1ComboBox"
        Me.LookupColumnExcel1ComboBox.Size = New System.Drawing.Size(289, 21)
        Me.LookupColumnExcel1ComboBox.TabIndex = 8
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(8, 209)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 13)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Target Column"
        '
        'TargetColumnExcel1ComboBox
        '
        Me.TargetColumnExcel1ComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.TargetColumnExcel1ComboBox.FormattingEnabled = True
        Me.TargetColumnExcel1ComboBox.Location = New System.Drawing.Point(11, 225)
        Me.TargetColumnExcel1ComboBox.Name = "TargetColumnExcel1ComboBox"
        Me.TargetColumnExcel1ComboBox.Size = New System.Drawing.Size(226, 21)
        Me.TargetColumnExcel1ComboBox.TabIndex = 8
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 352)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(49, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Export to"
        '
        'OutputTextBox
        '
        Me.OutputTextBox.Location = New System.Drawing.Point(156, 349)
        Me.OutputTextBox.Name = "OutputTextBox"
        Me.OutputTextBox.ReadOnly = True
        Me.OutputTextBox.Size = New System.Drawing.Size(451, 20)
        Me.OutputTextBox.TabIndex = 2
        '
        'OutputButton
        '
        Me.OutputButton.Location = New System.Drawing.Point(66, 347)
        Me.OutputButton.Name = "OutputButton"
        Me.OutputButton.Size = New System.Drawing.Size(75, 21)
        Me.OutputButton.TabIndex = 10
        Me.OutputButton.Text = "Browse"
        Me.OutputButton.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(12, 9)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(106, 31)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Source"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(314, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(162, 31)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "Destination"
        '
        'AddButton
        '
        Me.AddButton.Location = New System.Drawing.Point(245, 225)
        Me.AddButton.Name = "AddButton"
        Me.AddButton.Size = New System.Drawing.Size(55, 23)
        Me.AddButton.TabIndex = 12
        Me.AddButton.Text = "Add"
        Me.AddButton.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(11, 258)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(350, 82)
        Me.ListBox1.TabIndex = 13
        '
        'RemoveButton
        '
        Me.RemoveButton.Location = New System.Drawing.Point(306, 225)
        Me.RemoveButton.Name = "RemoveButton"
        Me.RemoveButton.Size = New System.Drawing.Size(55, 23)
        Me.RemoveButton.TabIndex = 12
        Me.RemoveButton.Text = "Remove"
        Me.RemoveButton.UseVisualStyleBackColor = True
        '
        'PrecisionComboBox1
        '
        Me.PrecisionComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.PrecisionComboBox1.FormattingEnabled = True
        Me.PrecisionComboBox1.Location = New System.Drawing.Point(68, 379)
        Me.PrecisionComboBox1.Name = "PrecisionComboBox1"
        Me.PrecisionComboBox1.Size = New System.Drawing.Size(73, 21)
        Me.PrecisionComboBox1.TabIndex = 15
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(12, 382)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(50, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Precision"
        '
        'ExcelMatcher
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(622, 450)
        Me.Controls.Add(Me.PrecisionComboBox1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.RemoveButton)
        Me.Controls.Add(Me.AddButton)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.OutputButton)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TargetColumnExcel1ComboBox)
        Me.Controls.Add(Me.LookupColumnExcel1ComboBox)
        Me.Controls.Add(Me.LookupColumnExcel2ComboBox)
        Me.Controls.Add(Me.SheetsExcel2ComboBox)
        Me.Controls.Add(Me.SheetsExcel1ComboBox)
        Me.Controls.Add(Me.MergeButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ExcelFile2TextBox)
        Me.Controls.Add(Me.OutputTextBox)
        Me.Controls.Add(Me.ExcelFile1TextBox)
        Me.Controls.Add(Me.ExcelFile2Button)
        Me.Controls.Add(Me.ExcelFile1Button)
        Me.Name = "ExcelMatcher"
        Me.Text = "ExcelMatcher"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ExcelFile1Button As System.Windows.Forms.Button
    Friend WithEvents ExcelFile2Button As System.Windows.Forms.Button
    Friend WithEvents ExcelFile1TextBox As System.Windows.Forms.TextBox
    Friend WithEvents ExcelFile2TextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MergeButton As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents OpenFileDialog2 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SheetsExcel1ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents SheetsExcel2ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents LookupColumnExcel2ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents LookupColumnExcel1ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TargetColumnExcel1ComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents OutputTextBox As System.Windows.Forms.TextBox
    Friend WithEvents OutputButton As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents AddButton As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents RemoveButton As System.Windows.Forms.Button
    Friend WithEvents PrecisionComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
End Class
