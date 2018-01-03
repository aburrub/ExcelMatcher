Imports DevExpress.Spreadsheet

Public Class ExcelMatcher

    Private Sub ExcelMatcher_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Fixed3D

        Dim d As New Dictionary(Of String, Integer)
        For i As Integer = 60 To 100
            d.Add(String.Format("{0}%", i.ToString), i)
        Next

        Me.PrecisionComboBox1.DataSource = New BindingSource(d, Nothing)
        Me.PrecisionComboBox1.ValueMember = "Value"
        Me.PrecisionComboBox1.DisplayMember = "Key"

        Me.PrecisionComboBox1.SelectedValue = 100
    End Sub

    Private Sub ExcelFile1Button_Click(sender As Object, e As EventArgs) Handles ExcelFile1Button.Click

        Me.OpenFileDialog1.Filter = "Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls"
        Me.OpenFileDialog1.FileName = ""
        Me.OpenFileDialog1.ShowDialog()
        Me.ExcelFile1TextBox.Text = Me.OpenFileDialog1.FileName

        Dim MySheets As Dictionary(Of String, MySheet) = Nothing
        AppModule.ReadExcelFile(Me.ExcelFile1TextBox.Text, MySheets)
        If MySheets IsNot Nothing AndAlso MySheets.Keys.Count > 0 Then
            AppModule.BindSheetsToComboBoxControl(MySheets, Me.SheetsExcel1ComboBox)
            AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel1ComboBox.SelectedValue, Me.LookupColumnExcel1ComboBox)
            AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel1ComboBox.SelectedValue, Me.TargetColumnExcel1ComboBox)
        End If

        ' by default choose the same file for the second file
        If Me.ExcelFile2TextBox.Text = "" Then
            Me.ExcelFile2TextBox.Text = Me.OpenFileDialog1.FileName
            MySheets = Nothing
            AppModule.ReadExcelFile(Me.ExcelFile1TextBox.Text, MySheets)
            If MySheets IsNot Nothing AndAlso MySheets.Keys.Count > 0 Then
                AppModule.BindSheetsToComboBoxControl(MySheets, Me.SheetsExcel2ComboBox)
                AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel2ComboBox.SelectedValue, Me.LookupColumnExcel2ComboBox)
            End If
        End If
     

        setMergeEnableFlag()

    
    End Sub
   
    Private Sub ExcelFile2Button_Click(sender As Object, e As EventArgs) Handles ExcelFile2Button.Click
        Me.OpenFileDialog2.Filter = "Excel Files(.xlsx)|*.xlsx|Excel Files(.xls)|*.xls"
        Me.OpenFileDialog2.FileName = ""
        Me.OpenFileDialog2.ShowDialog()
        Me.ExcelFile2TextBox.Text = Me.OpenFileDialog2.FileName

        Dim MySheets As Dictionary(Of String, MySheet) = Nothing
        AppModule.ReadExcelFile(Me.ExcelFile2TextBox.Text, MySheets)
        If MySheets IsNot Nothing AndAlso MySheets.Keys.Count > 0 Then
            AppModule.BindSheetsToComboBoxControl(MySheets, Me.SheetsExcel2ComboBox)
            AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel2ComboBox.SelectedValue, Me.LookupColumnExcel2ComboBox)
        End If

        setMergeEnableFlag()
    End Sub

    Private Sub SheetsExcel1ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SheetsExcel1ComboBox.SelectedIndexChanged
        AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel1ComboBox.SelectedValue, Me.LookupColumnExcel1ComboBox)
        AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel1ComboBox.SelectedValue, Me.TargetColumnExcel1ComboBox)
        setMergeEnableFlag()
    End Sub

    Private Sub SheetsExcel2ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SheetsExcel2ComboBox.SelectedIndexChanged
        AppModule.BindSheetCoumnsToComboBoxControl(Me.SheetsExcel2ComboBox.SelectedValue, Me.LookupColumnExcel2ComboBox)
        setMergeEnableFlag()
    End Sub



    Private Sub setMergeEnableFlag()
        Dim f As Boolean = True
        If Me.SheetsExcel1ComboBox.SelectedIndex >= 0 And Me.LookupColumnExcel1ComboBox.SelectedIndex >= 0 And Me.TargetColumnExcel1ComboBox.SelectedIndex >= 0 Then
            f = f And True
        Else
            f = f And False
        End If

        If Me.SheetsExcel2ComboBox.SelectedIndex >= 0 And Me.LookupColumnExcel2ComboBox.SelectedIndex >= 0 Then
            f = f And True
        Else
            f = f And False
        End If

        Me.MergeButton.Enabled = f
    End Sub

    Private Sub LookupColumnExcel1ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LookupColumnExcel1ComboBox.SelectedIndexChanged
        Dim tmp As String = Nothing
        For Each itm As String In Me.TargetColumns
            If Me.LookupColumnExcel1ComboBox.Text = itm Then
                tmp = itm               
            End If
        Next
        If tmp IsNot Nothing Then
            Me.TargetColumns.Remove(tmp)
            Me.ListBox1.DisplayMember = "Key"
            Me.ListBox1.ValueMember = "Value"
            Me.ListBox1.DataSource = New BindingSource(If(Me.TargetColumns.Count = 0, Nothing, Me.TargetColumns), Nothing)
        End If

        setMergeEnableFlag()
    End Sub

    Private Sub LookupColumnExcel2ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LookupColumnExcel2ComboBox.SelectedIndexChanged
        setMergeEnableFlag()
    End Sub

    Private Sub TargetColumnExcel1ComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TargetColumnExcel1ComboBox.SelectedIndexChanged
        setMergeEnableFlag()
    End Sub

    Private Sub TargetColumnExcel2TextBox_TextChanged(sender As Object, e As EventArgs) Handles OutputTextBox.TextChanged
        setMergeEnableFlag()
    End Sub

    Private Sub MergeButton_Click(sender As Object, e As EventArgs) Handles MergeButton.Click
        Dim sheet1 As MySheet = CType(Me.SheetsExcel1ComboBox.SelectedValue, MySheet)
        Dim sheet2 As MySheet = CType(Me.SheetsExcel2ComboBox.SelectedValue, MySheet)
        Dim Lookup1Column As String = Me.LookupColumnExcel1ComboBox.Text
        Dim Lookup2Column As String = Me.LookupColumnExcel2ComboBox.Text
        Dim TargetColumns As New List(Of String)

        For Each item As String In Me.ListBox1.Items
            TargetColumns.Add(item)
        Next
        AppModule.Merge(Me.ExcelFile1TextBox.Text, Me.ExcelFile2TextBox.Text, sheet1, sheet2, Lookup1Column, TargetColumns, Lookup2Column)

        Me.OutputTextBox.Text = My.Settings.OutputFile
    End Sub

    Private Sub OutputButton_Click(sender As Object, e As EventArgs) Handles OutputButton.Click
        Me.FolderBrowserDialog1.ShowDialog()
        My.Settings.OutputDirectory = Me.FolderBrowserDialog1.SelectedPath
        Me.OutputTextBox.Text = My.Settings.OutputDirectory

    End Sub

    Private TargetColumns As New List(Of String)
    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles AddButton.Click

        If Not String.IsNullOrEmpty(Me.TargetColumnExcel1ComboBox.Text) And Me.LookupColumnExcel1ComboBox.Text <> Me.TargetColumnExcel1ComboBox.Text Then
            If Not TargetColumns.Contains(Me.TargetColumnExcel1ComboBox.Text) Then
                Me.TargetColumns.Add(Me.TargetColumnExcel1ComboBox.Text)
            End If

            Me.ListBox1.DisplayMember = "Key"
            Me.ListBox1.ValueMember = "Value"
            Me.ListBox1.DataSource = New BindingSource(If(Me.TargetColumns.Count = 0, Nothing, Me.TargetColumns), Nothing)
        End If




    End Sub

    Private Sub RemoveButton_Click(sender As Object, e As EventArgs) Handles RemoveButton.Click
        If Not String.IsNullOrEmpty(Me.ListBox1.SelectedValue) Then
            If Me.TargetColumns.Contains(Me.ListBox1.SelectedValue.ToString) Then
                Me.TargetColumns.Remove(Me.ListBox1.SelectedValue.ToString)
            End If

            Me.ListBox1.DisplayMember = "Key"
            Me.ListBox1.ValueMember = "Value"
            Me.ListBox1.DataSource = New BindingSource(If(Me.TargetColumns.Count = 0, Nothing, Me.TargetColumns), Nothing)
        End If



    End Sub

    'Private Sub MatchExactCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles MatchExactCheckBox.CheckedChanged
    '    My.Settings.ExactMatch = Me.MatchExactCheckBox.Checked
    '    Me.PrecisionComboBox1.SelectedIndex = 0
    'End Sub


    Private Sub PrecisionComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles PrecisionComboBox1.SelectedValueChanged
        Try
            My.Settings.Precision = Me.PrecisionComboBox1.SelectedValue
        Catch ex As Exception

        End Try

    End Sub

   
    Private Sub PrecisionComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PrecisionComboBox1.SelectedIndexChanged

    End Sub
End Class