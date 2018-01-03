Imports DevExpress.Spreadsheet


Public Module AppModule

    Public Function ReadExcelFile(ByVal FilePath As String, ByRef Sheets As Dictionary(Of String, MySheet), Optional HeaderRowOrder As Integer = 0) As Boolean
        Try
            If Sheets Is Nothing Then
                Sheets = New Dictionary(Of String, MySheet)
            End If
            Dim workbook As New Workbook()
            workbook.LoadDocument(FilePath)

            For Each sheet As Worksheet In workbook.Sheets
                Dim Mysheet As New MySheet(sheet.Name)
                If sheet.GetUsedRange.ColumnCount = 0 Then
                    Continue For
                End If
                For i As Integer = 0 To sheet.GetUsedRange.ColumnCount - 1
                    If Not (sheet(HeaderRowOrder, i).Value.TextValue Is Nothing Or sheet(HeaderRowOrder, i).Value.TextValue = "") Then
                        Mysheet.Headers.Add(LTrim(RTrim(sheet(HeaderRowOrder, i).Value.TextValue)), i)
                    End If
                Next
                Sheets.Add(sheet.Name, Mysheet)
            Next

            workbook.Dispose()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function BindSheetsToComboBoxControl(ByRef Sheets As Dictionary(Of String, MySheet), ByRef ComboCtrl As ComboBox) As Boolean
        Try
            ComboCtrl.DisplayMember = "Key"
            ComboCtrl.ValueMember = "Value"

            ComboCtrl.DataSource = New BindingSource(Sheets, Nothing)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function


    Public Function BindSheetCoumnsToComboBoxControl(ByRef Sheet As MySheet, ByRef ComboCtrl As ComboBox) As Boolean
        Try
            ComboCtrl.DisplayMember = "Key"
            ComboCtrl.ValueMember = "Value"

            ComboCtrl.DataSource = New BindingSource(Sheet.Headers, Nothing)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function GetUniqueName(ByVal Token As String) As String
        Dim g As Guid = Guid.NewGuid
        Return String.Format("{0}-{1}", Token, g.ToString)
    End Function

    Class MySheet
        Public SheetName As String
        Public Headers As New Dictionary(Of String, Integer)

        Public Sub New(ByVal SheetName As String)
            Me.SheetName = SheetName
        End Sub
    End Class

    Class Comp
        Public Name As String, matched As Integer = 0, Notmatched As Integer = 0, percentage As Double = 0

        Public Sub New(ByVal Name As String, ByVal matched As Integer, ByVal Notmatched As Integer)
            Me.Name = Name
            Me.matched = matched
            Me.Notmatched = Notmatched
            Me.percentage = Math.Round(matched / (matched + Notmatched), 2)
        End Sub
    End Class

    Private Function IsMatchedWithDistance(ByVal Index As Integer, ByVal Name1 As String, ByVal Name2 As String) As Boolean
        Name2 = Name2.ToLower
        Name1 = Name1.ToLower
        Dim ch As Char = Name1(Index)

        Dim count As Integer
        Dim diff As Integer = Integer.MaxValue

        count = Name2.Count(Function(c As Char) c = ch)
        If count = 0 Then
            Return False
        End If

        If count = 1 Then ' occured 1 time, check the difference
            diff = Math.Abs(Index - Name2.IndexOf(ch))
            If diff <= 1 Then
                Return True
            Else
                Return False
            End If
        Else ' more than one occurance, take the nearest occurance to our index
            Dim near As Integer = Integer.MaxValue : Dim nearIndex As Integer = Integer.MaxValue
            For i As Integer = 0 To Name2.Length - 1
                If Name2(i) = ch Then
                    If near > Math.Abs(Index - i) Then
                        near = Math.Abs(Index - i)
                        nearIndex = i
                    End If
                End If
            Next

            'check as the first method
            diff = Math.Abs(Index - nearIndex)
            If diff <= 1 Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Public Function SmartMatch(ByVal Name As String, ByRef arr2 As Dictionary(Of String, String), Optional PreDefinedPercentage As Double = 0.8) As String
        Dim res As New List(Of Comp)
        Dim gPercent As Double = Double.MinValue : Dim comp As Comp = Nothing


        'First we search for exact match
        For Each SecondName As String In arr2.Keys
            If Name.ToLower = SecondName.ToLower Then
                Return SecondName
            End If
        Next

        If My.Settings.Precision = 100 Then
            Return Nothing
        Else
            PreDefinedPercentage = My.Settings.Precision / 100
        End If

        'Calculated the closest Name with Percentage
        res = New List(Of Comp)
        For Each SecondName As String In arr2.Keys
            If Not If(Name.Length > SecondName.Length, (SecondName.Length / Name.Length), (Name.Length / SecondName.Length)) >= PreDefinedPercentage Then
                res.Add(New Comp(SecondName, 0, 100))
            End If
            Dim matched As Integer = 0, Notmatched As Integer = 0
            For i As Integer = 0 To Name.Length - 1
                If IsMatchedWithDistance(i, Name, SecondName) Then
                    matched += 1
                Else
                    Notmatched += 1
                End If
            Next
            res.Add(New Comp(SecondName, matched, Notmatched))
        Next

        For Each itm As Comp In res
            If itm.percentage > gPercent Then
                gPercent = itm.percentage
                comp = itm
            End If
        Next

        If comp IsNot Nothing Then
            If comp.percentage >= PreDefinedPercentage Then
                Return comp.Name
            End If
        End If

        Return Nothing
    End Function

    Public Function Merge(ByVal FilePath1 As String, ByVal filePath2 As String, ByRef Sheet1 As MySheet, ByRef Sheet2 As MySheet, ByVal LookupColumn1 As String, ByVal TargetColumns As List(Of String), ByVal LookupColumn2 As String, Optional HeaderRowOrder As Integer = 0) As Boolean
        Try
            Dim dictionaries As New Dictionary(Of String, Dictionary(Of String, String))

            Dim workbook1 As New Workbook()
            Dim workbook2 As New Workbook()

            workbook1.LoadDocument(FilePath1)
            workbook2.LoadDocument(filePath2)

            Dim ExcelSheet1 As Worksheet = workbook1.Sheets(Sheet1.SheetName)
            Dim ExcelSheet2 As Worksheet = workbook2.Sheets(Sheet2.SheetName)

            For i As Integer = HeaderRowOrder + 1 To ExcelSheet1.GetUsedRange.RowCount - 1
                For Each NewColumn As String In TargetColumns
                    Dim dict As Dictionary(Of String, String) = Nothing
                    Dim key As String = ExcelSheet1(i, Sheet1.Headers(LookupColumn1)).Value.ToString()
                    Dim val As String = ExcelSheet1(i, Sheet1.Headers(NewColumn)).Value.ToString()


                    'get dictionary from dictionaries
                    If Not dictionaries.Keys.Contains(NewColumn) Then
                        dict = New Dictionary(Of String, String)
                        dict.Add(key, val)
                        dictionaries.Add(NewColumn, dict)
                    Else                        
                        If dictionaries(NewColumn).Keys.Contains(val) Then
                            dictionaries(NewColumn)(val) = val
                        Else
                            dictionaries(NewColumn).Add(key, val)
                        End If
                    End If

                Next
            Next

            'second file 
            For Each NewColumnName As String In TargetColumns
                Dim NewColIndex As Integer = ExcelSheet2.GetUsedRange.ColumnCount
                ExcelSheet2(HeaderRowOrder, NewColIndex).Value = NewColumnName

                For i As Integer = HeaderRowOrder + 1 To ExcelSheet2.GetUsedRange.RowCount - 1
                    Dim key As String = ExcelSheet2(i, Sheet2.Headers(LookupColumn2)).Value.ToString()
                    Dim res As String = SmartMatch(key, dictionaries(NewColumnName))
                    If res IsNot Nothing Then
                        ExcelSheet2(i, NewColIndex).Value = dictionaries(NewColumnName)(res)
                    End If
                Next
            Next




            workbook1.Dispose()
            My.Settings.OutputFile = String.Format("{0}\{1}.xlsx", My.Settings.OutputDirectory, GetUniqueName("Result"))
            workbook2.SaveDocument(My.Settings.OutputFile)
            workbook2.Dispose()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Module
