Attribute VB_Name = "TableImport"
Sub RunTableImport()

ThisWorksheetName = InputWorksheetName()
ThisFirstCell = InputFirstCell()
ThisTableColumns = ReadColumns(ThisFirstCell)

MultipleSelection = IsMultipleSelection()

OtherPath = SelectImport(MultipleSelection)
OtherOpenWorkbook = OpenWorkbook(MultipleSelection, OtherPath)
OtherWorksheetName = InputWorksheetName()
OtherFirstCell = InputFirstCell()

Performance = TogglePerformanceSettings(False)

ImportStart = StartImport(ThisWorksheetName, ThisFirstCell, ThisTableColumns, OtherPath, OtherWorksheetName, OtherFirstCell, _
OtherTableColumns, MultipleSelection)
 
Performance = TogglePerformanceSettings(True)
 
End Sub

Function TogglePerformanceSettings(ByVal Settings As Boolean)
    If Settings = False Then
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
    End If
End Function

Function IsWorksheet(ByVal WorksheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = WorksheetName Then IsWorksheet = True
    Next ws
End Function

Function IsCell(ByVal CellAddress As String) As Boolean
    On Error Resume Next
        If IsError(Range(CellAddress).Select) = True Then
            IsCell = False
        Else
            IsCell = True
        End If
    On Error GoTo 0
End Function

Function IsCellEmpty(ByVal CellAddress As String) As Boolean
    If Range(CellAddress) = "" Then
        IsCellEmpty = True
    Else
        IsCellEmpty = False
    End If
End Function

Function IsMultipleSelection() As Boolean
    Answer = MsgBox("The data will be imported from more than one file?", vbYesNo)
    If Answer = vbYes Then IsMultipleSelection = True Else IsMultipleSelection = False
End Function

Function ReadColumns(ByVal FirstCellAddress As String) As Variant
    ReadColumns = Application.Transpose(Application.Transpose(Range(FirstCellAddress, Range(FirstCellAddress).End(xlToRight))))
End Function

Function MatchIndex(ByVal SearchArray As Variant, ByVal BaseArray As Variant) As Variant
    For x = LBound(SearchArray) To UBound(SearchArray)
        For y = LBound(BaseArray) To UBound(BaseArray)
            If x = LBound(SearchArray) And SearchArray(x) = BaseArray(y) Then
                MatchIndex = y - 1
            ElseIf x <> LBound(SearchArray) And SearchArray(x) = BaseArray(y) Then
                MatchIndex = MatchIndex & "," & y - 1
            End If
        Next y
    Next x
    MatchIndex = Split(MatchIndex, ",")
End Function

Function CountTableRows(ByVal CellAddress As String) As Long
    CountTableRows = Range(CellAddress).CurrentRegion.Rows.Count
End Function

Function GetRangeAddress(ByVal FirstCellAddress As String, ByVal ColumnRelativeIndexArray As Variant, ByVal TableRows As Long) As Variant
    For i = LBound(ColumnRelativeIndexArray) To UBound(ColumnRelativeIndexArray)
        ColumnAbsoluteIndex = Range(FirstCellAddress).Column + ColumnRelativeIndexArray(i)
        RowAbsoluteIndex = Range(FirstCellAddress).Row + TableRows
        If i = LBound(ColumnRelativeIndexArray) Then
            GetRangeAddress = Range(Cells(Range(FirstCellAddress).Row + 1, ColumnAbsoluteIndex), Cells(RowAbsoluteIndex, ColumnAbsoluteIndex)).Address
        Else
            GetRangeAddress = GetRangeAddress & "," & _
            Range(Cells(Range(FirstCellAddress).Row + 1, ColumnAbsoluteIndex), Cells(RowAbsoluteIndex, ColumnAbsoluteIndex)).Address
        End If
    Next i
    GetRangeAddress = Split(GetRangeAddress, ",")
End Function

Function InputWorksheetName() As Variant
    Do Until WorksheetExists = True
        InputWorksheetName = Application.InputBox("Write the Worksheet name to be used.")
        If InputWorksheetName = False Then Exit Do
            WorksheetExists = IsWorksheet(InputWorksheetName)
        If WorksheetExists = False Then
            ErrorMsg = MsgBox("This Worksheet does not exist in the actual Workbook.", vbExclamation, "Error")
        End If
    Loop
    Sheets(InputWorksheetName).Activate
End Function

Function InputFirstCell() As String
    Do
        InputFirstCell = Application.InputBox("Write the first cell address from the table to be used.")
        If InputFirstCell = "" Then End
        CheckCell = IsCell(InputFirstCell)
        If CheckCell = False Then
            ErrorMsg = MsgBox("This cell does not exist.", vbExclamation, "Error")
        Else
            FirstCellEmpty = IsCellEmpty(InputFirstCell)
            If FirstCellEmpty = True Then ErrorMsg = MsgBox("This cell is empty.", vbExclamation, "Error")
        End If
    Loop Until CheckCell = True And FirstCellEmpty = False
End Function

Function SelectImport(ByVal MultipleSelection As Boolean)
    If MultipleSelection = True Then
        SelectImport = SelectMultipleFiles()
    Else
        SelectImport = SelectSingleFile()
    End If
End Function

Function SelectMultipleFiles() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then
            SelectMultipleFiles = .SelectedItems(1)
        End If
    End With
End Function

Function SelectSingleFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Select a file"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> 0 Then
            SelectSingleFile = .SelectedItems(1)
        End If
    End With
End Function

Function OpenWorkbook(ByVal MultipleSelection As Boolean, ByVal OtherPath As String)
    If MultipleSelection = True Then
        OtherWorkbookPath = Dir(OtherPath & "\" & "*.xls*")
        Workbooks.Open OtherPath & "\" & OtherWorkbookPath
    Else
        Workbooks.Open (OtherPath)
    End If
End Function

Function StartImport(ByVal ThisWorksheetName As String, ByVal ThisFirstCell As String, ByVal ThisTableColumns As Variant, _
ByVal OtherPath As String, ByVal OtherWorksheetName As String, ByVal OtherFirstCell As String, ByVal OtherTableColumns As String, _
ByVal IsMultipleSelection As Boolean)
    If IsMultipleSelection = True Then
        GetImportMode = MultipleImport(ThisWorksheetName, ThisFirstCell, ThisTableColumns, OtherPath, OtherWorksheetName, _
        OtherFirstCell, OtherTableColumns)
    Else
        GetImportMode = SingleImport(ThisWorksheetName, ThisFirstCell, ThisTableColumns, OtherPath, OtherWorksheetName, _
        OtherFirstCell, OtherTableColumns)
    End If
End Function

Function MultipleImport(ByVal ThisWorksheetName As String, ByVal ThisFirstCell As String, ByVal ThisTableColumns As Variant, _
ByVal OtherPath As String, ByVal OtherWorksheetName As String, ByVal OtherFirstCell As String, ByVal OtherTableColumns As Variant)
    OtherWorkbookPath = Dir(OtherPath & "\" & "*.xls*")
    Do
        Workbooks.Open OtherPath & "\" & OtherWorkbookPath
        Set OtherWorkbook = ActiveWorkbook
        Worksheets(OtherWorksheetName).Activate
        OtherTableColumns = ReadColumns(OtherFirstCell)
        ColumnsIndex = MatchIndex(ThisTableColumns, OtherTableColumns)
        TableRows = CountTableRows(OtherFirstCell)
        RangeAddress = GetRangeAddress(OtherFirstCell, ColumnsIndex, TableRows)
        CopyPaste = StartCopyPaste(OtherWorkbook, ColumnsIndex, RangeAddress, ThisFirstCell)
        OtherWorkbookPath = Dir
    Loop While OtherWorkbookPath <> ""
End Function

Function SingleImport(ByVal ThisWorksheetName As String, ByVal ThisFirstCell As String, ByVal ThisTableColumns As Variant, _
ByVal OtherPath As String, ByVal OtherWorksheetName As String, ByVal OtherFirstCell As String, ByVal OtherTableColumns As String)
    OtherWorkbookPath = OtherPath
    Workbooks.Open OtherWorkbookPath
    Set OtherWorkbook = ActiveWorkbook
    Worksheets(OtherWorksheetName).Activate
    OtherTableColumns = ReadColumns(OtherFirstCell)
    ColumnsIndex = MatchIndex(ThisTableColumns, OtherTableColumns)
    TableRows = CountTableRows(OtherFirstCell)
    RangeAddress = GetRangeAddress(OtherFirstCell, ColumnsIndex, TableRows)
    CopyPaste = StartCopyPaste(OtherWorkbook, ColumnsIndex, RangeAddress, ThisFirstCell)
End Function

Function StartCopyPaste(ByVal OtherWorkbook As Workbook, ByVal ColumnsIndex As Variant, ByVal RangeAddress As Variant, _
ByVal ThisFirstCell As String)
    ThisWorkbook.Activate
    TableRows = CountTableRows(ThisFirstCell)
    For i = LBound(RangeAddress) To UBound(RangeAddress)
        OtherWorkbook.Activate
        Range(RangeAddress(i)).Copy
        ThisWorkbook.Activate
        Range(ThisFirstCell).Offset(TableRows, i).PasteSpecial xlPasteValues
    Next i
    OtherWorkbook.Close
End Function
