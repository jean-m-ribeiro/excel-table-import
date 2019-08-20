Attribute VB_Name = "TableImport"
Sub data_table_import()

'Altera configurações para aumentar a velocidade de execução da macro'
Application.DisplayAlerts = False: Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual

Set this_workbook = ActiveWorkbook
this_worksheet_name = input_worksheet_name()
this_first_cell = input_first_cell()
this_table_columns = read_columns(this_first_cell)

multiple_selection = is_multiple_selection()

other_path = select_import(multiple_selection)
other_workbook_open = workbook_open(multiple_selection, other_path)
other_worksheet_name = input_worksheet_name()
other_first_cell = input_first_cell()

import = start_import(this_workbook, this_worksheet_name, this_first_cell, this_table_columns, other_path, other_worksheet_name, other_first_cell, other_table_columns, multiple_selection)
 
 'Restaura configurações padrão'
Application.DisplayAlerts = True: Application.ScreenUpdating = True: Application.Calculation = xlCalculationAutomatic
 
End Sub

Function is_worksheet(ByVal worksheet_name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = worksheet_name Then is_worksheet = True
    Next ws
End Function

Function is_cell(ByVal cell_address As String) As Boolean
    On Error Resume Next
        If IsError(Range(cell_address).Select) = True Then
            is_cell = False
        Else
            is_cell = True
        End If
    On Error GoTo 0
End Function

Function is_cell_empty(ByVal cell_address As String) As Boolean
    If Range(cell_address) = "" Then
        is_cell_empty = True
    Else
        is_cell_empty = False
    End If
End Function

Function is_multiple_selection() As Boolean
    answer = MsgBox("A importação de dados será feita de todos arquivos dentro de uma pasta?", vbYesNo)
    If answer = vbYes Then is_multiple_selection = True Else is_multiple_selection = False
End Function

Function read_columns(ByVal first_cell_address As String) As Variant
    read_columns = Application.Transpose(Application.Transpose(Range(first_cell_address, Range(first_cell_address).End(xlToRight))))
End Function

Function match_index(ByVal search_array As Variant, ByVal base_array As Variant) As Variant
    For x = LBound(search_array) To UBound(search_array)
        For y = LBound(base_array) To UBound(base_array)
            If x = LBound(search_array) And search_array(x) = base_array(y) Then
                match_index = y - 1
            ElseIf x <> LBound(search_array) And search_array(x) = base_array(y) Then
                match_index = match_index & "," & y - 1
            End If
        Next y
    Next x
    match_index = Split(match_index, ",")
End Function

Function count_table_rows(ByVal cell_address As String) As Long
    count_table_rows = Range(cell_address).CurrentRegion.Rows.Count
End Function

Function get_range_address(ByVal first_cell_address As String, ByVal column_relative_index_array As Variant, ByVal table_rows As Long) As Variant
    For i = LBound(column_relative_index_array) To UBound(column_relative_index_array)
        column_absolute_index = Range(first_cell_address).Column + column_relative_index_array(i)
        row_absolute_index = Range(first_cell_address).Row + table_rows
        If i = LBound(column_relative_index_array) Then
            get_range_address = Range(Cells(Range(first_cell_address).Row + 1, column_absolute_index), Cells(row_absolute_index, column_absolute_index)).Address
        Else
            get_range_address = get_range_address & "," & Range(Cells(Range(first_cell_address).Row + 1, column_absolute_index), Cells(row_absolute_index, column_absolute_index)).Address
        End If
    Next i
    get_range_address = Split(get_range_address, ",")
End Function

Function input_worksheet_name() As Variant
    Do Until worksheet_exists = True
        input_worksheet_name = Application.InputBox("Digite o nome da planilha para onde os dados devem ser importados")
        If input_worksheet_name = False Then Exit Do
            worksheet_exists = is_worksheet(input_worksheet_name)
        If worksheet_exists = False Then
            error_msg = MsgBox("Esta planilha não existe na pasta de trabalho atual.", vbExclamation, "Erro")
        End If
    Loop
    Sheets(input_worksheet_name).Activate
End Function

Function input_first_cell() As String
    Do
        input_first_cell = Application.InputBox("Digite a coluna e a linha da primeira célula da tabela para onde os dados devem ser importados")
        If input_first_cell = "" Then End
        check_cell = is_cell(input_first_cell)
        If check_cell = False Then
            error_msg = MsgBox("Esta célula não existe.", vbExclamation, "Erro")
        Else
            first_cell_empty = is_cell_empty(input_first_cell)
            If first_cell_empty = True Then error_msg = MsgBox("Esta célula está vazia.", vbExclamation, "Erro")
        End If
    Loop Until check_cell = True And first_cell_empty = False
End Function

Function select_import(ByVal multiple_selection As Boolean)
    If multiple_selection = True Then
        select_import = select_multiple_files()
    Else
        select_import = select_single_file()
    End If
End Function

Function select_multiple_files() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Selecione uma pasta"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show = -1 Then
            select_multiple_files = .SelectedItems(1)
        End If
    End With
End Function

Function select_single_file() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Selecione um arquivo"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> 0 Then
            select_single_file = .SelectedItems(1)
        End If
    End With
End Function

Function workbook_open(ByVal multiple_selection As Boolean, ByVal other_path As String)
    If multiple_selection = True Then
        other_workbook_path = Dir(other_path & "\" & "*.xls*")
        Workbooks.Open other_path & "\" & other_workbook_path
    Else
        Workbooks.Open (other_path)
    End If
End Function

Function start_import(ByVal this_workbook As Workbook, ByVal this_worksheet_name As String, ByVal this_first_cell As String, ByVal this_table_columns As Variant, ByVal other_path As String, ByVal other_worksheet_name As String, ByVal other_first_cell As String, ByVal other_table_columns As String, ByVal is_multiple_selection As Boolean)
    If is_multiple_selection = True Then
        get_import_mode = multiple_import(this_workbook, this_worksheet_name, this_first_cell, this_table_columns, other_path, other_worksheet_name, other_first_cell, other_table_columns)
    Else
        get_import_mode = single_import(this_workbook, this_worksheet_name, this_first_cell, this_table_columns, other_path, other_worksheet_name, other_first_cell, other_table_columns)
    End If
End Function

Function multiple_import(ByVal this_workbook As Workbook, ByVal this_worksheet_name As String, ByVal this_first_cell As String, ByVal this_table_columns As Variant, ByVal other_path As String, ByVal other_worksheet_name As String, ByVal other_first_cell As String, ByVal other_table_columns As Variant)
    other_workbook_path = Dir(other_path & "\" & "*.xls*")
    Do
        Workbooks.Open other_path & "\" & other_workbook_path
        Set other_workbook = ActiveWorkbook
        Worksheets(other_worksheet_name).Activate
        other_table_columns = read_columns(other_first_cell)
        columns_index = match_index(this_table_columns, other_table_columns)
        table_rows = count_table_rows(other_first_cell)
        range_address = get_range_address(other_first_cell, columns_index, table_rows)
        copy_paste = start_copy_paste(this_workbook, other_workbook, columns_index, range_address, this_first_cell)
        other_workbook_path = Dir
    Loop While other_workbook_path <> ""
End Function

Function single_import(ByVal this_workbook As Workbook, ByVal this_worksheet_name As String, ByVal this_first_cell As String, ByVal this_table_columns As Variant, ByVal other_path As String, ByVal other_worksheet_name As String, ByVal other_first_cell As String, ByVal other_table_columns As String)
    other_workbook_path = other_path
    Workbooks.Open other_workbook_path
    Set other_workbook = ActiveWorkbook
    Worksheets(other_worksheet_name).Activate
    other_table_columns = read_columns(other_first_cell)
    columns_index = match_index(this_table_columns, other_table_columns)
    table_rows = count_table_rows(other_first_cell)
    range_address = get_range_address(other_first_cell, columns_index, table_rows)
    copy_paste = start_copy_paste(this_workbook, other_workbook, columns_index, range_address, this_first_cell)
End Function

Function start_copy_paste(ByVal this_workbook As Workbook, ByVal other_workbook As Workbook, ByVal columns_index As Variant, ByVal range_address As Variant, ByVal this_first_cell As String)
    this_workbook.Activate
    table_rows = count_table_rows(this_first_cell)
    For i = LBound(range_address) To UBound(range_address)
        other_workbook.Activate
        Range(range_address(i)).Copy
        this_workbook.Activate
        Range(this_first_cell).Offset(table_rows, i).PasteSpecial xlPasteValues
    Next i
    other_workbook.Close
End Function
