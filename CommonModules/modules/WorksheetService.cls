VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorksheetService

' #############################################################################
'!
'! @brief
'! ワークシートへのアクセスを管理するサービスの実装です。
'!
' #############################################################################

'
' 範囲を検索します
'
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    If RangeBounds.IsEmpty Then Exit Function

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)
    
    Dim look_in As Long
    If LookInValues Then
        look_in = xlValues
    Else
        look_in = xlFormulas
    End If
    
    Dim look_at As Long
    If LookAtWhole Then
        look_at = xlWhole
    Else
        look_at = xlPart
    End If
    
    Dim target_range As Range
    If RangeBounds.IsEntireSheet Then
        ' シート全体を検索
        Set target_range = target_sheet.Cells
    Else
        ' 範囲を検索
        Set target_range = target_sheet.Range(RangeBounds.ToString(CellOnly:=True))
    End If
    
    ' 1 個目を検索
    Dim first_cell As Range
    Set first_cell = target_range.Find( _
            What:=What, _
            After:=target_range.Cells(target_range.Rows.Count, target_range.Columns.Count), _
            LookIn:=look_in, _
            LookAt:=look_at, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte, _
            SearchFormat:=False)
    
    ' 1 個も見つからない場合は即終了
    If first_cell Is Nothing Then
        Call pResetFindState(target_sheet)
        Exit Function
    End If
    
    Dim result_arr() As WorksheetRangeBounds
    Dim found_cell As Range
    Set found_cell = first_cell
    Do
        Call pAddItemToRangeBoundsArray(result_arr, found_cell)
        Set found_cell = target_range.FindNext(After:=found_cell)
    Loop While found_cell.Address <> first_cell.Address
    
    ' 完全一致を解除する (通常使用時に完全一致は扱いづらい)
    Call pResetFindState(target_sheet)
    
    Find = result_arr
End Function
Private Function IWorksheetService_Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    IWorksheetService_Find = Find( _
            What, RangeBounds, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

Private Sub pResetFindState(ByVal TargetSheet As Worksheet)
    Call TargetSheet.Range("A1:A2").Find( _
            What:="", _
            After:=TargetSheet.Range("A1"), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=True, _
            MatchByte:=True, _
            SearchFormat:=False)
End Sub

Private Sub pAddItemToRangeBoundsArray(ByRef TargetArray() As WorksheetRangeBounds, ByVal TargetCell As Range)
    Dim item_idx As Long
    If (Not (Not TargetArray)) = 0 Then
        item_idx = 0
        ReDim TargetArray(0 To 0)
    Else
        item_idx = UBound(TargetArray) + 1
        ReDim Preserve TargetArray(LBound(TargetArray) To item_idx)
    End If
    Set TargetArray(item_idx) = New_RangeBounds(Row:=TargetCell.Row, Column:=TargetCell.Column, Sheet:=TargetCell.Worksheet.Name, Book:=TargetCell.Worksheet.Parent.Name)
End Sub

Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub
Private Sub IWorksheetService_Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub

Private Sub pSortCore(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    If IsEmptyArray(SortKeyAndOrder) Then Exit Sub
    If RangeBounds.IsEmpty Then Exit Sub
    If (UBound(SortKeyAndOrder) - LBound(SortKeyAndOrder)) Mod 2 = 0 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", "SortKeyAndOrder の要素数が偶数ではありません。(奇数番目の要素は、ソートのキーとなる列の番号、偶数番目の要素は、xlAscending か xlDescending)"
    End If
    Call pRaiseIfSortKeyAndOrderInvalid(RangeBounds, SortKeyAndOrder)

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Dim sort_range As Range
    Set sort_range = target_sheet.Range(RangeBounds.ToString(CellOnly:=True))

    With target_sheet.Sort
        .SortFields.Clear

        Dim item_idx As Long
        For item_idx = LBound(SortKeyAndOrder) To UBound(SortKeyAndOrder) - 1 Step 2
            Dim sort_key As Long
            sort_key = CLng(SortKeyAndOrder(item_idx))
            If IsEmpty(SortKeyAndOrder(item_idx + 1)) Then
                .SortFields.Add Key:=target_sheet.Cells(RangeBounds.Row, RangeBounds.Column + sort_key - 1)
            Else
                .SortFields.Add Key:=target_sheet.Cells(RangeBounds.Row, RangeBounds.Column + sort_key - 1), Order:=CLng(SortKeyAndOrder(item_idx + 1))
            End If
        Next item_idx

        .SetRange sort_range
        .Header = xlNo
        .Apply
    End With
End Sub

Private Sub pRaiseIfSortKeyAndOrderInvalid(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    Dim item_idx As Long
    For item_idx = LBound(SortKeyAndOrder) To UBound(SortKeyAndOrder) - 1 Step 2
        Call pGetValidatedRelativeColumnIndex(RangeBounds, SortKeyAndOrder(item_idx), "SortKeyAndOrder")
        If Not IsEmpty(SortKeyAndOrder(item_idx + 1)) Then
            Dim sort_order As Long
            sort_order = pGetValidatedLong(SortKeyAndOrder(item_idx + 1), "SortKeyAndOrder")
            If sort_order <> xlAscending And sort_order <> xlDescending Then
                Err.Raise vbObjectError + 1, "Class WorksheetService", "SortKeyAndOrder の並び順には xlAscending または xlDescending を指定してください。"
            End If
        End If
    Next item_idx
End Sub

' 範囲をアクティブにします。
'
' @details
' 範囲をアクティブにします。
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Dim target_book As Workbook
    Set target_book = pGetWorkbook(RangeBounds)

    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(RangeBounds.WorksheetName)

    Call target_book.Activate
    Call target_sheet.Activate
    target_sheet.Range(RangeBounds.ToString(CellOnly:=True)).Activate
End Sub
Private Sub IWorksheetService_ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call ActivateRange(RangeBounds)
End Sub

'
' 指定されたセルに文字列または数式を設定します。
' TypeConvert が True の場合、数値文字列は整数として表せる場合は整数、それ以外は小数として扱います。
'
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
    
    If RangeBounds.IsEmpty Then Exit Sub
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService", "RangeBounds がセルではありません。(" & RangeBounds.ToString() & ")"
    
    Dim target_cell As Range
    Set target_cell = pGetCell(RangeBounds)
    
    If Expression = "" And IgnoreEmpty Then Exit Sub
    
    ' 表示形式を設定する
    If NumberFormat <> "" Then
        target_cell.NumberFormatLocal = NumberFormat
    End If
    
    ' 空白の処理
    If Expression = "" Then
        If ClearWhenEmpty Then
            target_cell.ClearContents
        Else
            target_cell.Formula = "="""""
            Call target_cell.Copy
            Call target_cell.PasteSpecial(Paste:=xlPasteValues, SkipBlanks:=False)
            Application.CutCopyMode = False
        End If
        Exit Sub
    End If
    
    If AsFormula Then
        ' 通常数式の処理
        target_cell.Formula = Expression
    Else
        If TypeConvert Then
            ' 型変換する場合
            Dim err_val As Variant
            If pTryConvertErrorStringToCVErr(Expression, err_val) Then
                target_cell.Value = err_val  ' CVErr(...) をセット
                Exit Sub
            End If
            
            Dim long_value As Long
            Dim double_value As Double
            Dim long_error As Long
            Dim double_error As Long
            
            On Error Resume Next
            
            Err.Clear
            long_value = CLng(Expression)
            long_error = Err.Number
            
            Err.Clear
            double_value = CDbl(Expression)
            double_error = Err.Number
            
            On Error GoTo 0
            
            If long_error = 0 Or double_error = 0 Then
                If long_error = 0 And double_error = 0 Then
                    If double_value - CDbl(long_value) = 0# Then
                        target_cell.Value = long_value
                    Else
                        target_cell.Value = double_value
                    End If
                ElseIf double_error = 0 Then
                    target_cell.Value = double_value
                Else
                    target_cell.Value = long_value
                End If
                Exit Sub
            End If
            
            On Error Resume Next
            
            Err.Clear
            target_cell.Value = CBool(Expression)
            If Err.Number = 0 Then Exit Sub
            
            Err.Clear
            target_cell.Value = CDate(Expression)
            If Err.Number = 0 Then Exit Sub
            
            Err.Clear
            target_cell.Value = CCur(Expression)
            If Err.Number = 0 Then Exit Sub
            
            Err.Clear
            On Error GoTo 0
        End If
        
        ' 文字列として処理
        Dim current_format As String
        current_format = target_cell.NumberFormatLocal

        On Error GoTo RESTORE_NUMBER_FORMAT_ON_ERROR
        target_cell.NumberFormatLocal = "@"
        target_cell.Value = Expression
        target_cell.NumberFormatLocal = current_format
        On Error GoTo 0
    End If
    Exit Sub

RESTORE_NUMBER_FORMAT_ON_ERROR:
    Dim err_num As Long
    err_num = Err.Number

    Dim err_source As String
    err_source = Err.Source

    Dim err_desc As String
    err_desc = Err.Description

    On Error Resume Next
    target_cell.NumberFormatLocal = current_format
    On Error GoTo 0

    Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc
End Sub
Private Sub IWorksheetService_WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
    
    Call WriteCell( _
            RangeBounds, Expression, _
            NumberFormat:=NumberFormat, AsFormula:=AsFormula, ClearWhenEmpty:=ClearWhenEmpty, IgnoreEmpty:=IgnoreEmpty, TypeConvert:=TypeConvert)
End Sub

'
' 範囲に 2 次元配列を書き込みます。
'
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Dim value_row_count As Long
    Dim value_column_count As Long
    Call pGetTwoDimensionalArraySize(value_row_count, value_column_count, ValuesArray)
    Call pRaiseIfDestinationShapeIsInvalid(RangeBounds, value_row_count, value_column_count, "RangeBounds")
    Call pRaiseIfDestinationOutOfSheet(RangeBounds, value_row_count, value_column_count, "書き込み")

    Dim target_range As Range
    Set target_range = pGetDestinationRange(RangeBounds, value_row_count, value_column_count)

    target_range.Value = ValuesArray
End Sub
Private Sub IWorksheetService_WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call WriteRange(RangeBounds, ValuesArray)
End Sub

'
' 範囲に配列 (CSE) 数式を書き込みます。
'
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    target_range.FormulaArray = FormulaArray
End Sub
Private Sub IWorksheetService_WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call WriteArrayFormula(RangeBounds, FormulaArray)
End Sub

Private Function pGetWorkbook(ByVal RangeBounds As WorksheetRangeBounds) As Workbook
    If RangeBounds.WorkbookName <> "" Then
        Set pGetWorkbook = Workbooks(RangeBounds.WorkbookName)
    Else
        Set pGetWorkbook = ThisWorkbook
    End If
End Function

Private Function pGetWorksheet(ByVal RangeBounds As WorksheetRangeBounds) As Worksheet
    Set pGetWorksheet = pGetWorkbook(RangeBounds).Worksheets(RangeBounds.WorksheetName)
End Function

Private Function pGetRange(ByVal RangeBounds As WorksheetRangeBounds) As Range
    With pGetWorksheet(RangeBounds)
        Set pGetRange = .Range(.Cells(RangeBounds.Row, RangeBounds.Column), .Cells(RangeBounds.FinishRow, RangeBounds.FinishColumn))
    End With
End Function

Private Function pGetCell(ByVal RangeBounds As WorksheetRangeBounds) As Range
    Set pGetCell = pGetWorksheet(RangeBounds).Cells(RangeBounds.Row, RangeBounds.Column)
End Function

Private Function pGetDestinationRange( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long) As Range

    If DestinationRangeBounds.IsEmpty Or DestinationRangeBounds.IsCell Then
        Set pGetDestinationRange = pGetCell(DestinationRangeBounds).Resize(SourceRowCount, SourceColumnCount)
    Else
        Set pGetDestinationRange = pGetRange(DestinationRangeBounds)
    End If
End Function

Private Sub pGetTwoDimensionalArraySize( _
        ByRef RowCount As Long, _
        ByRef ColumnCount As Long, _
        ByVal TargetArray As Variant)

    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", "ValuesArray は 2 次元配列を指定してください。(" & TypeName(TargetArray) & ")"
    End If

    Dim lower_bounds() As Long
    Dim upper_bounds() As Long
    Call GetArrayBounds(lower_bounds, upper_bounds, TargetArray)

    If IsEmptyArray(lower_bounds) Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", "ValuesArray は 2 次元配列を指定してください。"
    End If
    If UBound(lower_bounds) - LBound(lower_bounds) + 1 <> 2 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", "ValuesArray は 2 次元配列を指定してください。"
    End If

    RowCount = upper_bounds(LBound(upper_bounds)) - lower_bounds(LBound(lower_bounds)) + 1
    ColumnCount = upper_bounds(LBound(upper_bounds) + 1) - lower_bounds(LBound(lower_bounds) + 1) + 1

    If RowCount < 1 Or ColumnCount < 1 Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", "ValuesArray は 1 要素以上の 2 次元配列を指定してください。"
    End If
End Sub

Private Sub pRaiseIfDestinationShapeIsInvalid( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long, _
        ByVal ArgumentName As String)

    If DestinationRangeBounds.IsEmpty Then Exit Sub
    If DestinationRangeBounds.IsCell Then Exit Sub
    If DestinationRangeBounds.RowCount = SourceRowCount And DestinationRangeBounds.ColumnCount = SourceColumnCount Then Exit Sub

    Err.Raise vbObjectError + 1, _
            "Class WorksheetService", _
            ArgumentName & " は空範囲、1 セル、または元と同じ行数・列数を指定してください。" & _
            "(destination: " & DestinationRangeBounds.RowCount & "x" & DestinationRangeBounds.ColumnCount & _
            ", source: " & SourceRowCount & "x" & SourceColumnCount & ")"
End Sub

Private Sub pRaiseIfDestinationOutOfSheet( _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        ByVal SourceRowCount As Long, _
        ByVal SourceColumnCount As Long, _
        ByVal OperationName As String)

    Dim destination_finish_row As Long
    destination_finish_row = DestinationRangeBounds.Row + SourceRowCount - 1
    If G_ROW_MAX < destination_finish_row Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", OperationName & "先範囲の終了行がシート上限を超過しています。(row: " & DestinationRangeBounds.Row & ", row_count: " & SourceRowCount & ", finish_row: " & destination_finish_row & ")"
    End If

    Dim destination_finish_column As Long
    destination_finish_column = DestinationRangeBounds.Column + SourceColumnCount - 1
    If G_COL_MAX < destination_finish_column Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", OperationName & "先範囲の終了列がシート上限を超過しています。(column: " & DestinationRangeBounds.Column & ", column_count: " & SourceColumnCount & ", finish_column: " & destination_finish_column & ")"
    End If
End Sub

Private Function pGetFormulaRangeAddress(ByVal RangeBounds As WorksheetRangeBounds) As String
    pGetFormulaRangeAddress = pGetRange(RangeBounds).Address( _
            RowAbsolute:=True, _
            ColumnAbsolute:=True, _
            ReferenceStyle:=xlA1, _
            External:=True)
End Function

Private Function pGetFormulaLiteral(ByVal Value As Variant) As String
    If IsError(Value) Then
        pGetFormulaLiteral = pConvertErrorToString(Value)
        Exit Function
    End If
    If IsEmpty(Value) Or IsNull(Value) Then
        pGetFormulaLiteral = """"""
        Exit Function
    End If
    
    Select Case VarType(Value)
     Case vbBoolean
        If CBool(Value) Then
            pGetFormulaLiteral = "TRUE"
        Else
            pGetFormulaLiteral = "FALSE"
        End If
     Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
        pGetFormulaLiteral = pGetFormulaNumber(CDbl(Value))
     Case vbDate
        pGetFormulaLiteral = pGetFormulaNumber(CDbl(CDate(Value)))
     Case Else
        pGetFormulaLiteral = """" & Replace(CStr(Value), """", """""") & """"
    End Select
End Function

Private Function pGetFormulaNumber(ByVal NumberValue As Double) As String
    Dim result_value As String
    result_value = LTrim$(Str$(NumberValue))
    If Left$(result_value, 1) = "." Then
        result_value = "0" & result_value
    ElseIf Left$(result_value, 2) = "-." Then
        result_value = "-0" & Mid$(result_value, 2)
    End If
    pGetFormulaNumber = result_value
End Function

Private Function pTryConvertErrorStringToCVErr(ByVal Expression As String, ByRef ErrorValue As Variant) As Boolean
    Select Case UCase(Expression)
     Case "#DIV/0!"
        ErrorValue = CVErr(xlErrDiv0)
        pTryConvertErrorStringToCVErr = True
     Case "#N/A"
        ErrorValue = CVErr(xlErrNA)
        pTryConvertErrorStringToCVErr = True
     Case "#NAME?"
        ErrorValue = CVErr(xlErrName)
        pTryConvertErrorStringToCVErr = True
     Case "#NULL!"
        ErrorValue = CVErr(xlErrNull)
        pTryConvertErrorStringToCVErr = True
     Case "#NUM!"
        ErrorValue = CVErr(xlErrNum)
        pTryConvertErrorStringToCVErr = True
     Case "#REF!"
        ErrorValue = CVErr(xlErrRef)
        pTryConvertErrorStringToCVErr = True
     Case "#VALUE!"
        ErrorValue = CVErr(xlErrValue)
        pTryConvertErrorStringToCVErr = True
     Case Else
        pTryConvertErrorStringToCVErr = False
    End Select
End Function

Private Function pConvertErrorToString(ByVal ErrValue As Variant) As String
    Select Case ErrValue
     Case CVErr(xlErrDiv0)
        pConvertErrorToString = "#DIV/0!"
     Case CVErr(xlErrNA)
        pConvertErrorToString = "#N/A"
     Case CVErr(xlErrName)
        pConvertErrorToString = "#NAME?"
     Case CVErr(xlErrNull)
        pConvertErrorToString = "#NULL!"
     Case CVErr(xlErrNum)
        pConvertErrorToString = "#NUM!"
     Case CVErr(xlErrRef)
        pConvertErrorToString = "#REF!"
     Case CVErr(xlErrValue)
        pConvertErrorToString = "#VALUE!"
     Case Else
        pConvertErrorToString = "Unknown error"
    End Select
End Function

'
' セルから値を読み取ります。
'
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    If RangeBounds.IsEmpty Then Exit Function
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService", "RangeBounds がセルではありません。(" & RangeBounds.ToString() & ")"
    
    Dim target_cell As Range
    Set target_cell = pGetCell(RangeBounds)
    
    If GetText Then
        Call pReadCellText(target_cell, Expression, NumberFormat)
    Else
        Call pReadCellValue(target_cell, Expression, NumberFormat)
    End If
    ReadCell = Expression
End Function
Private Function IWorksheetService_ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    IWorksheetService_ReadCell = ReadCell(RangeBounds, Expression:=Expression, NumberFormat:=NumberFormat, GetText:=GetText)
End Function

'
' 範囲から値を読み取ります。
'
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    If RangeBounds.IsEmpty Then Exit Function
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    If RangeBounds.IsCell Then
        Dim result_value(1 To 1, 1 To 1) As Variant
        result_value(1, 1) = target_range.Value
        ReadRange = result_value
    Else
        ReadRange = target_range.Value
    End If
End Function
Private Function IWorksheetService_ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    IWorksheetService_ReadRange = ReadRange(RangeBounds)
End Function

'
' 数式を評価します。
'
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    If ContextRangeBounds.IsEmpty Then Exit Function
    
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(ContextRangeBounds)
    
    EvaluateFormula = target_sheet.Evaluate(Formula)
End Function
Private Function IWorksheetService_EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    IWorksheetService_EvaluateFormula = EvaluateFormula(ContextRangeBounds, Formula)
End Function

'
' XLOOKUP の結果を取得します。
'
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
    
    If LookupRangeBounds.IsEmpty Then Exit Function
    If ReturnRangeBounds.IsEmpty Then Exit Function
    
    Dim xlookup_formula As String
    xlookup_formula = "=XLOOKUP(" & _
            pGetFormulaLiteral(LookupValue) & "," & _
            pGetFormulaRangeAddress(LookupRangeBounds) & "," & _
            pGetFormulaRangeAddress(ReturnRangeBounds) & "," & _
            pGetFormulaLiteral(IfNotFound) & "," & _
            CStr(MatchMode) & "," & _
            CStr(SearchMode) & ")"
    
    XLookup = EvaluateFormula(LookupRangeBounds, xlookup_formula)
End Function
Private Function IWorksheetService_XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
    
    IWorksheetService_XLookup = XLookup( _
            LookupValue, _
            LookupRangeBounds, _
            ReturnRangeBounds, _
            IfNotFound:=IfNotFound, _
            MatchMode:=MatchMode, _
            SearchMode:=SearchMode)
End Function

Private Sub pReadCellValue(ByVal TargetCell As Range, ByRef Expression As String, ByRef NumberFormat As String)
    If IsError(TargetCell.Value) Then
        Expression = pConvertErrorToString(TargetCell.Value)
    Else
        Expression = CStr(TargetCell.Value)
    End If
    NumberFormat = TargetCell.NumberFormatLocal
End Sub

Private Sub pReadCellText(ByVal TargetCell As Range, ByRef Expression As String, ByRef NumberFormat As String)
    Dim result_string As String
    result_string = TargetCell.Text
    
    If result_string <> "" And Replace(result_string, "#", "") = "" Then
        Dim current_width As Long
        current_width = TargetCell.Parent.Columns(TargetCell.Column).ColumnWidth
        TargetCell.Parent.Columns(TargetCell.Column).ColumnWidth = 254.88
        result_string = TargetCell.Text
        TargetCell.Parent.Columns(TargetCell.Column).ColumnWidth = current_width
        
        If result_string <> "" And Replace(result_string, "#", "") = "" Then
            result_string = "" & TargetCell.Value
        End If
    End If
    
    Expression = result_string
    NumberFormat = TargetCell.NumberFormat
End Sub

'
' 指定されたセルの値または数式をコピーします。
' 数式のコピーでは配列数式にも対応します。
'
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    If SourceRangeBounds.IsEmpty Then Exit Sub
    If DestinationRangeBounds.IsEmpty Then Exit Sub
    If Not SourceRangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService", "SourceRangeBounds がセルではありません。(" & SourceRangeBounds.ToString() & ")"
    If Not DestinationRangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService", "DestinationRangeBounds がセルではありません。(" & DestinationRangeBounds.ToString() & ")"
    
    Dim src_sheet As Worksheet
    Set src_sheet = pGetWorksheet(SourceRangeBounds)
    
    Dim dst_sheet As Worksheet
    Set dst_sheet = pGetWorksheet(DestinationRangeBounds)
    
    Call pCopyCellCore( _
        SourceRangeBounds.Row, SourceRangeBounds.Column, _
        DestinationRangeBounds.Row, DestinationRangeBounds.Column, _
        src_sheet, dst_sheet, AsValue, CopyNumberFormat, RemainRowCount:=1, RemainColumnCount:=1, ForcePasteIfArrayFormula:=True)
End Sub
Private Sub IWorksheetService_CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    Call CopyCell(SourceRangeBounds, DestinationRangeBounds, AsValue:=AsValue, CopyNumberFormat:=CopyNumberFormat)
End Sub

'
' 指定された範囲の値または数式をコピーします。
' 数式のコピーでは配列数式にも対応します。
'
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    If SourceRangeBounds.IsEmpty Then Exit Sub
    Call pRaiseIfDestinationShapeIsInvalid(DestinationRangeBounds, SourceRangeBounds.RowCount, SourceRangeBounds.ColumnCount, "DestinationRangeBounds")
    Call pRaiseIfCopyRangeDestinationOutOfSheet(SourceRangeBounds, DestinationRangeBounds)
    
    Dim src_sheet As Worksheet
    Set src_sheet = pGetWorksheet(SourceRangeBounds)
    
    Dim dst_sheet As Worksheet
    Set dst_sheet = pGetWorksheet(DestinationRangeBounds)
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = GetUsedRangeBounds(SourceRangeBounds, IgnoreEmpty:=False, IsUsedStartRow:=True, IsUsedStartColumn:=True)
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds( _
            Row:=DestinationRangeBounds.Row, Column:=DestinationRangeBounds.Column, _
            FinishRow:=DestinationRangeBounds.Row + SourceRangeBounds.RowCount - 1, _
            FinishColumn:=DestinationRangeBounds.Column + SourceRangeBounds.ColumnCount - 1, _
            Sheet:=DestinationRangeBounds.WorksheetName, Book:=DestinationRangeBounds.WorkbookName)
    Set dst_bounds = GetUsedRangeBounds(dst_bounds, IgnoreEmpty:=False, IsUsedStartRow:=True, IsUsedStartColumn:=True)
    
    Dim last_row_idx As Long, last_col_idx As Long
    last_row_idx = MaxLng(src_bounds.RowCount, dst_bounds.RowCount) - 1
    last_col_idx = MaxLng(src_bounds.ColumnCount, dst_bounds.ColumnCount) - 1

    Dim row_start_idx As Long, row_end_idx As Long, row_step As Long
    row_start_idx = 0
    row_end_idx = last_row_idx
    row_step = 1

    Dim col_start_idx As Long, col_end_idx As Long, col_step As Long
    col_start_idx = 0
    col_end_idx = last_col_idx
    col_step = 1

    If src_sheet Is dst_sheet Then
        If src_bounds.Row < dst_bounds.Row Then
            row_start_idx = last_row_idx
            row_end_idx = 0
            row_step = -1
        End If

        If src_bounds.Column < dst_bounds.Column Then
            col_start_idx = last_col_idx
            col_end_idx = 0
            col_step = -1
        End If
    End If

    ' 処理済みの配列数式を記録する配列を準備
    Dim procd_arr() As String: ReDim procd_arr(0)
    Dim row_idx As Long, col_idx As Long
    For row_idx = row_start_idx To row_end_idx Step row_step
        For col_idx = col_start_idx To col_end_idx Step col_step
            ' 配列数式のアンカーセルを得る (配列数式でない場合は Nothing)
            Dim arr_anchor As Range
            Set arr_anchor = pGetArrayFormulaAnchor(src_sheet, src_bounds.Row + row_idx, src_bounds.Column + col_idx)
            If pIsArrayFormulaToBeProcessed(arr_anchor, procd_arr) Then
                ' 配列数式であり、かつ、未処理の場合は、
                Dim arr_rows As Long, arr_cols As Long
                ' 配列数式の、貼り付け先でのサイズを算出
                Call pCalcurateArrayCount(arr_rows, arr_cols, arr_anchor, SourceRangeBounds)
                ' 配列数式として強制的に処理
                Call pCopyCellCore( _
                    src_bounds.Row + row_idx, src_bounds.Column + col_idx, _
                    dst_bounds.Row + row_idx, dst_bounds.Column + col_idx, _
                    src_sheet, dst_sheet, AsValue, CopyNumberFormat, _
                    RemainRowCount:=arr_rows, _
                    RemainColumnCount:=arr_cols, _
                    ForcePasteIfArrayFormula:=True)
                ' 処理済み配列に記録
                Call pAddItemToStringArray(procd_arr, arr_anchor.Address)
            Else
                ' 通常の処理
                Call pCopyCellCore( _
                    src_bounds.Row + row_idx, src_bounds.Column + col_idx, _
                    dst_bounds.Row + row_idx, dst_bounds.Column + col_idx, _
                    src_sheet, dst_sheet, AsValue, CopyNumberFormat)
            End If
        Next col_idx
    Next row_idx
End Sub
Private Sub IWorksheetService_CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    Call CopyRange(SourceRangeBounds, DestinationRangeBounds, AsValue:=AsValue, CopyNumberFormat:=CopyNumberFormat)
End Sub

Private Sub pRaiseIfCopyRangeDestinationOutOfSheet( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds)

    Call pRaiseIfDestinationOutOfSheet(DestinationRangeBounds, SourceRangeBounds.RowCount, SourceRangeBounds.ColumnCount, "コピー")
End Sub

Private Sub pAddItemToStringArray(ByRef TargetArray() As String, ByVal ItemToBeAdded As String)
    Dim new_ubound As Long
    
    If (Not TargetArray) = -1 Then
        ReDim Preserve TargetArray(0 To 0) As String
    Else
        new_ubound = UBound(TargetArray) + 1
        ReDim Preserve TargetArray(LBound(TargetArray) To new_ubound) As String
    End If
    TargetArray(new_ubound) = ItemToBeAdded
End Sub

Private Function pContainsItemInStringArray(ByRef TargetArray() As String, ByVal CheckItem As String) As Boolean
    Dim arr_item As Variant
    For Each arr_item In TargetArray
        If arr_item = CheckItem Then
            pContainsItemInStringArray = True
            Exit Function
        End If
    Next arr_item
    pContainsItemInStringArray = False
End Function

Private Function pGetArrayFormulaAnchor(ByVal TargetSheet As Worksheet, ByVal RowIndex As Long, ByVal ColumnIndex As Long) As Range
    If TargetSheet.Cells(RowIndex, ColumnIndex).HasArray Then
        Set pGetArrayFormulaAnchor = TargetSheet.Cells(RowIndex, ColumnIndex).CurrentArray.Cells(1, 1)
    Else
        Set pGetArrayFormulaAnchor = Nothing
    End If
End Function

Private Function pIsArrayFormulaToBeProcessed(ByVal TargetCell As Range, ByRef ProcessedArray() As String) As Boolean
    If TargetCell Is Nothing Then
        pIsArrayFormulaToBeProcessed = False
        Exit Function
    End If
    
    If pContainsItemInStringArray(ProcessedArray, TargetCell.Address) Then
        pIsArrayFormulaToBeProcessed = False
    Else
        pIsArrayFormulaToBeProcessed = True
    End If
End Function

Private Sub pCalcurateArrayCount(ByRef ArrayRowCount As Long, ByRef ArrayColumnCount As Long, ByVal TargetArrayAnchor As Range, ByVal RangeBounds As WorksheetRangeBounds)
    Dim array_last_row As Long, array_last_col As Long
    array_last_row = TargetArrayAnchor.Row + TargetArrayAnchor.CurrentArray.Rows.Count - 1
    array_last_col = TargetArrayAnchor.Column + TargetArrayAnchor.CurrentArray.Columns.Count - 1
    
    Dim paste_start_row As Long, paste_start_col As Long
    paste_start_row = MaxLng(TargetArrayAnchor.Row, RangeBounds.Row)
    paste_start_col = MaxLng(TargetArrayAnchor.Column, RangeBounds.Column)
    
    Dim paste_finish_row As Long, paste_finish_col As Long
    paste_finish_row = MinLng(array_last_row, RangeBounds.FinishRow)
    paste_finish_col = MinLng(array_last_col, RangeBounds.FinishColumn)
    
    ArrayRowCount = MaxLng(paste_finish_row - paste_start_row + 1, 0)
    ArrayColumnCount = MaxLng(paste_finish_col - paste_start_col + 1, 0)
End Sub

Private Sub pCopyCellCore( _
        ByVal SourceRow As Long, ByVal SourceColumn As Long, _
        ByVal DestinationRow As Long, ByVal DestinationColumn As Long, _
        ByVal SourceSheet As Worksheet, ByVal DestinationSheet As Worksheet, _
        ByVal AsValue As Boolean, _
        ByVal CopyNumberFormat As Boolean, _
        Optional ByVal RemainRowCount As Long = 2147483647#, _
        Optional ByVal RemainColumnCount As Long = 2147483647#, _
        Optional ByVal ForcePasteIfArrayFormula As Boolean = False)
    
    If DestinationSheet Is Nothing Then Set DestinationSheet = SourceSheet
    
    Dim src_cell As Range
    Set src_cell = SourceSheet.Cells(SourceRow, SourceColumn)
    
    Dim dst_cell As Range
    Set dst_cell = DestinationSheet.Cells(DestinationRow, DestinationColumn)
    
    ' 値コピー
    If AsValue Then
        If IsEmpty(src_cell.Value) Then
            dst_cell.ClearContents
        Else
            dst_cell.Value = src_cell.Value
        End If
                
        If CopyNumberFormat Then dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
        Exit Sub
    End If
    
    ' 数式コピー
    If src_cell.HasFormula Then
        ' これだと相対参照が処理されない。(元のままとなってしまう)
        'dst_cell.Formula = src_cell.Formula
        If src_cell.HasArray Then
            If src_cell.Address = src_cell.CurrentArray.Cells(1, 1).Address Or ForcePasteIfArrayFormula Then
                Dim formula_str As String
                ' 一旦、元セルの通常の関数に入れる (正常に計算できなくともよい)
                formula_str = src_cell.FormulaArray
                src_cell = formula_str
                ' 宛先セルに入れる (FormulaR1C1 から取り出すことにより、なぜか相対参照が正しく処理される)
                If Application.ReferenceStyle = xlA1 Then
                    dst_cell = Application.ConvertFormula(src_cell.FormulaR1C1, xlR1C1, ToReferenceStyle:=xlA1)
                Else
                    dst_cell = Application.ConvertFormula(src_cell.FormulaR1C1, xlR1C1, ToReferenceStyle:=xlR1C1)
                End If
                ' 元セルを配列関数に戻す
                src_cell.FormulaArray = formula_str
                
                ' 宛先セルの通常関数を配列関数に入れる。
                formula_str = dst_cell.Formula
                Dim resized_dst As Range
                Set resized_dst = dst_cell.Resize( _
                        MinLng(src_cell.CurrentArray.Rows.Count, RemainRowCount), _
                        MinLng(src_cell.CurrentArray.Columns.Count, RemainColumnCount))
                resized_dst.FormulaArray = formula_str
                If CopyNumberFormat Then resized_dst.NumberFormatLocal = src_cell.NumberFormatLocal
            Else
                ' マルチセル配列の起点セル以外は何もしない
            End If
        Else
            Dim dst_number_format As String
            dst_number_format = dst_cell.NumberFormatLocal

            On Error GoTo RESTORE_FORMULA_COPY_NUMBER_FORMAT_ON_ERROR
            dst_cell.NumberFormat = "General"
            dst_cell.FormulaR1C1 = src_cell.FormulaR1C1
            If CopyNumberFormat Then
                dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
            Else
                dst_cell.NumberFormatLocal = dst_number_format
            End If
            On Error GoTo 0
        End If
    Else
        ' 値コピー
        If IsEmpty(src_cell.Value) Then
            dst_cell.ClearContents
        Else
            dst_cell.Value = src_cell.Value
        End If
        If CopyNumberFormat Then dst_cell.NumberFormatLocal = src_cell.NumberFormatLocal
    End If
    Exit Sub

RESTORE_FORMULA_COPY_NUMBER_FORMAT_ON_ERROR:
    Dim err_num As Long
    err_num = Err.Number

    Dim err_source As String
    err_source = Err.Source

    Dim err_desc As String
    err_desc = Err.Description

    On Error Resume Next
    dst_cell.NumberFormatLocal = dst_number_format
    On Error GoTo 0

    Err.Raise Number:=err_num, Source:=err_source, Description:=err_desc
End Sub

'
' セルが空かどうかを調べます。IgnoreEmptyString を True に指定した場合、数式が入っていても計算結果が空文字列であれば、空であると判定されます。
'
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    If RangeBounds.IsEmpty Then
        IsEmptyCell = True
        Exit Function
    End If
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    Dim cell_item As Range
    For Each cell_item In target_range
        If Not pIsEmptyCellCore(cell_item, IgnoreEmptyString:=IgnoreEmptyString) Then
            IsEmptyCell = False
            Exit Function
        End If
    Next cell_item
    
    IsEmptyCell = True
End Function
Private Function IWorksheetService_IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    IWorksheetService_IsEmptyCell = IsEmptyCell(RangeBounds, IgnoreEmptyString:=IgnoreEmptyString)
End Function

'
' セルに数式が入力されているかを調べます。
'
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    If RangeBounds.IsEmpty Then Exit Function
    If Not RangeBounds.IsCell Then Err.Raise vbObjectError + 1, "Class WorksheetService", "RangeBounds がセルではありません。(" & RangeBounds.ToString() & ")"
    HasFormula = pGetCell(RangeBounds).HasFormula
End Function
Private Function IWorksheetService_HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    IWorksheetService_HasFormula = HasFormula(RangeBounds)
End Function
Private Function pIsEmptyCellCore(ByVal TargetCell As Range, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    Dim cell_value As Variant
    cell_value = TargetCell.Value
    
    If IsEmpty(cell_value) Then
        pIsEmptyCellCore = True
        Exit Function
    End If
    
    If Not IgnoreEmptyString Then
        pIsEmptyCellCore = False
        Exit Function
    End If
    
    If VarType(cell_value) = vbString Then
        pIsEmptyCellCore = (Len(cell_value) = 0)
    Else
        pIsEmptyCellCore = False
    End If
End Function

'
' 指定した範囲内の、実際に使用されている範囲を返します。
'
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
    
    ' 空範囲はそのまま返す。
    If RangeBounds.IsEmpty Then
        Set GetUsedRangeBounds = RangeBounds
        Exit Function
    End If
    
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)
    
    ' 生の UsedRange を得る。
    Dim raw_range As WorksheetRangeBounds
    Set raw_range = pGetRawUsedRange(target_sheet, RangeBounds, IsUsedStartRow, IsUsedStartColumn)
    
    ' GetRawRange が True か、UsedRange の時点で空の場合はここで処理終了。
    If GetRawRange Or raw_range.IsEmpty Then
        Set GetUsedRangeBounds = raw_range
        Exit Function
    End If
    
    Dim row_start As Long, row_finish As Long
    Dim col_start As Long, col_finish As Long
    
    ' 最終行を得る。
    row_finish = pGetLastRowCore(target_sheet, raw_range.Row, raw_range.Column, raw_range.FinishRow, raw_range.FinishColumn, IgnoreEmpty)
    
    ' 最初の行を得る。
    If IsUsedStartRow Then
        row_start = RangeBounds.Row
    Else
        row_start = pGetFirstRowCore(target_sheet, raw_range.Row, raw_range.Column, row_finish, raw_range.FinishColumn, IgnoreEmpty)
    End If
    
    ' 最終列を得る。
    col_finish = pGetLastColumnCore(target_sheet, row_start, raw_range.Column, row_finish, raw_range.FinishColumn, IgnoreEmpty)
    
    ' 最初の列を得る。
    If IsUsedStartColumn Then
        col_start = RangeBounds.Column
    Else
        col_start = pGetFirstColumnCore(target_sheet, row_start, raw_range.Column, row_finish, col_finish, IgnoreEmpty)
    End If
    
    ' 処理結果を整形
    If row_finish < row_start Or col_finish < col_start Then
        Set GetUsedRangeBounds = New_RangeBounds(Row:=RangeBounds.Row, Column:=RangeBounds.Column, FinishRow:=0, FinishColumn:=0, Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    Else
        Set GetUsedRangeBounds = New_RangeBounds(Row:=row_start, Column:=col_start, FinishRow:=row_finish, FinishColumn:=col_finish, Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    End If
    
    ' 完全一致を解除する (通常使用時に完全一致は扱いづらい)
    Call pResetFindState(target_sheet)
End Function
Private Function IWorksheetService_GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
    
    Set IWorksheetService_GetUsedRangeBounds = GetUsedRangeBounds(RangeBounds, IgnoreEmpty:=IgnoreEmpty, IsUsedStartRow:=IsUsedStartRow, IsUsedStartColumn:=IsUsedStartColumn, GetRawRange:=GetRawRange)
End Function

Private Function pGetRawUsedRange( _
        ByVal TargetSheet As Worksheet, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal IsUsedStartRow As Boolean, _
        ByVal IsUsedStartColumn As Boolean) As WorksheetRangeBounds
    
    Dim used_s_row As Long, used_f_row As Long
    Dim used_s_col As Long, used_f_col As Long
    If 0 < TargetSheet.UsedRange.Rows.Count Then
        used_s_row = TargetSheet.UsedRange.Row
        used_f_row = TargetSheet.UsedRange.Rows(TargetSheet.UsedRange.Rows.Count).Row
    Else
        used_s_row = 1
        used_s_col = 1
    End If
    If 0 < TargetSheet.UsedRange.Columns.Count Then
        used_s_col = TargetSheet.UsedRange.Column
        used_f_col = TargetSheet.UsedRange.Columns(TargetSheet.UsedRange.Columns.Count).Column
    Else
        used_s_row = 1
        used_s_col = 1
    End If
    
    Dim used_bounds As WorksheetRangeBounds
    If used_s_row = used_f_row And used_s_col = used_f_col _
            And IsEmpty(TargetSheet.Cells(used_s_row, used_s_col).Value) _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeTop).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeBottom).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone _
            And TargetSheet.Cells(used_s_row, used_s_col).Borders(xlEdgeRight).LineStyle = xlLineStyleNone Then
        ' 空の領域
        Set used_bounds = New_RangeBounds( _
                Row:=used_s_row, _
                Column:=used_s_col, _
                FinishRow:=0, _
                FinishColumn:=0, _
                Sheet:=TargetSheet.Name, _
                Book:=TargetSheet.Parent.Name)
    Else
        ' 普通の領域
        Set used_bounds = New_RangeBounds( _
                Row:=used_s_row, _
                Column:=used_s_col, _
                FinishRow:=used_f_row, _
                FinishColumn:=used_f_col, _
                Sheet:=TargetSheet.Name, _
                Book:=TargetSheet.Parent.Name)
    End If
    
    ' 元の範囲との交差を取る
    Dim intersect_bounds As WorksheetRangeBounds
    Set intersect_bounds = RangeBounds.Intersect(used_bounds)
    
    If IsUsedStartRow And IsUsedStartColumn Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=RangeBounds.Row, Column:=RangeBounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    ElseIf IsUsedStartRow Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=RangeBounds.Row, Column:=intersect_bounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    ElseIf IsUsedStartColumn Then
        Set pGetRawUsedRange = New_RangeBounds( _
            Row:=intersect_bounds.Row, Column:=RangeBounds.Column, _
            FinishRow:=intersect_bounds.FinishRow, FinishColumn:=intersect_bounds.FinishColumn, _
            Sheet:=RangeBounds.WorksheetName, Book:=RangeBounds.WorkbookName)
    Else
        Set pGetRawUsedRange = intersect_bounds
    End If
End Function

Private Function pGetLastRowCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long
    
    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetLastRowCore = 0
        Exit Function
    End If
    
    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )
    
    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If
    
    Dim fnd_cell As Range
    ' 最終行 → SearchOrder=xlByRows, SearchDirection=xlPrevious
    ' After は検索範囲の左上セル
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(1), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)
    
    If fnd_cell Is Nothing Then
        pGetLastRowCore = 0
    Else
        pGetLastRowCore = fnd_cell.Row
    End If
End Function

Private Function pGetLastColumnCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long
    
    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetLastColumnCore = 0
        Exit Function
    End If
    
    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )
    
    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If
    
    Dim fnd_cell As Range
    ' 最終列 → SearchOrder=xlByColumns, SearchDirection=xlPrevious
    ' After は検索範囲の左上セル
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(1), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)
    
    If fnd_cell Is Nothing Then
        pGetLastColumnCore = 0
    Else
        pGetLastColumnCore = fnd_cell.Column
    End If
End Function

Private Function pGetFirstRowCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long
    
    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetFirstRowCore = 2147483647#
        Exit Function
    End If
    
    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )
    
    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If
    
    Dim fnd_cell As Range
    ' 最初の行 → SearchOrder=xlByRows, SearchDirection=xlNext
    ' After は検索範囲の右下セル
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(srch_rng.Rows.Count, srch_rng.Columns.Count), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlNext, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)
    
    If fnd_cell Is Nothing Then
        pGetFirstRowCore = 2147483647#
    Else
        pGetFirstRowCore = fnd_cell.Row
    End If
End Function

Private Function pGetFirstColumnCore( _
        ByVal TargetSheet As Worksheet, _
        ByVal MinRow As Long, _
        ByVal MinColumn As Long, _
        ByVal MaxRow As Long, _
        ByVal MaxColumn As Long, _
        ByVal IgnoreEmpty As Boolean) As Long
    
    If MaxRow < MinRow Or MaxColumn < MinColumn Then
        pGetFirstColumnCore = 2147483647#
        Exit Function
    End If
    
    Dim srch_rng As Range
    Set srch_rng = TargetSheet.Range( _
        TargetSheet.Cells(MinRow, MinColumn), _
        TargetSheet.Cells(MaxRow, MaxColumn) _
    )
    
    Dim look_in_mode As XlFindLookIn
    If IgnoreEmpty Then
        look_in_mode = xlValues
    Else
        look_in_mode = xlFormulas
    End If
    
    Dim fnd_cell As Range
    ' 最初の列 → SearchOrder=xlByColumns, SearchDirection=xlNext
    ' After は検索範囲の右下セル
    Set fnd_cell = srch_rng.Find(What:="*", _
                                 After:=srch_rng.Cells(srch_rng.Rows.Count, srch_rng.Columns.Count), _
                                 LookAt:=xlPart, _
                                 LookIn:=look_in_mode, _
                                 SearchOrder:=xlByColumns, _
                                 SearchDirection:=xlNext, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)
    
    If fnd_cell Is Nothing Then
        pGetFirstColumnCore = 2147483647#
    Else
        pGetFirstColumnCore = fnd_cell.Column
    End If
End Function

'
' 範囲を指定してクリアします。
'
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
    
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    ' ひとまず範囲全体を処理
    Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
    
    ' 使用中の範囲を取得
    Dim used_row As Long, used_col As Long
    If 0 < target_sheet.UsedRange.Rows.Count Then
        used_row = target_sheet.UsedRange.Rows(target_sheet.UsedRange.Rows.Count).Row
    End If
    If 0 < target_sheet.UsedRange.Columns.Count Then
        used_col = target_sheet.UsedRange.Columns(target_sheet.UsedRange.Columns.Count).Column
    End If
    
    Dim item_idx As Long
    For item_idx = RangeBounds.Row To MinLng(used_row, RangeBounds.FinishRow)
        If target_sheet.Rows(item_idx).Hidden Then
            ' 範囲内の非表示状態の行を処理
            Set target_range = target_sheet.Range(target_sheet.Cells(item_idx, RangeBounds.Column), target_sheet.Cells(item_idx, RangeBounds.FinishColumn))
            Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
        End If
    Next item_idx
    
    For item_idx = RangeBounds.Column To MinLng(used_col, RangeBounds.FinishColumn)
        If target_sheet.Columns(item_idx).Hidden Then
            ' 範囲内の非表示状態の列を処理
            Set target_range = target_sheet.Range(target_sheet.Cells(RangeBounds.Row, item_idx), target_sheet.Cells(RangeBounds.FinishRow, item_idx))
            Call pClearRangeCore(target_range, ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
        End If
    Next item_idx
End Sub
Private Sub IWorksheetService_ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
    
    Call ClearRange( _
            RangeBounds, ClearAll:=ClearAll, _
            ClearContents:=ClearContents, ClearNumberFormats:=ClearNumberFormats, ClearColors:=ClearColors, _
            ClearComments:=ClearComments, ClearHyperlinks:=ClearHyperlinks)
End Sub

Private Sub pClearRangeCore( _
        ByVal TargetRange As Range, _
        ByVal ClearAll As Boolean, ByVal ClearContents As Boolean, ByVal ClearNumberFormats As Boolean, _
        ByVal ClearColors As Boolean, ByVal ClearComments As Boolean, ByVal ClearHyperlinks As Boolean)
    
    If ClearAll Or ClearContents Then Call TargetRange.ClearContents
    If ClearAll Or ClearNumberFormats Then TargetRange.NumberFormat = "General"
    If ClearAll Or ClearColors Then
        TargetRange.Font.ColorIndex = 0
        TargetRange.Interior.ColorIndex = xlColorIndexNone
    End If
    If ClearAll Or ClearComments Then Call TargetRange.ClearComments
    If ClearAll Or ClearHyperlinks Then Call TargetRange.ClearHyperlinks
End Sub

'
' 範囲内の重複行を削除します。
'
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    If RangeBounds.IsEmpty Then Exit Sub

    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)

    Call target_range.RemoveDuplicates(Columns:=pBuildDuplicateColumnsArgument(RangeBounds, DuplicateColumns), Header:=xlNo)
End Sub
Private Sub IWorksheetService_RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call RemoveDuplicates(RangeBounds, DuplicateColumns)
End Sub

'
' 行を挿入します。
'
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call pRaiseIfNotEntireRowRange(RangeBounds, "InsertRows")

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Call target_sheet.Range(CStr(RangeBounds.Row) & ":" & CStr(RangeBounds.FinishRow)).Insert
End Sub
Private Sub IWorksheetService_InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call InsertRows(RangeBounds)
End Sub

'
' 行を削除します。
'
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call pRaiseIfNotEntireRowRange(RangeBounds, "DeleteRows")

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    Call target_sheet.Range(CStr(RangeBounds.Row) & ":" & CStr(RangeBounds.FinishRow)).Delete
End Sub
Private Sub IWorksheetService_DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call DeleteRows(RangeBounds)
End Sub

'
' すべてのデータを表示します。
' Sheet.ShowAllData、Rows.Hidden = False、Columns.Hidden = False します。
'
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)
    
    If target_sheet.FilterMode Then
        Call target_sheet.ShowAllData
    End If
    target_sheet.Rows.Hidden = False
    target_sheet.Columns.Hidden = False
End Sub
Private Sub IWorksheetService_SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call SetAllDataVisible(RangeBounds)
End Sub

'
' 構成済みのアウトラインに対して、表示するアウトライン レベルを設定します。
'
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    If RowLevels = 0 And ColumnLevels = 0 Then Exit Sub

    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)

    If RowLevels > 0 And ColumnLevels > 0 Then
        Call target_sheet.Outline.ShowLevels(RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
    ElseIf RowLevels > 0 Then
        Call target_sheet.Outline.ShowLevels(RowLevels:=RowLevels)
    Else
        Call target_sheet.Outline.ShowLevels(ColumnLevels:=ColumnLevels)
    End If
End Sub
Private Sub IWorksheetService_SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call SetSheetOutlineLevel(RangeBounds, RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
End Sub

Private Function pBuildDuplicateColumnsArgument(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant) As Variant
    If IsArray(DuplicateColumns) Then
        If IsEmptyArray(DuplicateColumns) Then
            Err.Raise vbObjectError + 1, "Class WorksheetService", "DuplicateColumns には 1 つ以上の列番号を指定してください。"
        End If

        Dim result_columns() As Variant
        Dim column_count As Long
        Dim column_item As Variant
        For Each column_item In DuplicateColumns
            If column_count = 0 Then
                ReDim result_columns(0 To 0)
            Else
                ReDim Preserve result_columns(0 To column_count)
            End If
            result_columns(column_count) = pGetValidatedRelativeColumnIndex(RangeBounds, column_item, "DuplicateColumns")
            column_count = column_count + 1
        Next column_item

        If column_count = 0 Then
            Err.Raise vbObjectError + 1, "Class WorksheetService", "DuplicateColumns には 1 つ以上の列番号を指定してください。"
        End If

        pBuildDuplicateColumnsArgument = result_columns
    Else
        pBuildDuplicateColumnsArgument = pGetValidatedRelativeColumnIndex(RangeBounds, DuplicateColumns, "DuplicateColumns")
    End If
End Function

Private Function pGetValidatedRelativeColumnIndex(ByVal RangeBounds As WorksheetRangeBounds, ByVal ColumnIndex As Variant, ByVal ArgumentName As String) As Long
    Dim column_num As Long
    column_num = pGetValidatedLong(ColumnIndex, ArgumentName)
    If column_num < 1 Or RangeBounds.ColumnCount < column_num Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", ArgumentName & " は対象範囲内の相対列番号を指定してください。"
    End If

    pGetValidatedRelativeColumnIndex = column_num
End Function

Private Function pGetValidatedLong(ByVal TargetValue As Variant, ByVal ArgumentName As String) As Long
    On Error GoTo INVALID_VALUE
    If IsObject(TargetValue) Then GoTo INVALID_VALUE
    If IsError(TargetValue) Then GoTo INVALID_VALUE
    If IsEmpty(TargetValue) Then GoTo INVALID_VALUE
    If IsNull(TargetValue) Then GoTo INVALID_VALUE
    If Not IsNumeric(TargetValue) Then GoTo INVALID_VALUE

    Dim result_value As Long
    result_value = CLng(TargetValue)
    If CDbl(TargetValue) <> CDbl(result_value) Then GoTo INVALID_VALUE

    pGetValidatedLong = result_value
    On Error GoTo 0
    Exit Function

INVALID_VALUE:
    On Error GoTo 0
    Err.Raise vbObjectError + 1, "Class WorksheetService", ArgumentName & " には整数を指定してください。"
End Function

Private Sub pRaiseIfNotEntireRowRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal MethodName As String)
    If RangeBounds.IsEmpty Or Not RangeBounds.IsEntireRow Then
        Err.Raise vbObjectError + 1, "Class WorksheetService", MethodName & " には行全体の範囲を指定してください。"
    End If
End Sub

'
' シートのタブ色を設定します。インデックスと値を同時に指定した場合、値指定が優先されます。
'
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
    Dim target_sheet As Worksheet
    Set target_sheet = pGetWorksheet(RangeBounds)
    
    If TabColor >= 0 Then
        target_sheet.Tab.Color = TabColor
    Else
        target_sheet.Tab.ColorIndex = TabColorIndex
    End If
End Sub
Private Sub IWorksheetService_SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
    
    Call SetSheetTabColor(RangeBounds, TabColorIndex:=TabColorIndex, TabColor:=TabColor)
End Sub

'
' 範囲の色を設定します。インデックスと値を同時に指定した場合、値指定が優先されます。
'
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    If FontColor >= 0 Then
        target_range.Font.Color = FontColor
    Else
        target_range.Font.ColorIndex = FontColorIndex
    End If
    
    If InteriorColor >= 0 Then
        target_range.Interior.Color = InteriorColor
    Else
        target_range.Interior.ColorIndex = InteriorColorIndex
    End If
End Sub
Private Sub IWorksheetService_SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
    
    Call SetRangeColor(RangeBounds, FontColorIndex:=FontColorIndex, FontColor:=FontColor, InteriorColorIndex:=InteriorColorIndex, InteriorColor:=InteriorColor)
End Sub

'
' テキストを折り返して表示するかを設定する。
'
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    target_range.WrapText = WrapText
End Sub
Private Sub IWorksheetService_SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call SetWrapText(RangeBounds, WrapText)
End Sub

'
' テキストを縮小して全体を表示するかを設定する。
'
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    target_range.ShrinkToFit = ShrinkToFit
End Sub
Private Sub IWorksheetService_SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call SetShrinkToFit(RangeBounds, ShrinkToFit)
End Sub

'
' セル内の配置を設定する。
'
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = xlGeneral, _
        Optional ByVal VerticalAlignment As Long = xlBottom, _
        Optional ByVal Orientation As Long = 0, _
        Optional ByVal IndentLevel As Long = 0)
    If RangeBounds.IsEmpty Then Exit Sub
    
    Dim target_range As Range
    Set target_range = pGetRange(RangeBounds)
    
    target_range.Orientation = Orientation ' 文字の方向
    target_range.IndentLevel = IndentLevel ' インデント
    
    Select Case HorizontalAlignment
     Case xlLeft ' 左詰め
        target_range.HorizontalAlignment = xlLeft
     Case xlCenter ' 中央揃え
        target_range.HorizontalAlignment = xlCenter
     Case xlRight ' 右詰め
        target_range.HorizontalAlignment = xlRight
     Case xlFill ' 繰り返し
        target_range.HorizontalAlignment = xlFill
     Case xlJustify ' 両端揃え
        target_range.HorizontalAlignment = xlJustify
     Case xlCenterAcrossSelection ' 選択範囲内で中央
        target_range.HorizontalAlignment = xlCenterAcrossSelection
     Case xlDistributed ' 均等割り付け
        target_range.HorizontalAlignment = xlDistributed
     Case Else ' 標準
        target_range.HorizontalAlignment = xlGeneral
    End Select
    
    Select Case VerticalAlignment
     Case xlTop ' 上詰め
        target_range.VerticalAlignment = xlTop
     Case xlCenter ' 中央揃え
        target_range.VerticalAlignment = xlCenter
     Case xlFill ' 繰り返し
        target_range.VerticalAlignment = xlFill
     Case xlJustify ' 両端揃え
        target_range.VerticalAlignment = xlJustify
     Case xlDistributed ' 均等割り付け
        target_range.VerticalAlignment = xlDistributed
     Case Else ' 下詰め
        target_range.VerticalAlignment = xlBottom
    End Select
End Sub
Private Sub IWorksheetService_SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = xlGeneral, _
        Optional ByVal VerticalAlignment As Long = xlBottom, _
        Optional ByVal Orientation As Long = 0, _
        Optional ByVal IndentLevel As Long = 0)
    
    Call SetAlignment( _
            RangeBounds, _
            HorizontalAlignment:=HorizontalAlignment, _
            VerticalAlignment:=VerticalAlignment, _
            Orientation:=Orientation, _
            IndentLevel:=IndentLevel)
End Sub
