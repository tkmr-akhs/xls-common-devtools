Attribute VB_Name = "Test_WorksheetService"
Option Explicit

' #############################################################################
'!
'! @brief
'! WorksheetService クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private Function pPrepareTestSheet(ByRef SheetName As String) As Worksheet
    Dim target_sheet As Worksheet

    On Error Resume Next
        Set target_sheet = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If Not target_sheet Is Nothing Then
        target_sheet.Rows.Hidden = False
        target_sheet.Columns.Hidden = False
        Call target_sheet.Cells.ClearOutline
        Call target_sheet.Cells.Clear
    Else
        Set target_sheet = ThisWorkbook.Worksheets.Add
        target_sheet.Name = SheetName
    End If

    Set pPrepareTestSheet = target_sheet
End Function

Private Sub pSetStaleFindFormat(ByVal TargetRange As Range)
    Application.FindFormat.Clear
    Application.FindFormat.Interior.Color = RGB(255, 0, 0)

    Dim ignored_cell As Range
    Set ignored_cell = TargetRange.Find( _
            What:="__codex_not_found__", _
            After:=TargetRange.Cells(1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=True)
End Sub

Private Sub pClearStaleFindFormat(ByVal TargetRange As Range)
    Application.FindFormat.Clear

    Dim ignored_cell As Range
    Set ignored_cell = TargetRange.Find( _
            What:="__codex_not_found__", _
            After:=TargetRange.Cells(1), _
            LookIn:=xlValues, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, _
            MatchCase:=False, _
            SearchFormat:=False)
End Sub

Private Sub pWriteEmptyStringValue(ByVal TargetCell As Range)
    TargetCell.Formula = "="""""
    Call TargetCell.Copy
    Call TargetCell.PasteSpecial(Paste:=xlPasteValues, SkipBlanks:=False)
    Application.CutCopyMode = False
End Sub

' -----------------------------------------------------------------------------
' ActivateRange
' -----------------------------------------------------------------------------

Public Sub Test_ActivateRange_NonActiveWorkbook_ActivatesTargetBookSheetAndCell(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_start"
    target_book.Worksheets(2).Name = "target_activate"
    Call target_book.Worksheets("target_start").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=4, Sheet:="target_activate", Book:=target_book_name)

    ' Act
    Err.Clear
    Call sheet_srv.ActivateRange(range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_active_cell_address As String
    If actual_error_number = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_active_cell_address = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals target_book_name, actual_active_book_name
    Assert.Equals "target_activate", actual_active_sheet_name
    Assert.Equals "D3", actual_active_cell_address
End Sub

Public Sub Test_ActivateRange_NonActiveWorksheet_ActivatesTargetSheetAndCell(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_start"
    target_book.Worksheets(2).Name = "target_activate"
    Call target_book.Worksheets("target_start").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=4, Column:=5, Sheet:="target_activate", Book:=target_book_name)

    ' Act
    Err.Clear
    Call sheet_srv.ActivateRange(range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_active_cell_address As String
    If actual_error_number = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_active_cell_address = ActiveCell.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals target_book_name, actual_active_book_name
    Assert.Equals "target_activate", actual_active_sheet_name
    Assert.Equals "E4", actual_active_cell_address
End Sub

' -----------------------------------------------------------------------------
' Find
' -----------------------------------------------------------------------------

Public Sub Test_Find_EmptyRange_ReturnsEmptyResult(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = "target"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, FinishRow:=2, FinishColumn:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_ranges() As WorksheetRangeBounds
    actual_ranges = sheet_srv.Find("target", range_bounds)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue IsEmptyArray(actual_ranges)
End Sub

Public Sub Test_Find_IgnoresPreviousSearchFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = "target"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Call pSetStaleFindFormat(target_sheet.Range("A1:C3"))

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_ranges() As WorksheetRangeBounds
    actual_ranges = sheet_srv.Find("target", range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    Call pClearStaleFindFormat(target_sheet.Range("A1:C3"))

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim actual_is_empty As Boolean
    actual_is_empty = IsEmptyArray(actual_ranges)
    Assert.IsFalse actual_is_empty
    If actual_is_empty Then Exit Sub

    Assert.EqualsNumeric 0, LBound(actual_ranges)
    Assert.EqualsNumeric 0, UBound(actual_ranges)
    Assert.EqualsNumeric 2, actual_ranges(0).Row
    Assert.EqualsNumeric 2, actual_ranges(0).Column
End Sub

Public Sub Test_Find_PreviousSearchOrder_DoesNotChangeResultOrder(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = "target"
    target_sheet.Cells(2, 4).Value = "target"
    target_sheet.Cells(4, 2).Value = "target"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=5, FinishColumn:=5, Sheet:="test_output")

    Dim ignored_cell As Range
    Set ignored_cell = target_sheet.Range("A1:E5").Find( _
            What:="target", _
            After:=target_sheet.Range("A1"), _
            LookIn:=xlValues, _
            LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious, _
            MatchCase:=False, _
            SearchFormat:=False)

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_ranges() As WorksheetRangeBounds
    actual_ranges = sheet_srv.Find("target", range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    Call pClearStaleFindFormat(target_sheet.Range("A1:E5"))

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Dim actual_is_empty As Boolean
    actual_is_empty = IsEmptyArray(actual_ranges)
    Assert.IsFalse actual_is_empty
    If actual_is_empty Then Exit Sub

    Assert.EqualsNumeric 0, LBound(actual_ranges)
    Assert.EqualsNumeric 2, UBound(actual_ranges)
    Assert.EqualsNumeric 2, actual_ranges(0).Row
    Assert.EqualsNumeric 2, actual_ranges(0).Column
    Assert.EqualsNumeric 2, actual_ranges(1).Row
    Assert.EqualsNumeric 4, actual_ranges(1).Column
    Assert.EqualsNumeric 4, actual_ranges(2).Row
    Assert.EqualsNumeric 2, actual_ranges(2).Column
End Sub

' -----------------------------------------------------------------------------
' Sort
' -----------------------------------------------------------------------------

Public Sub Test_Sort_WhenCalls_HasBeenSorted(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(2, 2).Value = 2
    target_sheet.Cells(2, 3).Value = 22
    target_sheet.Cells(3, 2).Value = 1
    target_sheet.Cells(3, 3).Value = 111
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.Sort(range_bounds, 1, xlAscending)
    
    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 111, target_sheet.Cells(2, 3).Value
    Assert.EqualsNumeric 22, target_sheet.Cells(3, 3).Value
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Sort_EmptyRange_DoesNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = 2
    target_sheet.Cells(2, 3).Value = 22

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, FinishRow:=2, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.Sort(range_bounds, 1, xlAscending)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, target_sheet.Cells(2, 2).Value
    Assert.EqualsNumeric 22, target_sheet.Cells(2, 3).Value
End Sub
Public Sub Test_Sort_KeyColumnZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.Sort(range_bounds, 0, xlAscending)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

Public Sub Test_Sort_KeyColumnPastRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.Sort(range_bounds, 3, xlAscending)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

Public Sub Test_Sort_InvalidOrder_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.Sort(range_bounds, 1, 9999)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

' -----------------------------------------------------------------------------
' WriteCell
' -----------------------------------------------------------------------------

Public Sub Test_WriteCell_String_WritesAsString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "Hello VBA")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(2, 2).Value

    Assert.Equals "Hello VBA", actual_value
End Sub


Public Sub Test_WriteCell_UnconvertibleString_DoesNotLeaveErrNumber(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=13, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Call sheet_srv.WriteCell(range_bounds, "ABC123")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "ABC123", target_sheet.Cells(13, 2).Value
End Sub

Public Sub Test_WriteCell_NumericString_WritesAsNumber(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "12345")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(3, 2).Value

    Assert.EqualsNumeric 12345, actual_value
End Sub

Public Sub Test_WriteCell_DecimalString_WritesAsDouble(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=8, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "123.45")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(8, 2).Value

    Assert.EqualsNumeric 123.45, actual_value
End Sub

Public Sub Test_WriteCell_NumericStringAndConvertFalse_WritesAsString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=6, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "ABC123", TypeConvert:=False)

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(6, 2).Value

    Assert.Equals "ABC123", actual_value
    Assert.Equals "String", TypeName(actual_value)
End Sub

Public Sub Test_WriteCell_TextWriteError_RestoresNumberFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("B7").NumberFormatLocal = "0.0000"
    target_sheet.Protect Contents:=True, AllowFormattingCells:=True

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=7, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.WriteCell range_bounds, "ABC123", TypeConvert:=False

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    Dim actual_format As String
    actual_format = target_sheet.Range("B7").NumberFormatLocal

    target_sheet.Unprotect
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "0.0000", actual_format
End Sub

Public Sub Test_WriteCell_TextWriteError_RestoresCellContents(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim original_enable_events As Boolean
    original_enable_events = Application.EnableEvents
    Application.EnableEvents = True

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("B8").Formula = "=1+1"
    target_sheet.Range("B8").NumberFormatLocal = "0.0000"
    target_sheet.Range("Z1").ClearContents

    Dim protect_probe As WorksheetChangeTestProbe
    Set protect_probe = New WorksheetChangeTestProbe
    Call protect_probe.Initialize(target_sheet, "B8", "Z1")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=8, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.WriteCell range_bounds, "ABC123", TypeConvert:=False

    Dim actual_error_number As Long
    actual_error_number = Err.Number

    Dim actual_error_source As String
    actual_error_source = Err.Source

    Dim actual_error_description As String
    actual_error_description = Err.Description

    If target_sheet.ProtectContents Then target_sheet.Unprotect
    Set protect_probe = Nothing
    Application.EnableEvents = original_enable_events
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "=1+1", target_sheet.Range("B8").Formula
    Assert.EqualsNumeric 2, target_sheet.Range("B8").Value
    Assert.Equals "0.0000", target_sheet.Range("B8").NumberFormatLocal
End Sub

Public Sub Test_WriteCell_DateString_WritesAsDate(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=4, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "2025/01/01")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(4, 2).Value

    Assert.Equals CDate("2025/01/01"), actual_value
End Sub

Public Sub Test_WriteCell_BooleanString_WritesAsBoolean(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=5, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "True")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(5, 2).Value

    Assert.Equals True, actual_value
End Sub

Public Sub Test_WriteCell_AsFormula_WritesAsFormula(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=7, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "=SUM(1,2,3)", AsFormula:=True)

    ' Assert
    Dim actual_formula As String
    actual_formula = target_sheet.Cells(7, 2).Formula  ' =SUM(1,2,3)
    Assert.Equals "=SUM(1,2,3)", actual_formula

    Dim actual_value As Variant
    actual_value = target_sheet.Cells(7, 2).Value      ' = 6
    Assert.EqualsNumeric 1 + 2 + 3, actual_value
End Sub

Public Sub Test_WriteCell_ClearWhenEmpty_ClearsCell(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(9, 2).Value = "InitialValue"  ' もともと何か入っている

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=9, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "", ClearWhenEmpty:=True)

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(9, 2).Value
    ' ClearContents されている → Empty
    Assert.IsTrue IsEmpty(actual_value)
End Sub

Public Sub Test_WriteCell_NotClearWhenEmpty_WritesEmptyString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(10, 2).Value = "InitialValue"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=10, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    ' 空を渡すが ClearWhenEmpty=False なので "" を書き込み
    Call sheet_srv.WriteCell(range_bounds, "", ClearWhenEmpty:=False)

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(10, 2).Value

    ' 空文字列の値になっているか
    Assert.Equals "", actual_value
    Assert.IsFalse IsEmpty(actual_value)
    Assert.IsFalse target_sheet.Cells(10, 2).HasFormula
    Assert.Equals "String", TypeName(actual_value)
End Sub

Public Sub Test_WriteCell_IgnoreEmpty_NotChange(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(11, 2).Value = "KeepValue"
    target_sheet.Cells(11, 2).NumberFormatLocal = "0.00"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=11, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    ' 空を渡す & IgnoreEmpty=True → なにもしない
    Call sheet_srv.WriteCell(range_bounds, "", NumberFormat:="@", IgnoreEmpty:=True)

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(11, 2).Value
    ' "KeepValue" がそのまま残る
    Assert.Equals "KeepValue", actual_value
    Assert.IsFalse target_sheet.Cells(11, 2).HasFormula
    Assert.Equals "0.00", target_sheet.Cells(11, 2).NumberFormatLocal
End Sub

Public Sub Test_WriteCell_IgnoreEmpty_KeepsFormulaAndNumberFormat(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).Formula = "=1+2"
    target_sheet.Cells(12, 2).NumberFormatLocal = "0.00"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "", NumberFormat:="@", IgnoreEmpty:=True)

    ' Assert
    Assert.IsTrue target_sheet.Cells(12, 2).HasFormula
    Assert.Equals "=1+2", target_sheet.Cells(12, 2).Formula
    Assert.Equals "0.00", target_sheet.Cells(12, 2).NumberFormatLocal
End Sub

Public Sub Test_WriteCell_NumberFormat_SetsNumberFormat(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteCell(range_bounds, "23", NumberFormat:="m/dd")

    ' Assert
    Dim actual_value As String
    actual_value = target_sheet.Cells(12, 2).Text
    Assert.Equals "1/23", actual_value
End Sub

Public Sub Test_WriteCell_ErrorString_CreatesError(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=20, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    ' "#DIV/0!" を書き込む -> セルには CVErr(xlErrDiv0) が代入されるはず
    Call sheet_srv.WriteCell(range_bounds, "#DIV/0!")

    ' Assert
    Dim actual_value As Variant
    actual_value = target_sheet.Cells(20, 2).Value

    ' セルがエラー値になっているか
    Assert.Equals CVErr(xlErrDiv0), actual_value
End Sub

Public Sub Test_WriteCell_ModernErrorStrings_CreateErrors(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim error_texts As Variant
    error_texts = Array( _
            "#GETTING_DATA", _
            "#SPILL!", _
            "#CONNECT!", _
            "#BLOCKED!", _
            "#UNKNOWN!", _
            "#FIELD!", _
            "#CALC!")

    Dim error_values As Variant
    error_values = Array( _
            CVErr(xlErrGettingData), _
            CVErr(xlErrSpill), _
            CVErr(xlErrConnect), _
            CVErr(xlErrBlocked), _
            CVErr(xlErrUnknown), _
            CVErr(xlErrField), _
            CVErr(xlErrCalc))

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    Dim err_idx As Long
    For err_idx = LBound(error_texts) To UBound(error_texts)
        Dim range_bounds As WorksheetRangeBounds
        Set range_bounds = New_RangeBounds(Row:=22 + err_idx, Column:=2, Sheet:="test_output")

        ' Act
        Err.Clear
        Call sheet_srv.WriteCell(range_bounds, CStr(error_texts(err_idx)))

        Dim actual_error_number As Long
        Dim actual_error_source As String
        Dim actual_error_description As String
        actual_error_number = Err.Number
        actual_error_source = Err.Source
        actual_error_description = Err.Description

        ' Assert
        If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
            On Error GoTo 0
            Exit Sub
        End If
        Assert.Equals error_values(err_idx), target_sheet.Cells(22 + err_idx, 2).Value
    Next err_idx

    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' WriteRange
' -----------------------------------------------------------------------------

Public Sub Test_WriteRange_Range_WritesValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim values_arr(1 To 3, 1 To 3) As Variant
    Dim expect_arr(1 To 3, 1 To 3) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 3
        For idx2 = 1 To 3
            values_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
            expect_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
        Next idx2
    Next idx1
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Call sheet_srv.WriteRange(range_bounds, values_arr)
    
    Dim actual_arr() As Variant
    actual_arr = target_sheet.Range("C2:E4").Value
    
    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_WriteRange_Cell_WritesValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim values_arr(1 To 1, 1 To 1) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 1
        For idx2 = 1 To 1
            values_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
        Next idx2
    Next idx1
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=2, FinishColumn:=3, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Call sheet_srv.WriteRange(range_bounds, values_arr)
    
    Dim actual_value As Variant
    actual_value = target_sheet.Range("C2").Value
    
    ' Assert
    Assert.Equals "idx: 1-1", actual_value
End Sub

Public Sub Test_WriteRange_CellDestination_ResizesToArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim values_arr(1 To 2, 1 To 3) As Variant
    Dim expect_arr(1 To 2, 1 To 3) As Variant
    Dim row_idx As Long, col_idx As Long
    For row_idx = 1 To 2
        For col_idx = 1 To 3
            values_arr(row_idx, col_idx) = "cell: " & row_idx & "-" & col_idx
            expect_arr(row_idx, col_idx) = "cell: " & row_idx & "-" & col_idx
        Next col_idx
    Next row_idx

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteRange(range_bounds, values_arr)

    Dim actual_arr() As Variant
    actual_arr = target_sheet.Range("C2:E3").Value

    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_WriteRange_EmptyDestination_UsesAnchor(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim values_arr(1 To 2, 1 To 2) As Variant
    Dim expect_arr(1 To 2, 1 To 2) As Variant
    values_arr(1, 1) = "A"
    values_arr(1, 2) = "B"
    values_arr(2, 1) = "C"
    values_arr(2, 2) = "D"
    expect_arr(1, 1) = "A"
    expect_arr(1, 2) = "B"
    expect_arr(2, 1) = "C"
    expect_arr(2, 2) = "D"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=0, FinishColumn:=0, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteRange(range_bounds, values_arr)

    Dim actual_arr() As Variant
    actual_arr = target_sheet.Range("C2:D3").Value

    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_WriteRange_MismatchedDestination_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("C2:D4").Value = "keep"

    Dim values_arr(1 To 2, 1 To 2) As Variant
    values_arr(1, 1) = "A"
    values_arr(1, 2) = "B"
    values_arr(2, 1) = "C"
    values_arr(2, 2) = "D"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=4, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.WriteRange(range_bounds, values_arr)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("C2").Value
    Assert.Equals "keep", target_sheet.Range("D4").Value
End Sub

' -----------------------------------------------------------------------------
' ReadCell
' -----------------------------------------------------------------------------

Public Sub Test_ReadCell_Call_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).NumberFormatLocal = "m/dd"
    target_sheet.Cells(12, 2).Value = "92"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As String
    Dim actual_format As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, NumberFormat:=actual_format)

    ' Assert
    Assert.Equals "1900/04/01", actual_value
    Assert.Equals "m/dd", actual_format
End Sub

Public Sub Test_ReadCell_CallWithGetText_ReturnsText1(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).NumberFormatLocal = "m/dd"
    target_sheet.Cells(12, 2).Value = "92"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As String
    Dim actual_format As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, NumberFormat:=actual_format, GetText:=True)

    ' Assert
    Assert.Equals "4/01", actual_value
    Assert.Equals "m/dd", actual_format
End Sub

Public Sub Test_ReadCell_CallWithGetText_ReturnsText2(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).NumberFormatLocal = "#,##0_ "
    target_sheet.Cells(12, 2).Value = "12345678"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As String
    Dim actual_format As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, NumberFormat:=actual_format, GetText:=True)

    ' Assert
    Assert.Equals "12,345,678 ", actual_value
    Assert.Equals "#,##0_ ", actual_format
End Sub

'Public Sub Test_ReadCell_EmptyCell_ReturnsEmpty(ByVal Assert As UnitTestAssert)
'    ' Arrange
'    Dim target_sheet As Worksheet
'    Set target_sheet = pPrepareTestSheet("test_output")
'
'    ' 該当セルをクリア
'    target_sheet.Cells(13, 2).ClearContents
'
'    Dim range_bounds As WorksheetRangeBounds
'    Set range_bounds = New_RangeBounds(Row:=13, Column:=2, Sheet:="test_output")
'
'    Dim sheet_srv As IWorksheetService
'    Set sheet_srv = New WorksheetService
'
'    ' Act (Value モードで読み取り)
'    Dim actual_value As String
'    Call sheet_srv.ReadCell(range_bounds, actual_value, GetText:=False)
'
'    ' Assert
'    ' 空セル → Value は Empty、Text も "" になる
'    Assert.IsTrue IsEmpty(actual_value))
'End Sub


Public Sub Test_ReadCell_GetText_PreservesFractionalColumnWidth(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).Value = DateSerial(2026, 5, 15)
    target_sheet.Cells(12, 2).NumberFormatLocal = "yyyy/mm/dd"
    target_sheet.Columns(2).ColumnWidth = 4.71

    Dim expected_width As Double
    expected_width = target_sheet.Columns(2).ColumnWidth

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, GetText:=True)

    Dim actual_width As Double
    actual_width = target_sheet.Columns(2).ColumnWidth

    ' Assert
    Assert.EqualsNumeric expected_width, actual_width
End Sub

Public Sub Test_ReadCell_GetText_ReturnsNumberFormatLocal(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(12, 2).Value = 123

    Dim expected_format As String
    expected_format = target_sheet.Cells(12, 2).NumberFormatLocal

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=12, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As String
    Dim actual_format As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, NumberFormat:=actual_format, GetText:=True)

    ' Assert
    Assert.Equals expected_format, actual_format
End Sub

Public Sub Test_ReadCell_EmptyCellWithGetText_ReturnsEmptyString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' 該当セルをクリア
    target_sheet.Cells(13, 2).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=13, Column:=2, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_value As String
    sheet_srv.ReadCell range_bounds, actual_value, GetText:=True
    
    ' Assert
    Assert.Equals "", actual_value
End Sub

Public Sub Test_ReadCell_ErrorCell_ReturnsErrorValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(15, 2).Formula = "=1/0" ' #DIV/0! を作る
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=15, Column:=2, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Value モード
    Dim actual_value As String
    Call sheet_srv.ReadCell(range_bounds, actual_value, GetText:=False)
    
    ' Assert
    Assert.Equals "#DIV/0!", actual_value
End Sub

Public Sub Test_ReadCell_ErrorCell_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' #DIV/0! を作る
    target_sheet.Cells(16, 2).Formula = "=1/0"
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=16, Column:=2, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_text As String
    Call sheet_srv.ReadCell(range_bounds, actual_text, GetText:=True)
    
    ' Assert
    ' #DIV/0! は Text プロパティなら "#DIV/0!" の文字列になる
    Assert.Equals "#DIV/0!", actual_text
End Sub

Public Sub Test_ReadCell_AllExcelErrorValues_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim error_values As Variant
    error_values = Array( _
            CVErr(xlErrDiv0), _
            CVErr(xlErrNA), _
            CVErr(xlErrName), _
            CVErr(xlErrNull), _
            CVErr(xlErrNum), _
            CVErr(xlErrRef), _
            CVErr(xlErrValue), _
            CVErr(xlErrGettingData), _
            CVErr(xlErrSpill), _
            CVErr(xlErrConnect), _
            CVErr(xlErrBlocked), _
            CVErr(xlErrUnknown), _
            CVErr(xlErrField), _
            CVErr(xlErrCalc))

    Dim error_texts As Variant
    error_texts = Array( _
            "#DIV/0!", _
            "#N/A", _
            "#NAME?", _
            "#NULL!", _
            "#NUM!", _
            "#REF!", _
            "#VALUE!", _
            "#GETTING_DATA", _
            "#SPILL!", _
            "#CONNECT!", _
            "#BLOCKED!", _
            "#UNKNOWN!", _
            "#FIELD!", _
            "#CALC!")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    Dim err_idx As Long
    For err_idx = LBound(error_values) To UBound(error_values)
        target_sheet.Cells(20 + err_idx, 2).Value = error_values(err_idx)

        Dim range_bounds As WorksheetRangeBounds
        Set range_bounds = New_RangeBounds(Row:=20 + err_idx, Column:=2, Sheet:="test_output")

        ' Act
        Err.Clear
        Dim actual_text As String
        Call sheet_srv.ReadCell(range_bounds, actual_text, GetText:=False)

        Dim actual_error_number As Long
        Dim actual_error_source As String
        Dim actual_error_description As String
        actual_error_number = Err.Number
        actual_error_source = Err.Source
        actual_error_description = Err.Description

        ' Assert
        If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
            On Error GoTo 0
            Exit Sub
        End If
        Assert.Equals CStr(error_texts(err_idx)), actual_text
    Next err_idx

    On Error GoTo 0
End Sub

Public Sub Test_ReadCell_GetTextNarrowErrorCell_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(30, 2).Value = CVErr(xlErrDiv0)
    target_sheet.Columns(2).ColumnWidth = 4

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=30, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_text As String
    Call sheet_srv.ReadCell(range_bounds, actual_text, GetText:=True)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "#DIV/0!", actual_text
End Sub

Public Sub Test_ReadCell_DynamicArrayCalcError_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(32, 2).Formula2 = "=FILTER({1,2},{FALSE,FALSE})"
    Application.CalculateFull

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=32, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_text As String
    Call sheet_srv.ReadCell(range_bounds, actual_text, GetText:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "#CALC!", actual_text
End Sub

Public Sub Test_ReadCell_SpillError_ReturnsErrorText(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(34, 3).Value = "blocked"
    target_sheet.Cells(34, 2).Formula2 = "=SEQUENCE(1,2)"
    Application.CalculateFull

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=34, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_text As String
    Call sheet_srv.ReadCell(range_bounds, actual_text, GetText:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals "#SPILL!", actual_text
End Sub

' -----------------------------------------------------------------------------
' ReadRange
' -----------------------------------------------------------------------------

Public Sub Test_ReadRange_Range_ReturnsCorrectArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim expect_arr(1 To 3, 1 To 3) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 3
        For idx2 = 1 To 3
            expect_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
        Next idx2
    Next idx1
    
    target_sheet.Range("C2:E4").Value = expect_arr
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=4, FinishColumn:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_arr() As Variant
    actual_arr = sheet_srv.ReadRange(range_bounds)
    
    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_ReadRange_Cell_ReturnsCorrectArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim expect_arr(1 To 1, 1 To 1) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 1
        For idx2 = 1 To 1
            expect_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
        Next idx2
    Next idx1
    
    target_sheet.Range("C2:C2").Value = expect_arr
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=2, FinishColumn:=3, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_arr() As Variant
    actual_arr = sheet_srv.ReadRange(range_bounds)
    
    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_ReadRange_OneRowRange_ReturnsCorrectArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim expect_arr(1 To 1, 1 To 3) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 1
        For idx2 = 1 To 3
            expect_arr(idx1, idx2) = "idx: " & idx1 & "-" & idx2
        Next idx2
    Next idx1
    
    target_sheet.Range("C2:E2").Value = expect_arr
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=2, FinishColumn:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_arr() As Variant
    actual_arr = sheet_srv.ReadRange(range_bounds)
    
    ' Assert
    Assert.EqualsArray expect_arr, actual_arr
End Sub

' -----------------------------------------------------------------------------
' IsEmptyCell
' -----------------------------------------------------------------------------

Public Sub Test_IsEmptyCell_EmptyCell_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).ClearContents

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsEmptyCell_EmptyStringValueWithIgnoreEmptyStringFalse_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    Call pWriteEmptyStringValue(target_sheet.Cells(3, 2))

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds, IgnoreEmptyString:=False)

    ' Assert
    Assert.IsFalse actual_value
End Sub

Public Sub Test_IsEmptyCell_EmptyStringValueWithIgnoreEmptyStringTrue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    Call pWriteEmptyStringValue(target_sheet.Cells(4, 2))

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=4, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds, IgnoreEmptyString:=True)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsEmptyCell_FormulaEmptyStringWithIgnoreEmptyStringFalse_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(5, 2).Formula = "="""""

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=5, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds, IgnoreEmptyString:=False)

    ' Assert
    Assert.IsFalse actual_value
End Sub

Public Sub Test_IsEmptyCell_FormulaEmptyStringWithIgnoreEmptyStringTrue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(6, 2).Formula = "="""""

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=6, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds, IgnoreEmptyString:=True)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_IsEmptyCell_HiddenValueWithIgnoreEmptyStringTrue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(7, 2).Value = 123
    target_sheet.Cells(7, 2).NumberFormatLocal = ";;;"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=7, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.IsEmptyCell(range_bounds, IgnoreEmptyString:=True)

    ' Assert
    Assert.IsFalse actual_value
End Sub

' -----------------------------------------------------------------------------
' HasFormula
' -----------------------------------------------------------------------------

Public Sub Test_HasFormula_FormulaCell_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Formula = "=1+2"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.HasFormula(range_bounds)

    ' Assert
    Assert.IsTrue actual_value
End Sub

Public Sub Test_HasFormula_ValueCell_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = "ABC"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.HasFormula(range_bounds)

    ' Assert
    Assert.IsFalse actual_value
End Sub
' -----------------------------------------------------------------------------
' GetUsedRangeBounds
' -----------------------------------------------------------------------------

Public Sub Test_GetUsedRangeBounds_NotEmpty_ReturnsCorrectBounds(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Value = "B1"
    target_sheet.Cells(3, 4).Value = "D3"
    target_sheet.Cells(5, 6).Value = "F5"
    target_sheet.Cells(7, 8).Value = "H7"
    Call target_sheet.Cells(1, 2).ClearContents
    Call target_sheet.Cells(7, 8).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=7, FinishColumn:=8, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds)
    
    ' Assert
    Assert.EqualsNumeric 1#, actual_bounds.Row
    Assert.EqualsNumeric 2#, actual_bounds.Column
    Assert.EqualsNumeric 5#, actual_bounds.FinishRow
    Assert.EqualsNumeric 6#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_IgnoresPreviousSearchFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Value = "start"
    target_sheet.Cells(6, 5).Value = "finish"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=8, FinishColumn:=8, Sheet:="test_output")

    Call pSetStaleFindFormat(target_sheet.Range("A1:H8"))

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    Call pClearStaleFindFormat(target_sheet.Range("A1:H8"))

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric 2, actual_bounds.Row
    Assert.EqualsNumeric 2, actual_bounds.Column
    Assert.EqualsNumeric 6, actual_bounds.FinishRow
    Assert.EqualsNumeric 5, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_Empty_ReturnsEmptyBounds(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Value = "B1"
    target_sheet.Cells(3, 4).Value = "D3"
    target_sheet.Cells(5, 6).Value = "F5"
    target_sheet.Cells(7, 8).Value = "H7"
    Call target_sheet.Cells(1, 2).ClearContents
    Call target_sheet.Cells(3, 4).ClearContents
    Call target_sheet.Cells(5, 6).ClearContents
    Call target_sheet.Cells(7, 8).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=7, FinishColumn:=8, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds)
    
    ' Assert
    Assert.EqualsNumeric 1#, actual_bounds.Row
    Assert.EqualsNumeric 2#, actual_bounds.Column
    Assert.EqualsNumeric 0#, actual_bounds.FinishRow
    Assert.EqualsNumeric 0#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_NotEmpty_ReturnsCorrectBounds2(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Value = "B1"
    target_sheet.Cells(3, 4).Value = "D3"
    target_sheet.Cells(5, 6).Value = "F5"
    target_sheet.Cells(7, 8).Value = "H7"
    Call target_sheet.Cells(1, 2).ClearContents
    Call target_sheet.Cells(7, 8).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=7, FinishColumn:=8, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 3#, actual_bounds.Row
    Assert.EqualsNumeric 4#, actual_bounds.Column
    Assert.EqualsNumeric 5#, actual_bounds.FinishRow
    Assert.EqualsNumeric 6#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_Empty_ReturnsEmptyBounds2(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Value = "B1"
    target_sheet.Cells(3, 4).Value = "D3"
    target_sheet.Cells(5, 6).Value = "F5"
    target_sheet.Cells(7, 8).Value = "H7"
    Call target_sheet.Cells(1, 2).ClearContents
    Call target_sheet.Cells(3, 4).ClearContents
    Call target_sheet.Cells(5, 6).ClearContents
    Call target_sheet.Cells(7, 8).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=7, FinishColumn:=8, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 1#, actual_bounds.Row
    Assert.EqualsNumeric 2#, actual_bounds.Column
    Assert.EqualsNumeric 0#, actual_bounds.FinishRow
    Assert.EqualsNumeric 0#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_NotEmpty_ReturnsCorrectBounds3(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(5, 2).Value = "B5"
    target_sheet.Cells(1, 4).Value = "D1"
    target_sheet.Cells(7, 6).Value = "F7"
    target_sheet.Cells(3, 8).Value = "H3"
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=8, FinishColumn:=9, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 3#, actual_bounds.Row
    Assert.EqualsNumeric 6#, actual_bounds.Column
    Assert.EqualsNumeric 7#, actual_bounds.FinishRow
    Assert.EqualsNumeric 8#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_Empty_ReturnsEmptyBounds3(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Value = "B1"
    target_sheet.Cells(3, 4).Value = "D3"
    target_sheet.Cells(5, 6).Value = "F5"
    target_sheet.Cells(7, 8).Value = "H7"
    Call target_sheet.Cells(3, 4).ClearContents
    Call target_sheet.Cells(5, 6).ClearContents
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=7, FinishColumn:=7, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 2#, actual_bounds.Row
    Assert.EqualsNumeric 2#, actual_bounds.Column
    Assert.EqualsNumeric 0#, actual_bounds.FinishRow
    Assert.EqualsNumeric 0#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_NotEmpty_ReturnsCorrectBounds4(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Formula = "="""""
    target_sheet.Cells(3, 4).Formula = "="""""
    target_sheet.Cells(5, 6).Formula = "="""""
    target_sheet.Cells(7, 8).Formula = "="""""
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=7, FinishColumn:=7, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IgnoreEmpty:=False, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 3#, actual_bounds.Row
    Assert.EqualsNumeric 4#, actual_bounds.Column
    Assert.EqualsNumeric 5#, actual_bounds.FinishRow
    Assert.EqualsNumeric 6#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_Empty_ReturnsEmptyBounds4(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    target_sheet.Cells(1, 2).Formula = "="""""
    target_sheet.Cells(3, 4).Formula = "="""""
    target_sheet.Cells(5, 6).Formula = "="""""
    target_sheet.Cells(7, 8).Formula = "="""""
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=7, FinishColumn:=7, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - Text モード
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)
    
    ' Assert
    Assert.EqualsNumeric 2#, actual_bounds.Row
    Assert.EqualsNumeric 2#, actual_bounds.Column
    Assert.EqualsNumeric 0#, actual_bounds.FinishRow
    Assert.EqualsNumeric 0#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_EmptySheet_ReturnsCorrectBounds(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, GetRawRange:=True)
    
    ' Assert
    Assert.EqualsNumeric 1#, actual_bounds.Row
    Assert.EqualsNumeric 1#, actual_bounds.Column
    Assert.EqualsNumeric 0#, actual_bounds.FinishRow
    Assert.EqualsNumeric 0#, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_OneColumnUsedRangeWithClearedTopCell_DoesNotTreatAsEmpty(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Cells(2, 2).Interior.Color = RGB(255, 255, 0)
    target_sheet.Cells(4, 2).Value = "bottom"

    Dim used_range_address As String
    used_range_address = target_sheet.UsedRange.Address

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)

    ' Assert
    Assert.EqualsNumeric 4, actual_bounds.Row
    Assert.EqualsNumeric 2, actual_bounds.Column
    Assert.EqualsNumeric 4, actual_bounds.FinishRow
    Assert.EqualsNumeric 2, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_EntireSheetRange_DoesNotOverflow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_full_sheet")
    target_sheet.Cells(1, 1).Value = "start"
    target_sheet.Cells(target_sheet.Rows.Count, target_sheet.Columns.Count).Value = "finish"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Sheet:="test_full_sheet")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Dim actual_bounds As WorksheetRangeBounds
    Err.Clear
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        target_sheet.Cells.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric 1, actual_bounds.Row
    Assert.EqualsNumeric 1, actual_bounds.Column
    Assert.EqualsNumeric target_sheet.Rows.Count, actual_bounds.FinishRow
    Assert.EqualsNumeric target_sheet.Columns.Count, actual_bounds.FinishColumn

    target_sheet.Cells.Clear
End Sub

Public Sub Test_GetUsedRangeBounds_IgnoreEmptyTrue_FormulaBlankCellNotCounted(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' 2箇所に値を入れる
    target_sheet.Cells(2, 3).Value = "XYZ"
    ' 1箇所に空文字を返す数式を入れる
    target_sheet.Cells(4, 4).Formula = "="""""  ' このセルは計算結果が "" (空文字)
    
    ' RangeBounds: 行1~5, 列2~5
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=5, FinishColumn:=5, Sheet:="test_output")
    
    ' Act
    ' IgnoreEmpty=True → 空文字セルは未使用と見なす
    Dim sheet_srv As WorksheetService
    Set sheet_srv = New WorksheetService
    
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False, IgnoreEmpty:=True)
    
    ' Assert
    ' (2,3) だけが有効セル => Row=2,Col=3
    ' => FinishRow=2, FinishColumn=3 となるはず
    Assert.EqualsNumeric 2, actual_bounds.Row
    Assert.EqualsNumeric 3, actual_bounds.Column
    Assert.EqualsNumeric 2, actual_bounds.FinishRow
    Assert.EqualsNumeric 3, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_IgnoreEmptyFalse_FormulaBlankCellIsCounted(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' 2箇所に値を入れる
    target_sheet.Cells(2, 3).Value = "ABC"
    ' 1箇所に空文字式
    target_sheet.Cells(4, 4).Formula = "="""""  ' → 計算結果は "" だがセルには数式あり
    
    ' RangeBounds: 行1~5, 列2~5
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=5, FinishColumn:=5, Sheet:="test_output")
    
    ' Act
    ' IgnoreEmpty=False → 数式があれば、空文字結果でも使用中と見なす
    Dim sheet_srv As WorksheetService
    Set sheet_srv = New WorksheetService
    
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, IsUsedStartRow:=False, IsUsedStartColumn:=False, IgnoreEmpty:=False)
    
    ' Assert
    ' 使われているのは (2,3) と (4,4)
    ' → Row最小=2,Row最大=4; Col最小=3,Col最大=4
    Assert.EqualsNumeric 2, actual_bounds.Row
    Assert.EqualsNumeric 3, actual_bounds.Column
    Assert.EqualsNumeric 4, actual_bounds.FinishRow
    Assert.EqualsNumeric 4, actual_bounds.FinishColumn
End Sub

Public Sub Test_GetUsedRangeBounds_GetRawRangeTrue_IncludingEmptyCells(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' シート全体には何も入っていないケースでも
    ' "GetRawRange=True" だとUsedRange全体を返す → 実装次第
    ' あるいはUsedRangeそのものがrow=1,col=1, finishRow=1, finishCol=1か、などExcelの挙動をベースにする
    ' ここでは既存のUsedRangeがある程度あると想定
    target_sheet.Range("B2").Value = "Something"
    target_sheet.Range("D5").Value = "Other"
    
    Dim range_bounds As WorksheetRangeBounds
    ' 大きめの範囲
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=10, FinishColumn:=10, Sheet:="test_output")
    
    Dim sheet_srv As WorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    ' GetRawRange=True → A1からUsedRangeの右下まで取得する実装を想定
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, GetRawRange:=True)
    
    ' Assert
    ' ExcelのUsedRangeが B2..D5 だったとすると
    ' Row=2, Column=2, FinishRow=5,FinishColumn=4 などが typical
    ' しかし "GetRawRange=True" の実装によっては A1..E6 とか
    ' ここは実装に合わせて調整。
    Assert.IsTrue actual_bounds.Row <= 2
    Assert.IsTrue actual_bounds.Column <= 2
    Assert.IsTrue actual_bounds.FinishRow >= 5
    Assert.IsTrue actual_bounds.FinishColumn >= 4
End Sub

Public Sub Test_GetUsedRangeBounds_IsUsedStartRowFalse_IsUsedStartColumnTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    ' For example:
    target_sheet.Cells(5, 3).Value = "DataX" ' row=5,col=3
    ' range= 1..10,2..10
    
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=2, FinishRow:=10, FinishColumn:=10, Sheet:="test_output")
    
    Dim sheet_srv As WorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    ' IsUsedStartRow=False => 使われていない開始行はスキップ
    ' IsUsedStartColumn=True => 開始列は 2 のまま
    Dim actual_bounds As WorksheetRangeBounds
    Set actual_bounds = sheet_srv.GetUsedRangeBounds(range_bounds, _
                                IgnoreEmpty:=True, _
                                IsUsedStartRow:=False, _
                                IsUsedStartColumn:=True)
    
    ' Assert
    ' Rowが 5 になる(expected) / Column=2 remains
    Assert.EqualsNumeric 5, actual_bounds.Row
    Assert.EqualsNumeric 2, actual_bounds.Column
    ' Finishes => row=5, col=3, or bigger if more used cells
    Assert.IsTrue actual_bounds.FinishRow >= 5
    Assert.IsTrue actual_bounds.FinishColumn >= 3
End Sub

' -----------------------------------------------------------------------------
' CopyCell - Additional Test Cases
' -----------------------------------------------------------------------------

Public Sub Test_CopyCell_WithValue_CopiesToDestination(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' コピー元に数値を設定
    src_sheet.Range("B2").Value = 12345
    
    ' Set RangeBounds for source and destination
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act - CopyCell
    sheet_srv.CopyCell src_bounds, dst_bounds
    
    ' Assert
    Assert.EqualsNumeric 12345, dst_sheet.Range("E2").Value
End Sub

Public Sub Test_CopyCell_WithFormula_CopiesFormula(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' コピー元に通常の数式を設定
    src_sheet.Range("C2").Formula = "=SUM(10,20)"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds
    
    ' Assert
    Assert.Equals "=SUM(10,20)", dst_sheet.Range("C2").Formula
    Assert.EqualsNumeric 30, dst_sheet.Range("C2").Value
End Sub

Public Sub Test_CopyCell_RelativeFormula_AdjustsReference(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim original_reference_style As XlReferenceStyle
    original_reference_style = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("A2").Value = "source"
    dst_sheet.Range("D2").Value = "expected"
    dst_sheet.Range("F2").NumberFormatLocal = "@"
    src_sheet.Range("C2").Formula = "=A2"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    Application.ReferenceStyle = original_reference_style

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue dst_sheet.Range("F2").HasFormula
    Assert.Equals "=D2", dst_sheet.Range("F2").Formula
    Assert.Equals "expected", dst_sheet.Range("F2").Value
End Sub

Public Sub Test_CopyCell_ArrayFormulaWriteError_KeepsSourceArrayFormula(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("D2:D4").FormulaArray = "=TRANSPOSE({1,2,3})"
    dst_sheet.Protect Contents:=True

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    dst_sheet.Unprotect

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue src_sheet.Range("D2:D4").HasArray
    Assert.Equals "=TRANSPOSE({1,2,3})", src_sheet.Range("D2").FormulaArray
End Sub

Public Sub Test_CopyCell_FormulaCopyError_RestoresDestination(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim original_enable_events As Boolean
    original_enable_events = Application.EnableEvents
    Application.EnableEvents = True

    ' Arrange
    Dim src_sheet As Worksheet
    Dim dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_formula_restore_input")
    Set dst_sheet = pPrepareTestSheet("test_formula_restore_output")

    src_sheet.Range("A1").Value = 10
    src_sheet.Range("C2").Formula = "=A1"
    src_sheet.Range("C2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("F2").Formula = "=1+1"
    dst_sheet.Range("F2").NumberFormatLocal = "0.0000"
    dst_sheet.Range("Z1").ClearContents

    Dim protect_probe As WorksheetChangeTestProbe
    Set protect_probe = New WorksheetChangeTestProbe
    Call protect_probe.Initialize(dst_sheet, "F2", "Z1")

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:=src_sheet.Name)

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:=dst_sheet.Name)

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds, CopyNumberFormat:=True

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    If dst_sheet.ProtectContents Then dst_sheet.Unprotect
    Set protect_probe = Nothing
    Application.EnableEvents = original_enable_events

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "=1+1", dst_sheet.Range("F2").Formula
    Assert.EqualsNumeric 2, dst_sheet.Range("F2").Value
    Assert.Equals "0.0000", dst_sheet.Range("F2").NumberFormatLocal
End Sub

Public Sub Test_CopyCell_WithArrayFormula_CopiesArrayFormula(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 配列数式を設定
    src_sheet.Range("D2:D4").FormulaArray = "=TRANSPOSE({1,2,3})"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds
    
    ' Assert
    Assert.IsTrue dst_sheet.Range("G2").HasArray
    Assert.Equals "=TRANSPOSE({1,2,3})", dst_sheet.Range("G2").FormulaArray
End Sub

Public Sub Test_CopyCell_AsValue_FromFormula_CopiesAsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 元セルに数式を入れておく
    src_sheet.Range("E2").Formula = "=1+2+3"  ' 結果は 6
    src_sheet.Range("E2").NumberFormatLocal = "#,##0.0"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds, AsValue:=True
    
    ' Assert
    Assert.EqualsNumeric 6, dst_sheet.Range("E2").Value
    Assert.Equals "6", dst_sheet.Range("E2").Formula
End Sub

Public Sub Test_CopyCell_WithErrorValue_CopiesError(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' #DIV/0! 作成
    src_sheet.Range("F2").Formula = "=1/0"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds
    
    ' Assert
    Assert.Equals CVErr(xlErrDiv0), dst_sheet.Range("F2").Value
End Sub

Public Sub Test_CopyCell_EmptySource_ClearsDestination(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 空セル
    src_sheet.Range("G2").ClearContents
    
    ' 先にコピー先に値を入れておく
    dst_sheet.Range("G2").Value = "Initial"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds
    
    ' Assert
    Assert.IsTrue IsEmpty(dst_sheet.Range("G2").Value)
End Sub

' -----------------------------------------------------------------------------
' CopyRange
' -----------------------------------------------------------------------------
Public Sub Test_CopyRange_OverlappedSameSheetDown_CopiesOriginalValues(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("A1").Value = "src1"
    target_sheet.Range("A2").Value = "src2"
    target_sheet.Range("A3").Value = "dst"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=1, Sheet:="test_output")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src1", target_sheet.Range("A2").Value
    Assert.Equals "src2", target_sheet.Range("A3").Value
End Sub

Public Sub Test_CopyRange_OverlappedSameSheetRight_CopiesOriginalValues(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("A1").Value = "src1"
    target_sheet.Range("B1").Value = "src2"
    target_sheet.Range("C1").Value = "dst"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=1, FinishColumn:=2, Sheet:="test_output")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=1, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src1", target_sheet.Range("B1").Value
    Assert.Equals "src2", target_sheet.Range("C1").Value
End Sub

Public Sub Test_CopyRange_OverlappedSameSheetDownRight_CopiesOriginalValues(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("A1").Value = "src11"
    target_sheet.Range("B1").Value = "src12"
    target_sheet.Range("A2").Value = "src21"
    target_sheet.Range("B2").Value = "src22"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=2, FinishColumn:=2, Sheet:="test_output")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src11", target_sheet.Range("B2").Value
    Assert.Equals "src12", target_sheet.Range("C2").Value
    Assert.Equals "src21", target_sheet.Range("B3").Value
    Assert.Equals "src22", target_sheet.Range("C3").Value
End Sub

Public Sub Test_CopyRange_OverlappedSameSheetUpLeft_CopiesOriginalValues(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("B2").Value = "src11"
    target_sheet.Range("C2").Value = "src12"
    target_sheet.Range("B3").Value = "src21"
    target_sheet.Range("C3").Value = "src22"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src11", target_sheet.Range("A1").Value
    Assert.Equals "src12", target_sheet.Range("B1").Value
    Assert.Equals "src21", target_sheet.Range("A2").Value
    Assert.Equals "src22", target_sheet.Range("B2").Value
End Sub

Public Sub Test_CopyRange_Call_CopyToDestination(ByVal Assert As UnitTestAssert)
    ' Arrange
    ' シートの準備
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 配列の準備
    Dim src_arr(1 To 5, 1 To 5) As String
    Dim dst_arr(1 To 5, 1 To 5) As String
    Dim expect_arr(1 To 5, 1 To 5) As Variant
    Dim idx1 As Long, idx2 As Long
    For idx1 = 1 To 5
        For idx2 = 1 To 5
            src_arr(idx1, idx2) = "src"
            dst_arr(idx1, idx2) = "dst"
            expect_arr(idx1, idx2) = "dst"
        Next idx2
    Next idx1
    
    ' シートに配列を書き込み
    src_sheet.Range("D3:H7").Value = src_arr
    dst_sheet.Range("E4:I8").Value = dst_arr
    
    ' RangeBounds を準備
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=4, Column:=5, FinishRow:=6, FinishColumn:=7, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=5, Column:=6, Sheet:="test_output")
    
    ' テスト対象のインスタンスを生成
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Call sheet_srv.CopyRange(src_bounds, dst_bounds)
    
    Dim actual_arr() As Variant
    actual_arr() = dst_sheet.Range("E4:I8").Value
    
    ' Assert
    expect_arr(2, 2) = "src"
    expect_arr(2, 3) = "src"
    expect_arr(2, 4) = "src"
    expect_arr(3, 2) = "src"
    expect_arr(3, 3) = "src"
    expect_arr(3, 4) = "src"
    expect_arr(4, 2) = "src"
    expect_arr(4, 3) = "src"
    expect_arr(4, 4) = "src"
    
    'For idx1 = 1 To 5
    '    Dim test_str As String
    '    test_str = expect_arr(idx1, 1)
    '    For idx2 = 1 To 5
    '        test_str = test_str & expect_arr(idx1, idx2)
    '    Next idx2
    '    Debug.Print test_str
    'Next idx1
    
    'Dim expect_item As Variant
    'For Each expect_item In expect_arr
    '    Debug.Print expect_item
    'Next expect_item
    
    Assert.EqualsArray expect_arr, actual_arr
End Sub

Public Sub Test_CopyRange_RelativeFormula_AdjustsReferences(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    Dim original_reference_style As XlReferenceStyle
    original_reference_style = Application.ReferenceStyle
    Application.ReferenceStyle = xlA1

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("A2:A3").Value = "source"
    dst_sheet.Range("D2:D3").Value = "expected"
    dst_sheet.Range("F2:F3").NumberFormatLocal = "@"
    src_sheet.Range("C2:C3").Formula = "=A2"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, FinishRow:=3, FinishColumn:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    Application.ReferenceStyle = original_reference_style

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue dst_sheet.Range("F2").HasFormula
    Assert.IsTrue dst_sheet.Range("F2:F3").HasFormula
    Assert.Equals "=D2", dst_sheet.Range("F2").Formula
    Assert.Equals "expected", dst_sheet.Range("F2").Value
    Assert.Equals "=D3", dst_sheet.Range("F3").Formula
    Assert.Equals "expected", dst_sheet.Range("F3").Value
End Sub

Public Sub Test_CopyRange_DestinationPastLastRow_RaisesBeforeCopy(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_edge_input_row")
    Set dst_sheet = pPrepareTestSheet("test_edge_output_row")

    src_sheet.Range("B2").Value = "src1"
    src_sheet.Range("B3").Value = "src2"
    src_sheet.Range("B4").Value = "src3"
    dst_sheet.Cells(G_ROW_MAX - 1, 2).Value = "keep1"
    dst_sheet.Cells(G_ROW_MAX, 2).Value = "keep2"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_edge_input_row")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=G_ROW_MAX - 1, Column:=2, Sheet:="test_edge_output_row")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "keep1", dst_sheet.Cells(G_ROW_MAX - 1, 2).Value
    Assert.Equals "keep2", dst_sheet.Cells(G_ROW_MAX, 2).Value
End Sub

Public Sub Test_CopyRange_DestinationPastLastColumn_RaisesBeforeCopy(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_edge_input_col")
    Set dst_sheet = pPrepareTestSheet("test_edge_output_col")

    src_sheet.Range("B2").Value = "src1"
    src_sheet.Range("C2").Value = "src2"
    src_sheet.Range("D2").Value = "src3"
    dst_sheet.Cells(2, G_COL_MAX - 1).Value = "keep1"
    dst_sheet.Cells(2, G_COL_MAX).Value = "keep2"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=4, Sheet:="test_edge_input_col")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=G_COL_MAX - 1, Sheet:="test_edge_output_col")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "keep1", dst_sheet.Cells(2, G_COL_MAX - 1).Value
    Assert.Equals "keep2", dst_sheet.Cells(2, G_COL_MAX).Value
End Sub

Public Sub Test_CopyRange_AsValue_CopiesValuesOnly(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 複数セルに数式を入れる
    src_sheet.Range("B2").Formula = "=1+2"
    src_sheet.Range("B3").Formula = "=2+3"
    src_sheet.Range("B4").Formula = "=3+4"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds, AsValue:=True
    
    ' Assert
    Assert.EqualsNumeric 3, dst_sheet.Range("E2").Value
    Assert.EqualsNumeric 5, dst_sheet.Range("E3").Value
    Assert.EqualsNumeric 7, dst_sheet.Range("E4").Value
    Assert.Equals "3", dst_sheet.Range("E2").Formula
    Assert.Equals "5", dst_sheet.Range("E3").Formula
    Assert.Equals "7", dst_sheet.Range("E4").Formula
End Sub

Public Sub Test_CopyRange_CopyNumberFormat_CopiesFormat(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 数値と書式を設定
    src_sheet.Range("C2").Value = 1234.56
    src_sheet.Range("C2").NumberFormatLocal = "#,##0.00"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=4, Column:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds, CopyNumberFormat:=True
    
    ' Assert
    Assert.EqualsNumeric 1234.56, dst_sheet.Range("E4").Value
    Assert.Equals "#,##0.00", dst_sheet.Range("E4").NumberFormatLocal
End Sub

Public Sub Test_CopyRange_WithArrayFormula_CopiesArrayFormula(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 2行×3列の配列数式を設定
    src_sheet.Range("D2:F3").FormulaArray = "=TRANSPOSE({10,20})"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, FinishRow:=3, FinishColumn:=5, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=4, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds
    
    ' Assert
    Assert.IsTrue dst_sheet.Range("D2:E3").HasArray
    Assert.Equals "=TRANSPOSE({10,20})", dst_sheet.Range("D2").FormulaArray
    Assert.EqualsNumeric 10, dst_sheet.Range("D2").Value
    Assert.EqualsNumeric 20, dst_sheet.Range("D3").Value
End Sub

Public Sub Test_CopyRange_ArrayFormulaFullHeight_DoesNotDropLastRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("D2:D4").FormulaArray = "=TRANSPOSE({10,20,30})"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, FinishRow:=4, FinishColumn:=4, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue dst_sheet.Range("G2:G4").HasArray
    Assert.Equals "=TRANSPOSE({10,20,30})", dst_sheet.Range("G2").FormulaArray
    Assert.EqualsNumeric 10, dst_sheet.Range("G2").Value
    Assert.EqualsNumeric 20, dst_sheet.Range("G3").Value
    Assert.EqualsNumeric 30, dst_sheet.Range("G4").Value
End Sub

Public Sub Test_CopyRange_ArrayFormulaFullWidth_DoesNotDropLastColumn(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("D2:F2").FormulaArray = "={10,20,30}"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, FinishRow:=2, FinishColumn:=6, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue dst_sheet.Range("G2:I2").HasArray
    Assert.Equals "={10,20,30}", dst_sheet.Range("G2").FormulaArray
    Assert.EqualsNumeric 10, dst_sheet.Range("G2").Value
    Assert.EqualsNumeric 20, dst_sheet.Range("H2").Value
    Assert.EqualsNumeric 30, dst_sheet.Range("I2").Value
End Sub

Public Sub Test_CopyRange_MultipleCells_EmptySource_ClearsDestination(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' Source の一部セルに値、他セルは空
    src_sheet.Range("B2").Value = "A"
    src_sheet.Range("B4").Value = "B"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=5, FinishColumn:=2, Sheet:="test_input")
    
    ' Copy 先に初期値を入れておく
    dst_sheet.Range("C2:C5").Value = "X"
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds
    
    ' Assert
    Assert.Equals "A", dst_sheet.Range("C2").Value
    Assert.IsTrue IsEmpty(dst_sheet.Range("C3").Value)
    Assert.Equals "B", dst_sheet.Range("C4").Value
    Assert.IsTrue IsEmpty(dst_sheet.Range("C5").Value)
End Sub

Public Sub Test_CopyRange_WithErrorCells_CopiesErrors(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")
    
    ' 複数セルにエラー式
    src_sheet.Range("E2").Formula = "=1/0"
    src_sheet.Range("E3").Formula = "=NA()"
    
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=5, FinishRow:=3, FinishColumn:=5, Sheet:="test_input")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=4, Column:=5, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds
    
    ' Assert
    Assert.Equals CVErr(xlErrDiv0), dst_sheet.Range("E4").Value
    Assert.Equals CVErr(xlErrNA), dst_sheet.Range("E5").Value
End Sub

Public Sub Test_CopyRange_SameShapeDestination_CopiesValues(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("B2:C3").Value = "src"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=4, Column:=5, FinishRow:=5, FinishColumn:=6, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src", dst_sheet.Range("E4").Value
    Assert.Equals "src", dst_sheet.Range("F5").Value
End Sub

Public Sub Test_CopyRange_EmptyDestination_UsesAnchor(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("B2").Value = "src"
    dst_sheet.Range("C5").Value = "keep"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=5, Column:=3, FinishRow:=0, FinishColumn:=0, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    ' Assert
    Assert.Equals "src", dst_sheet.Range("C5").Value
End Sub

Public Sub Test_CopyRange_MismatchedDestination_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("B2:C3").Value = "src"
    dst_sheet.Range("E4:F6").Value = "keep"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=4, Column:=5, FinishRow:=6, FinishColumn:=6, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", dst_sheet.Range("E4").Value
    Assert.Equals "keep", dst_sheet.Range("F6").Value
End Sub

Public Sub Test_CopyCell_CopyNumberFormatFalse_Value_KeepsDestinationFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("B2").Value = 1234.56
    src_sheet.Range("B2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("E2").NumberFormatLocal = "0.0000"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds, CopyNumberFormat:=False

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric 1234.56, dst_sheet.Range("E2").Value
    Assert.Equals "0.0000", dst_sheet.Range("E2").NumberFormatLocal
End Sub

Public Sub Test_CopyCell_CopyNumberFormatFalse_Formula_KeepsDestinationFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("C2").Formula = "=10/4"
    src_sheet.Range("C2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("F2").NumberFormatLocal = "0.0000"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=6, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds, CopyNumberFormat:=False

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "=10/4", dst_sheet.Range("F2").Formula
    Assert.Equals "0.0000", dst_sheet.Range("F2").NumberFormatLocal
End Sub

Public Sub Test_CopyCell_AsValueCopyNumberFormatFalse_KeepsDestinationFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("D2").Formula = "=10/4"
    src_sheet.Range("D2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("G2").NumberFormatLocal = "0.0000"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=4, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=7, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds, AsValue:=True, CopyNumberFormat:=False

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric 2.5, dst_sheet.Range("G2").Value
    Assert.Equals "0.0000", dst_sheet.Range("G2").NumberFormatLocal
End Sub

Public Sub Test_CopyCell_CopyNumberFormatFalse_EmptySource_KeepsDestinationFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("H2").ClearContents
    src_sheet.Range("H2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("H2").Value = "delete me"
    dst_sheet.Range("H2").NumberFormatLocal = "0.0000"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=8, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=8, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyCell src_bounds, dst_bounds, CopyNumberFormat:=False

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue IsEmpty(dst_sheet.Range("H2").Value)
    Assert.Equals "0.0000", dst_sheet.Range("H2").NumberFormatLocal
End Sub

Public Sub Test_CopyRange_CopyNumberFormatFalse_KeepsDestinationFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_sheet As Worksheet, dst_sheet As Worksheet
    Set src_sheet = pPrepareTestSheet("test_input")
    Set dst_sheet = pPrepareTestSheet("test_output")

    src_sheet.Range("B2").Value = 1234.56
    src_sheet.Range("C2").Value = 7890.12
    src_sheet.Range("B2:C2").NumberFormatLocal = "#,##0.00"
    dst_sheet.Range("E2:F2").NumberFormatLocal = "0.0000"

    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=2, FinishColumn:=3, Sheet:="test_input")

    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(Row:=2, Column:=5, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.CopyRange src_bounds, dst_bounds, CopyNumberFormat:=False

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric 1234.56, dst_sheet.Range("E2").Value
    Assert.EqualsNumeric 7890.12, dst_sheet.Range("F2").Value
    Assert.Equals "0.0000", dst_sheet.Range("E2").NumberFormatLocal
    Assert.Equals "0.0000", dst_sheet.Range("F2").NumberFormatLocal
End Sub


' -----------------------------------------------------------------------------
' RemoveDuplicates / Rows / Outline
' -----------------------------------------------------------------------------

Public Sub Test_RemoveDuplicates_ScalarColumn_RemovesDuplicateRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2").Value = "A"
    target_sheet.Range("C2").Value = "keep-first"
    target_sheet.Range("B3").Value = "A"
    target_sheet.Range("C3").Value = "remove"
    target_sheet.Range("B4").Value = "B"
    target_sheet.Range("C4").Value = "keep-second"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "A", target_sheet.Range("B2").Value
    Assert.Equals "keep-first", target_sheet.Range("C2").Value
    Assert.Equals "B", target_sheet.Range("B3").Value
    Assert.Equals "keep-second", target_sheet.Range("C3").Value
End Sub

Public Sub Test_RemoveDuplicates_ArrayColumns_RemovesDuplicateRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2").Value = "A"
    target_sheet.Range("C2").Value = "X"
    target_sheet.Range("D2").Value = "keep-first"
    target_sheet.Range("B3").Value = "A"
    target_sheet.Range("C3").Value = "X"
    target_sheet.Range("D3").Value = "remove"
    target_sheet.Range("B4").Value = "A"
    target_sheet.Range("C4").Value = "Y"
    target_sheet.Range("D4").Value = "keep-second"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=4, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, Array(1, 2))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "keep-first", target_sheet.Range("D2").Value
    Assert.Equals "keep-second", target_sheet.Range("D3").Value
End Sub

Public Sub Test_RemoveDuplicates_ColumnZero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, 0)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

Public Sub Test_RemoveDuplicates_ColumnPastRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, 3)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

Public Sub Test_RemoveDuplicates_EmptyColumnsArray_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2:C3").Value = "keep"

    Dim duplicate_columns() As Long

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, duplicate_columns)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("B2").Value
End Sub

Public Sub Test_InsertRows_EntireRowRange_InsertsRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A2").Value = "before"
    target_sheet.Range("A3").Value = "after"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, FinishRow:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.InsertRows(range_bounds)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "before", target_sheet.Range("A2").Value
    Assert.IsTrue IsEmpty(target_sheet.Range("A3").Value)
    Assert.Equals "after", target_sheet.Range("A4").Value
End Sub

Public Sub Test_InsertRows_CellRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, FinishRow:=4, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.InsertRows(range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("A3").Value
End Sub

Public Sub Test_InsertRows_EmptyRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=1, FinishRow:=0, FinishColumn:=0, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.InsertRows(range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
End Sub

Public Sub Test_DeleteRows_EntireRowRange_DeletesRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A2").Value = "before"
    target_sheet.Range("A3").Value = "delete"
    target_sheet.Range("A4").Value = "after"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, FinishRow:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.DeleteRows(range_bounds)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "before", target_sheet.Range("A2").Value
    Assert.Equals "after", target_sheet.Range("A3").Value
End Sub

Public Sub Test_DeleteRows_CellRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A3").Value = "keep"

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=3, Column:=2, FinishRow:=4, FinishColumn:=3, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.DeleteRows(range_bounds)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.Equals "Class WorksheetService", actual_error_source
    Assert.Equals "keep", target_sheet.Range("A3").Value
End Sub

Public Sub Test_SetSheetOutlineLevel_Default_DoesNotRaise(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.SetSheetOutlineLevel(range_bounds)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SetSheetOutlineLevel_RowOnly_DoesNotRaise(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Rows("2:3").Group
    Err.Clear

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.SetSheetOutlineLevel(range_bounds, RowLevels:=1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SetSheetOutlineLevel_InvalidRowLevel_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.SetSheetOutlineLevel(range_bounds, RowLevels:=9)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
End Sub

Public Sub Test_SetSheetOutlineLevel_InvalidColumnLevel_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Call sheet_srv.SetSheetOutlineLevel(range_bounds, ColumnLevels:=-1)

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description

    ' Assert
    If Not Assert.ErrorRaised(0, actual_error_number, actual_error_source, actual_error_description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
End Sub

' -----------------------------------------------------------------------------
' ClearRange / Color / Alignment
' -----------------------------------------------------------------------------

Public Sub Test_ClearRange_ClearColorsOnly_DoesNotClearNumberFormat(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("B2").Value = DateSerial(2026, 5, 9)
    target_sheet.Range("B2").NumberFormatLocal = "yyyy/mm/dd"
    target_sheet.Range("B2").Font.Color = RGB(0, 0, 255)
    target_sheet.Range("B2").Interior.Color = RGB(255, 0, 0)

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.ClearRange range_bounds, ClearContents:=False, ClearNumberFormats:=False, ClearColors:=True

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric CDbl(DateSerial(2026, 5, 9)), target_sheet.Range("B2").Value2
    Assert.Equals "yyyy/mm/dd", target_sheet.Range("B2").NumberFormatLocal
End Sub

Public Sub Test_ClearRange_HiddenRowAndColumnOutsideRange_DoesNotClearOutsideRange(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_clear_range")
    target_sheet.Rows.Hidden = False
    target_sheet.Columns.Hidden = False

    target_sheet.Range("B2:C3").Value = "target"
    target_sheet.Range("B8:C8").Value = "outside row"
    target_sheet.Range("F2:F3").Value = "outside col"
    target_sheet.Rows(8).Hidden = True
    target_sheet.Columns(6).Hidden = True

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=3, FinishColumn:=3, Sheet:="test_clear_range")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.ClearRange range_bounds

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        target_sheet.Rows(8).Hidden = False
        target_sheet.Columns(6).Hidden = False
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.IsTrue IsEmpty(target_sheet.Range("B2").Value)
    Assert.IsTrue IsEmpty(target_sheet.Range("C3").Value)
    Assert.Equals "outside row", target_sheet.Range("B8").Value
    Assert.Equals "outside row", target_sheet.Range("C8").Value
    Assert.Equals "outside col", target_sheet.Range("F2").Value
    Assert.Equals "outside col", target_sheet.Range("F3").Value

    target_sheet.Rows(8).Hidden = False
    target_sheet.Columns(6).Hidden = False
End Sub

Public Sub Test_SetSheetTabColor_BlackColorValue_TakesPrecedenceOverColorIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.SetSheetTabColor range_bounds, TabColorIndex:=3, TabColor:=vbBlack

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric CLng(vbBlack), CLng(target_sheet.Tab.Color)
End Sub

Public Sub Test_SetRangeColor_BlackColorValue_TakesPrecedenceOverColorIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.SetRangeColor range_bounds, FontColorIndex:=3, FontColor:=vbBlack, InteriorColorIndex:=3, InteriorColor:=vbBlack

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric CLng(vbBlack), CLng(target_sheet.Range("B2").Font.Color)
    Assert.EqualsNumeric CLng(vbBlack), CLng(target_sheet.Range("B2").Interior.Color)
End Sub

Public Sub Test_SetAlignment_RightBottom_SetsHorizontalAndVerticalAlignment(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    target_sheet.Range("B2").HorizontalAlignment = xlGeneral
    target_sheet.Range("B2").VerticalAlignment = xlTop

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    sheet_srv.SetAlignment range_bounds, HorizontalAlignment:=xlRight, VerticalAlignment:=xlBottom

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Assert.EqualsNumeric CLng(xlRight), CLng(target_sheet.Range("B2").HorizontalAlignment)
    Assert.EqualsNumeric CLng(xlBottom), CLng(target_sheet.Range("B2").VerticalAlignment)
End Sub
' -----------------------------------------------------------------------------
' EvaluateFormula / XLookup
' -----------------------------------------------------------------------------

Public Sub Test_EvaluateFormula_SimpleFormula_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    
    Dim context_bounds As WorksheetRangeBounds
    Set context_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.EvaluateFormula(context_bounds, "=SUM(1,2,3)")
    
    ' Assert
    Assert.EqualsNumeric 6, actual_value
End Sub

Public Sub Test_EvaluateFormula_SheetContext_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("B2").Value = 12
    
    Dim context_bounds As WorksheetRangeBounds
    Set context_bounds = New_RangeBounds(Row:=1, Column:=1, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.EvaluateFormula(context_bounds, "=B2*2")
    
    ' Assert
    Assert.EqualsNumeric 24, actual_value
End Sub

Public Sub Test_XLookup_Found_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A2:A4").Value = WorksheetFunction.Transpose(Array("A", "B", "C"))
    target_sheet.Range("B2:B4").Value = WorksheetFunction.Transpose(Array("Alpha", "Beta", "Gamma"))
    
    Dim lookup_bounds As WorksheetRangeBounds
    Set lookup_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=1, Sheet:="test_output")
    
    Dim return_bounds As WorksheetRangeBounds
    Set return_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.XLookup("B", lookup_bounds, return_bounds, "missing")
    
    ' Assert
    Assert.Equals "Beta", CStr(actual_value)
End Sub

Public Sub Test_XLookup_NotFound_ReturnsIfNotFound(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A2:A4").Value = WorksheetFunction.Transpose(Array("A", "B", "C"))
    target_sheet.Range("B2:B4").Value = WorksheetFunction.Transpose(Array("Alpha", "Beta", "Gamma"))
    
    Dim lookup_bounds As WorksheetRangeBounds
    Set lookup_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=1, Sheet:="test_output")
    
    Dim return_bounds As WorksheetRangeBounds
    Set return_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_output")
    
    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.XLookup("Z", lookup_bounds, return_bounds, "missing")
    
    ' Assert
    Assert.Equals "missing", CStr(actual_value)
End Sub

Public Sub Test_XLookup_ErrorIfNotFound_ReturnsErrorValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")
    target_sheet.Range("A2:A4").Value = WorksheetFunction.Transpose(Array("A", "B", "C"))
    target_sheet.Range("B2:B4").Value = WorksheetFunction.Transpose(Array("Alpha", "Beta", "Gamma"))

    Dim lookup_bounds As WorksheetRangeBounds
    Set lookup_bounds = New_RangeBounds(Row:=2, Column:=1, FinishRow:=4, FinishColumn:=1, Sheet:="test_output")

    Dim return_bounds As WorksheetRangeBounds
    Set return_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=4, FinishColumn:=2, Sheet:="test_output")

    Dim sheet_srv As IWorksheetService
    Set sheet_srv = New WorksheetService

    ' Act
    Err.Clear
    Dim actual_value As Variant
    actual_value = sheet_srv.XLookup("Z", lookup_bounds, return_bounds, CVErr(xlErrNA))

    Dim actual_error_number As Long
    Dim actual_error_source As String
    Dim actual_error_description As String
    actual_error_number = Err.Number
    actual_error_source = Err.Source
    actual_error_description = Err.Description
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, actual_error_number, actual_error_source, actual_error_description) Then Exit Sub
    Assert.Equals CVErr(xlErrNA), actual_value
End Sub
