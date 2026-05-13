Attribute VB_Name = "Test_Lib_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Lib_Common モジュールのユニット テストです。
'!
' #############################################################################

Public Sub Test_InitializeCommonService_NotInitialized_InitializesWorkbookAndWorksheetServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing
    Set WsSrv = Nothing

    ' Act
    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "WorkbookService", WbSrv
    Assert.IsTypeOf "WorksheetService", WsSrv
End Sub
' -----------------------------------------------------------------------------
' GetExcelFileFormat
' -----------------------------------------------------------------------------

Public Sub Test_GetExcelFileFormat_UppercaseXlsm_ReturnsMacroEnabledWorkbook(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.XLSM")

    ' Assert
    Assert.EqualsNumeric xlOpenXMLWorkbookMacroEnabled, actual_value
End Sub

Public Sub Test_GetExcelFileFormat_MixedCaseCsv_ReturnsCsv(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.Csv")

    ' Assert
    Assert.EqualsNumeric xlCSV, actual_value
End Sub

' -----------------------------------------------------------------------------
' JoinStringList / JoinStringSet
' -----------------------------------------------------------------------------

Public Sub Test_JoinStringList_PassedStringList_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_lst As ObjectList
    Set src_lst = New ObjectList
    Call src_lst.Add("alpha")
    Call src_lst.Add("beta")
    Call src_lst.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringList(src_lst, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

Public Sub Test_JoinStringSet_PassedStringSet_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_set As ObjectSet
    Set src_set = New ObjectSet
    Call src_set.Add("alpha")
    Call src_set.Add("beta")
    Call src_set.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringSet(src_set, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

' -----------------------------------------------------------------------------
' Excel address split / range address shape
' -----------------------------------------------------------------------------

Public Sub Test_SplitExcelAddress_QuotedWorkbookSheet_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'[Book.xlsm]Data Sheet'!$A$1:$B$2")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "Book.xlsm", book_name
    Assert.Equals "Data Sheet", sheet_name
    Assert.Equals "$A$1:$B$2", cell_address
End Sub

Public Sub Test_SplitExcelAddress_InvalidAddress_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "Sheet1!!A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SplitA1RangeAddress_Cell_ReturnsIndexes(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$B$2")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, start_row
    Assert.EqualsNumeric 2, start_col
    Assert.EqualsNumeric 2, finish_row
    Assert.EqualsNumeric 2, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_RowRange_ReturnsOmittedColumns(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "1:3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, start_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_col
    Assert.EqualsNumeric 3, finish_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_ColumnRange_ReturnsOmittedRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A:C")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_row
    Assert.EqualsNumeric 1, start_col
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_row
    Assert.EqualsNumeric 3, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_MultiRange_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A1,B2")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_RangeAddressShapeFunctions_ExpectedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsCell("A1")
    Assert.IsFalse IsArea("A1")
    Assert.IsTrue IsOneRow("A1")
    Assert.IsTrue IsOneColumn("A1")
    Assert.IsFalse IsOneRowArea("A1")
    Assert.IsFalse IsOneColumnArea("A1")

    Assert.IsTrue IsArea("A1:B1")
    Assert.IsTrue IsOneRow("A1:B1")
    Assert.IsTrue IsOneRowArea("A1:B1")
    Assert.IsFalse IsOneColumnArea("A1:B1")

    Assert.IsTrue IsArea("A1:A2")
    Assert.IsTrue IsOneColumn("A1:A2")
    Assert.IsFalse IsOneRowArea("A1:A2")
    Assert.IsTrue IsOneColumnArea("A1:A2")

    Assert.IsTrue IsArea("A1:B2")
    Assert.IsFalse IsOneRowArea("A1:B2")
    Assert.IsFalse IsOneColumnArea("A1:B2")

    Assert.IsTrue IsEntireRow("1:1")
    Assert.IsTrue IsOneRowArea("1:1")
    Assert.IsTrue IsEntireColumn("A:A")
    Assert.IsTrue IsOneColumnArea("A:A")
    Assert.IsTrue IsMultiRange("A1,B2")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_RangeAddressShapeFunctions_MultiRange_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsMultiRange("A1,B2")
    Assert.IsFalse IsArea("A1,B2")
    Assert.IsFalse IsCell("A1,B2")
    Assert.IsFalse IsEntireRow("A1,B2")
    Assert.IsFalse IsEntireColumn("A1,B2")
    Assert.IsFalse IsOneRow("A1,B2")
    Assert.IsFalse IsOneColumn("A1,B2")
    Assert.IsFalse IsOneRowArea("A1,B2")
    Assert.IsFalse IsOneColumnArea("A1,B2")

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub
