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

Public Sub Test_SplitExcelAddress_QuotedSheetWithExclamation_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'入力!確認'!A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "", book_name
    Assert.Equals "入力!確認", sheet_name
    Assert.Equals "A1", cell_address
End Sub

Public Sub Test_SplitExcelAddress_QuotedSheetWithEscapedApostrophe_ReturnsParts(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String

    ' Act
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, "'O''Brien'!$A$1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", folder_path
    Assert.Equals "", book_name
    Assert.Equals "O'Brien", sheet_name
    Assert.Equals "$A$1", cell_address
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

Public Sub Test_SplitA1RangeAddress_AbsoluteRowRange_ReturnsOmittedColumns(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$1:$3")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, start_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_col
    Assert.EqualsNumeric 3, finish_row
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_AbsoluteColumnRange_ReturnsOmittedRows(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$A:$C")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, start_row
    Assert.EqualsNumeric 1, start_col
    Assert.EqualsNumeric G_OMIT_CELL_INDEX, finish_row
    Assert.EqualsNumeric 3, finish_col
End Sub

Public Sub Test_SplitA1RangeAddress_DollarInsideColumn_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A$B1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
End Sub

Public Sub Test_SplitA1RangeAddress_DoubleColumnDollar_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "$$A$1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
End Sub

Public Sub Test_SplitA1RangeAddress_TrailingDollar_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim start_row As Long
    Dim start_col As Long
    Dim finish_row As Long
    Dim finish_col As Long

    ' Act
    Call SplitA1RangeAddress(start_row, start_col, finish_row, finish_col, "A1$")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Sub SplitA1RangeAddress", Err.Source
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

Public Sub Test_RangeAddress_ColumnRangeMixedAbsoluteFlags_UsesFinishColumnFlag(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartColumn:=2, _
            FinishColumn:=4, _
            IsAbsoluteStartColumn:=True, _
            IsAbsoluteFinishColumn:=False, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "$B:E", actual_value
End Sub

Public Sub Test_RangeAddress_ColumnRangeMixedAbsoluteFlags_CanMakeOnlyFinishColumnAbsolute(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartColumn:=2, _
            FinishColumn:=4, _
            IsAbsoluteStartColumn:=False, _
            IsAbsoluteFinishColumn:=True, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:$D", actual_value
End Sub

Public Sub Test_ExcelA1ColumnAddress_MaxColumn_ReturnsXFD(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(G_COL_MAX)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "XFD", actual_value
End Sub

Public Sub Test_ExcelA1ColumnAddress_Zero_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(0)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelA1ColumnAddress", Err.Source
End Sub

Public Sub Test_ExcelA1ColumnAddress_OverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = ExcelA1ColumnAddress(G_COL_MAX + 1)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function ExcelA1ColumnAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1MaxColumn_ReturnsXFD(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "XFD1", actual_value
End Sub

Public Sub Test_RangeAddress_A1ColumnOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX + 1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1RelativeColumnOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=1, _
            StartColumn:=G_COL_MAX, _
            ReferenceColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1MaxRow_ReturnsMaxRow(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX, _
            StartColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "A1048576", actual_value
End Sub

Public Sub Test_RangeAddress_A1RowOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX + 1, _
            StartColumn:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
End Sub

Public Sub Test_RangeAddress_A1RelativeRowOverMax_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = RangeAddress( _
            StartRow:=G_ROW_MAX, _
            StartColumn:=1, _
            ReferenceRow:=1, _
            AddressType:="A1")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function RangeAddress", Err.Source
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

' -----------------------------------------------------------------------------
' Path helpers
' -----------------------------------------------------------------------------

Public Sub Test_GetAbsolutePathFromParent_RelativeParentSegment_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "..\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_RootRelative_UsesParentDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_AbsolutePath_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:\Base\Child", "C:\Other\..\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\File.txt", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_DriveRelativePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetAbsolutePathFromParent("C:\Base\Child", "C:File.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetAbsolutePathFromParent", Err.Source
End Sub

Public Sub Test_GetAbsolutePathFromParent_NonAbsoluteParentPath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetAbsolutePathFromParent("Base\Child", "File.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetAbsolutePathFromParent", Err.Source
End Sub


Public Sub Test_IsUrlPath_SchemeUrl_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsUrlPath("https://example.com/sites/book.xlsm")
    Assert.IsTrue IsUrlPath("custom+scheme://host/path")
    Assert.IsFalse IsUrlPath("C:\Base\Child")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlParentRelativePath_NormalizesPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "..\book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team/book.xlsm", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlRootRelativePath_UsesUrlRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "\shared\book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/shared/book.xlsm", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_UrlQueryFragment_PreservesSuffix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("https://example.com/sites/team/docs", "sub\..\book.xlsm?q=C:\Temp\A/B#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team/docs/book.xlsm?q=C:\Temp\A/B#frag/c", actual_value
End Sub

Public Sub Test_GetAbsolutePathFromParent_WindowsForwardSlash_ReturnsBackslash(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetAbsolutePathFromParent("C:/Base/Child", "../File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\File.txt", actual_value
End Sub

Public Sub Test_GetPathRoot_DrivePath_ReturnsDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("C:\Base\Child")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\", actual_value
End Sub

Public Sub Test_GetPathRoot_LowercaseDrivePath_ReturnsDriveRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("z:\Base\Child")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Z:\", actual_value
End Sub

Public Sub Test_GetPathRoot_UncPath_ReturnsShareRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("\\server\share\folder\file.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "\\server\share\", actual_value
End Sub

Public Sub Test_GetPathRoot_RootRelativePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetPathRoot("\folder\file.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetPathRoot", Err.Source
End Sub

Public Sub Test_GetPathRoot_NonAlphabetDrivePath_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Call GetPathRoot("1:\folder\file.txt")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetPathRoot", Err.Source
End Sub

Public Sub Test_IsAbsolutePath_DriveRelativePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsAbsolutePath("C:folder\file.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetPathRoot_Url_ReturnsAuthorityRoot(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetPathRoot("https://example.com/sites/team/book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/", actual_value
End Sub

Public Sub Test_IsDrivePath_DriveSpecifier_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDrivePath("C:\Base\File.txt")
    Assert.IsTrue IsDrivePath("C:/Base/File.txt")
    Assert.IsTrue IsDrivePath("C:Base\File.txt")
    Assert.IsTrue IsDrivePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDrivePath_NonDriveSpecifier_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDrivePath("\\Server\Share\File.txt")
    Assert.IsFalse IsDrivePath("Base\File.txt")
    Assert.IsFalse IsDrivePath("1:\Base\File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveAbsolutePath_DriveAbsolutePath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDriveAbsolutePath("C:\Base\File.txt")
    Assert.IsTrue IsDriveAbsolutePath("C:/Base/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveAbsolutePath_DriveRelativePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDriveAbsolutePath("C:Base\File.txt")
    Assert.IsFalse IsDriveAbsolutePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveRelativePath_DriveRelativePath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsDriveRelativePath("C:Base\File.txt")
    Assert.IsTrue IsDriveRelativePath("C:")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsDriveRelativePath_DriveAbsolutePath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsDriveRelativePath("C:\Base\File.txt")
    Assert.IsFalse IsDriveRelativePath("C:/Base/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsUncPath_ValidUncPath_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsTrue IsUncPath("\\Server\Share\File.txt")
    Assert.IsTrue IsUncPath("//Server/Share/File.txt")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_IsUncPath_IncompleteUncPath_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.IsFalse IsUncPath("\\Server")
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_SplitPath_WindowsPath_ReturnsParentAndLeaf(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\Base\Child\File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_WindowsForwardSlash_ReturnsBackslashParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:/Base/Child/File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_NoSeparator_ReturnsEmptyParent(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "File.txt")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "", parent_path
    Assert.Equals "File.txt", leaf_path
End Sub

Public Sub Test_SplitPath_EndSeparator_RespectsIgnoreEndSep(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "C:\Base\Child\")

    Dim ignored_parent_path As String
    Dim ignored_leaf_path As String
    Call SplitPath(ignored_parent_path, ignored_leaf_path, "C:\Base\Child\", IgnoreEndSep:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Base\Child", parent_path
    Assert.Equals "", leaf_path
    Assert.Equals "C:\Base", ignored_parent_path
    Assert.Equals "Child", ignored_leaf_path
End Sub

Public Sub Test_SplitPath_Url_KeepsPathSuffixInLeaf(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim parent_path As String
    Dim leaf_path As String
    Call SplitPath(parent_path, leaf_path, "https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team", parent_path
    Assert.Equals "book.xlsm?q=a/b#frag/c", leaf_path
End Sub

Public Sub Test_ParseLeafPath_FileName_ReturnsBaseNameAndExtension(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, "book.xlsm")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book", base_name
    Assert.Equals ".xlsm", file_ext
    Assert.Equals "", path_suffix
End Sub

Public Sub Test_ParseLeafPath_PathSuffix_ReturnsSuffixSeparately(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, "book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book", base_name
    Assert.Equals ".xlsm", file_ext
    Assert.Equals "?q=a/b#frag/c", path_suffix
End Sub

Public Sub Test_ParseLeafPath_DotFile_ReturnsNoExtension(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim base_name As String
    Dim file_ext As String
    Dim path_suffix As String
    Call ParseLeafPath(base_name, file_ext, path_suffix, ".gitignore#frag")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals ".gitignore", base_name
    Assert.Equals "", file_ext
    Assert.Equals "#frag", path_suffix
End Sub

Public Sub Test_GetParentPath_Url_IgnoresQueryFragment(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetParentPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "https://example.com/sites/team", actual_value
End Sub

Public Sub Test_GetLeafFromPath_Url_IgnoresQueryFragment(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetLeafFromPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "book.xlsm", actual_value
End Sub

Public Sub Test_GetLeafFromPath_UrlExtension_IgnoresPathSuffix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetLeafFromPath("https://example.com/sites/team/book.xlsm?q=a/b#frag/c", BaseName:=False, Extension:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals ".xlsm", actual_value
End Sub

Public Sub Test_GetTypeString_PrimitiveObjectAndArray_ReturnsTypePrefix(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim long_arr(0 To 1) As Long

    Dim matrix(1 To 2, 1 To 3) As String

    Dim variant_arr As Variant
    variant_arr = Array("A", "B")

    Dim item_obj As Test_ObjectSetEquatableDouble
    Set item_obj = New Test_ObjectSetEquatableDouble

    ' Act / Assert
    Assert.Equals "String", GetTypeString("A")
    Assert.Equals "Long", GetTypeString(CLng(1))
    Assert.Equals "Long[]", GetTypeString(long_arr)
    Assert.Equals "Long[]", GetTypeString(long_arr, IncludeArrayRank:=True)
    Assert.Equals "Long[0:1]", GetTypeString(long_arr, IncludeArrayBounds:=True)
    Assert.Equals "String[,]", GetTypeString(matrix, IncludeArrayRank:=True)
    Assert.Equals "String[1:2,1:3]", GetTypeString(matrix, IncludeArrayBounds:=True)
    Assert.Equals "Variant[]", GetTypeString(variant_arr)
    Assert.Equals "Variant[0:1]", GetTypeString(variant_arr, IncludeArrayBounds:=True)
    Assert.Equals "Object@Test_ObjectSetEquatableDouble", GetTypeString(item_obj)
    Assert.Equals "IEquatable@Test_ObjectSetEquatableDouble", GetTypeString(item_obj, G_TYPED_VALUE_STRING_OBJECT_I_EQUATABLE)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueString_PrimitiveValues_ReturnsTypedStrings(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.Equals "String(A\,B\(C\)\\D)", GetTypedValueString("A,B(C)\D")
    Assert.Equals "String(A\tB)", GetTypedValueString("A" & vbTab & "B")
    Assert.Equals "Long(1)", GetTypedValueString(CLng(1))
    Assert.Equals "String(1)", GetTypedValueString(CStr(1))
    Assert.NotEquals GetTypedValueString(CLng(1)), GetTypedValueString(CStr(1))
    Assert.Equals "Boolean(True)", GetTypedValueString(CBool(True))
    Assert.Equals "Null()", GetTypedValueString(Null)
    Assert.Equals "Empty()", GetTypedValueString(Empty)
    Assert.Equals "Error(2042)", GetTypedValueString(CVErr(2042))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueString_PrimitiveArrays_ReturnsArrayTypeAndBounds(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim string_arr(1 To 2) As String
    string_arr(1) = "A"
    string_arr(2) = "B"

    Dim long_arr(0 To 1) As Long
    long_arr(0) = 1
    long_arr(1) = 2

    Dim variant_arr As Variant
    variant_arr = Array("1", CLng(1), Empty, Null, CVErr(2042))

    ' Act / Assert
    Assert.Equals "String[1:2](A,B)", GetTypedValueString(string_arr)
    Assert.Equals "Long[0:1](1,2)", GetTypedValueString(long_arr)
    Assert.Equals "Variant[0:4](String(1),Long(1),Empty(),Null(),Error(2042))", GetTypedValueString(variant_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueString_MultidimensionalArrays_ReturnsNestedItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim matrix(1 To 2, 1 To 3) As String
    matrix(1, 1) = "A"
    matrix(1, 2) = "B"
    matrix(1, 3) = "C"
    matrix(2, 1) = "D"
    matrix(2, 2) = "E"
    matrix(2, 3) = "F"

    Dim cube(1 To 2, 1 To 2, 1 To 2) As Long
    cube(1, 1, 1) = 111
    cube(1, 1, 2) = 112
    cube(1, 2, 1) = 121
    cube(1, 2, 2) = 122
    cube(2, 1, 1) = 211
    cube(2, 1, 2) = 212
    cube(2, 2, 1) = 221
    cube(2, 2, 2) = 222

    ' Act / Assert
    Assert.Equals "String[1:2,1:3]((A,B,C),(D,E,F))", GetTypedValueString(matrix)
    Assert.Equals "Long[1:2,1:2,1:2](((111,112),(121,122)),((211,212),(221,222)))", GetTypedValueString(cube)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetTypedValueString_JaggedArray_ReturnsRecursiveArrayItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim jagged_arr(0 To 1) As Variant
    jagged_arr(0) = Array("A", "B")
    jagged_arr(1) = Array("C", "D")

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueString(jagged_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:1](Variant[0:1](String(A),String(B)),Variant[0:1](String(C),String(D)))", actual_value
End Sub

Public Sub Test_GetTypedValueString_EmptyArrays_ReturnsArrayTypeOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim long_arr() As Long

    Dim variant_arr As Variant
    variant_arr = Array()

    ' Act / Assert
    Assert.Equals "Long[]()", GetTypedValueString(long_arr)
    Assert.Equals "Variant[0:-1]()", GetTypedValueString(variant_arr)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetMultiKey_DifferentPrimitiveTypes_ReturnsDifferentKeys(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act / Assert
    Assert.NotEquals GetMultiKey(CLng(1)), GetMultiKey(CStr(1))
    Assert.NotEquals GetMultiKey(CBool(True)), GetMultiKey(CStr("True"))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
End Sub

Public Sub Test_GetMultiKey_Null_ReturnsTypedKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Null()", actual_value
End Sub

Public Sub Test_GetMultiKey_ErrorValue_ReturnsTypedKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(CVErr(2042))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Error(2042)", actual_value
End Sub

Public Sub Test_GetMultiKey_TabInValue_UsesTypedValueEscaping(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey("A" & vbTab & "B", CLng(1))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "String(A\tB)" & vbTab & "Long(1)", actual_value
End Sub

Public Sub Test_GetMultiKey_ArrayValue_ReturnsTypedArrayKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim variant_arr As Variant
    variant_arr = Array("A", "B")

    ' Act
    Dim actual_value As String
    actual_value = GetMultiKey(variant_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:1](String(A),String(B))", actual_value
End Sub

Public Sub Test_GetTypedValueString_IEquatableObject_ReturnsIdentityString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ObjectSetEquatableDouble
    Set item_obj = New Test_ObjectSetEquatableDouble
    item_obj.IdentityKey = "A" & vbTab & "B@(C)"

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueString(item_obj, G_TYPED_VALUE_STRING_OBJECT_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IEquatable@Test_ObjectSetEquatableDouble(A\tB\@\(C\))", actual_value
End Sub

Public Sub Test_GetTypedValueString_IDuplicateCheckableObject_ReturnsDuplicateKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Test_ObjectSetDupCheckDouble
    Set item_obj = New Test_ObjectSetDupCheckDouble
    item_obj.DuplicateKey = "Key=1"

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueString(item_obj, G_TYPED_VALUE_STRING_OBJECT_DUPLICATE_CHECKABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IDuplicateCheckable@Test_ObjectSetDupCheckDouble(Key\=1)", actual_value
End Sub

Public Sub Test_GetTypedValueString_IEquatableArray_ReturnsPolicyAndDeclaredArrayType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "A"

    Dim second_item As Test_ObjectSetEquatableDouble
    Set second_item = New Test_ObjectSetEquatableDouble
    second_item.IdentityKey = "B"

    Dim item_arr(0 To 1) As IEquatable
    Set item_arr(0) = first_item
    Set item_arr(1) = second_item

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueString(item_arr, G_TYPED_VALUE_STRING_OBJECT_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "IEquatable@IEquatable[0:1](IEquatable@Test_ObjectSetEquatableDouble(A),IEquatable@Test_ObjectSetEquatableDouble(B))", actual_value
End Sub

Public Sub Test_GetTypedValueString_IEquatableModeWithUnsupportedObject_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim item_obj As Object
    Set item_obj = CreateObject("Scripting.Dictionary")

    ' Act
    Dim actual_value As String
    actual_value = GetTypedValueString(item_obj, G_TYPED_VALUE_STRING_OBJECT_I_EQUATABLE)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Function GetTypedValueString", Err.Source
End Sub
