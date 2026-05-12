Attribute VB_Name = "Test_WorkbookService"
Option Explicit

' #############################################################################
'!
'! @brief
'! WorkbookService クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

' -----------------------------------------------------------------------------
' GetThisWorkbookDirectoryPath
' -----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookDirectoryPath_SavedThisWorkbook_ReturnsWorkbookFolder(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService
    
    Dim expected_path As String
    expected_path = GetParentPath(ThisWorkbook.FullName)
    
    ' 実行
    Dim actual_path As String
    actual_path = book_srv.GetThisWorkbookDirectoryPath()
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

Private Function pPrepareTestSheet(ByRef SheetName As String) As Worksheet
    Dim target_sheet As Worksheet

    On Error Resume Next
        Set target_sheet = ThisWorkbook.Worksheets(SheetName)
    On Error GoTo 0

    If Not target_sheet Is Nothing Then
        Call target_sheet.Cells.Clear
    Else
        Set target_sheet = ThisWorkbook.Worksheets.Add
        target_sheet.Name = SheetName
    End If

    Set pPrepareTestSheet = target_sheet
End Function

' -----------------------------------------------------------------------------
' ExistsWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_ExistsWorksheet_NotExistsWorkbook_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService
    
    ' Act
    Dim actual_value As Boolean
    actual_value = book_srv.ExistsWorksheet("Sheet1", Book:="NotExistsWorkbook")
    
    ' Assert
    Assert.Equals False, actual_value
End Sub

Public Sub Test_ExistsWorksheet_NotExistsWorksheet_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService
    
    ' Act
    Dim actual_value As Boolean
    actual_value = book_srv.ExistsWorksheet("NotExistsWorksheet", Book:=ThisWorkbook.Name)
    
    ' Assert
    Assert.Equals False, actual_value
End Sub

Public Sub Test_ExistsWorksheet_ExistsWorksheet_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim target_sheet As Worksheet
    Set target_sheet = pPrepareTestSheet("test_output")

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService
    
    ' Act
    Dim actual_value As Boolean
    actual_value = book_srv.ExistsWorksheet("test_output")
    
    ' Assert
    Assert.Equals True, actual_value
End Sub

' -----------------------------------------------------------------------------
' GetAllWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_GetAllWorksheet_Call_ReturnsAllWorksheets(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService
    
    ' Act
    Dim actual_value() As String
    actual_value = book_srv.GetAllWorksheet
    
    ' Assert
    Assert.EqualsNumeric ThisWorkbook.Worksheets.Count, UBound(actual_value) - LBound(actual_value) + 1
End Sub

' -----------------------------------------------------------------------------
' AddWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_AddWorksheet_NegativeSheetIndexMinusOne_AddsAfterLastSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 3
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop

    Dim book_name As String
    book_name = target_book.Name

    Dim initial_sheet_count As Long
    initial_sheet_count = target_book.Worksheets.Count

    ' Act
    Dim actual_sheet_name As String
    Err.Clear
    actual_sheet_name = book_srv.AddWorksheet(Sheet:="test_negative_index", Book:=book_name, SheetIndex:=-1, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_index As Long
    If err_num = 0 Then
        actual_index = target_book.Worksheets(actual_sheet_name).Index
    End If

    ' Cleanup
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.EqualsNumeric initial_sheet_count + 1, actual_index
End Sub

' -----------------------------------------------------------------------------
' SaveWorkbook
' -----------------------------------------------------------------------------

Public Sub Test_SaveWorkbook_VisibleWorkbook_KeepsWindowVisible(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Call InitializeCommonService
    Call InitializeFileSystemService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_workbook_service_visible.xlsx")
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Act
    book_name = book_srv.SaveWorkbook(book_name, file_path, Force:=True)

    Dim actual_visible As Boolean
    actual_visible = Workbooks(book_name).Windows(1).Visible

    ' Cleanup
    Call book_srv.CloseWorkbook(book_name)
    If FsSrv.IsFile(file_path) Then Call FsSrv.RemoveFile(file_path, Force:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_visible
End Sub

Public Sub Test_SaveWorkbook_SaveAsFailure_RaisesError(ByVal Assert As UnitTestAssert)
    ' Arrange
    Call InitializeCommonService

    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim book_name As String
    book_name = target_book.Name

    Dim file_path As String
    file_path = JoinPath(WbSrv.GetThisWorkbookDirectoryPath(), "tmp_invalid:name.xlsx")

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String

    ' Act
    On Error GoTo ON_SAVE_ERROR
        Call book_srv.SaveWorkbook(book_name, file_path, Force:=True)
    On Error GoTo 0
    GoTo AFTER_SAVE

ON_SAVE_ERROR:
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description
    Resume AFTER_SAVE

AFTER_SAVE:
    On Error Resume Next
    Call book_srv.CloseWorkbook(book_name, Force:=True)
    On Error GoTo 0

    ' Assert
    Assert.ErrorRaised 0, err_num, err_source, err_desc
End Sub

