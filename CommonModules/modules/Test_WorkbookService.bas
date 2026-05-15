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

Public Sub Test_AddWorksheet_NonActiveWorkbook_RestoresActiveWorkbookAndSheets(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_active"
    target_book.Worksheets(2).Name = "target_anchor"
    Call target_book.Worksheets("target_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Dim actual_sheet_name As String
    Err.Clear
    actual_sheet_name = book_srv.AddWorksheet(Sheet:="added_sheet", Book:=target_book_name, SheetIndex:=2, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "target_active", actual_target_active_sheet_name
End Sub

Public Sub Test_AddWorksheet_NameFailure_RestoresActiveWorkbookAndSheets(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "target_active"
    target_book.Worksheets(2).Name = "target_anchor"
    Call target_book.Worksheets("target_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Err.Clear
    Call book_srv.AddWorksheet(Sheet:="invalid:name", Book:=target_book_name, SheetIndex:=2, Before:=False)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    actual_active_book_name = ActiveWorkbook.Name
    actual_active_sheet_name = ActiveSheet.Name
    actual_target_active_sheet_name = target_book.ActiveSheet.Name

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "target_active", actual_target_active_sheet_name
End Sub

' -----------------------------------------------------------------------------
' RemoveWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_RemoveWorksheet_LastSheet_RestoresActiveWorkbook(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add

    Dim old_display_alerts As Boolean
    old_display_alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Do While target_book.Worksheets.Count > 1
        Call target_book.Worksheets(target_book.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = old_display_alerts

    target_book.Worksheets(1).Name = "remove_me"

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Err.Clear
    Call book_srv.RemoveWorksheet(Sheet:="remove_me", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_sheet_count As Long
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_sheet_count = target_book.Worksheets.Count
    End If

    ' Cleanup
    Application.DisplayAlerts = old_display_alerts
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.EqualsNumeric 1, actual_target_sheet_count
End Sub

Public Sub Test_RemoveWorksheet_NonActiveWorkbook_RestoresTargetActiveSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

    Dim target_book As Workbook
    Set target_book = Workbooks.Add
    Do While target_book.Worksheets.Count < 2
        Call target_book.Worksheets.Add(After:=target_book.Worksheets(target_book.Worksheets.Count))
    Loop
    target_book.Worksheets(1).Name = "keep_active"
    target_book.Worksheets(2).Name = "delete_me"
    Call target_book.Worksheets("keep_active").Activate

    Dim control_book As Workbook
    Set control_book = Workbooks.Add
    control_book.Worksheets(1).Name = "control_active"
    Call control_book.Worksheets("control_active").Activate

    Dim target_book_name As String
    target_book_name = target_book.Name

    Dim control_book_name As String
    control_book_name = control_book.Name

    ' Act
    Err.Clear
    Call book_srv.RemoveWorksheet(Sheet:="delete_me", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    Dim actual_target_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
        actual_target_active_sheet_name = target_book.ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals control_book_name, actual_active_book_name
    Assert.Equals "control_active", actual_active_sheet_name
    Assert.Equals "keep_active", actual_target_active_sheet_name
End Sub

' -----------------------------------------------------------------------------
' ActivateWorksheet
' -----------------------------------------------------------------------------

Public Sub Test_ActivateWorksheet_NonActiveWorkbook_ActivatesTargetBookAndSheet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim book_srv As IWorkbookService
    Set book_srv = New WorkbookService

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

    ' Act
    Err.Clear
    Call book_srv.ActivateWorksheet("target_activate", Book:=target_book_name)

    Dim err_num As Long
    Dim err_source As String
    Dim err_desc As String
    err_num = Err.Number
    err_source = Err.Source
    err_desc = Err.Description

    Dim actual_active_book_name As String
    Dim actual_active_sheet_name As String
    If err_num = 0 Then
        actual_active_book_name = ActiveWorkbook.Name
        actual_active_sheet_name = ActiveSheet.Name
    End If

    ' Cleanup
    Call control_book.Close(SaveChanges:=False)
    Call target_book.Close(SaveChanges:=False)
    On Error GoTo 0

    ' Assert
    If Not Assert.ErrorNotRaised(0, err_num, err_source, err_desc) Then Exit Sub
    Assert.Equals target_book_name, actual_active_book_name
    Assert.Equals "target_activate", actual_active_sheet_name
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

