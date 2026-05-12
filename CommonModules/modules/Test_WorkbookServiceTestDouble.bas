Attribute VB_Name = "Test_WorkbookServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! WorkbookServiceTestDouble クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private TUtl As New UnitTestUtils

' ----------------------------------------------------------------------------
' GetThisWorkbookName のテスト
' ----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookName_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    wb_srv_stub.GetThisWorkbookName_Value = "Tool.xlsm"
    
    Dim wb_srv As IWorkbookService
    Set wb_srv = wb_srv_stub
    
    ' 実行
    Dim actual_name As String
    actual_name = wb_srv.GetThisWorkbookName()
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Tool.xlsm", actual_name
End Sub
' ----------------------------------------------------------------------------
' GetThisWorkbookDirectoryPath のテスト
' ----------------------------------------------------------------------------

Public Sub Test_GetThisWorkbookDirectoryPath_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    wb_srv_stub.GetThisWorkbookDirectoryPath_Value = "C:\Temp"
    
    Dim wb_srv As IWorkbookService
    Set wb_srv = wb_srv_stub
    
    ' 実行
    Dim actual_path As String
    actual_path = wb_srv.GetThisWorkbookDirectoryPath()
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "C:\Temp", actual_path
End Sub

' ----------------------------------------------------------------------------
' ExistsWorkbook のテスト
' ----------------------------------------------------------------------------

Public Sub Test_ExistsWorkbook_SetValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "TestBook.xlsm"
    
    Call TUtl.SetValue(wb_srv_stub.ExistsWorkbook_Values, True, wb_name)
    
    ' Act
    Dim actual_result As Boolean
    actual_result = wb_srv_stub.ExistsWorkbook(wb_name)
    
    ' Assert
    Assert.Equals True, actual_result
End Sub

Public Sub Test_ExistsWorkbook_SetValueButNotMatch_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_srv_stub As WorkbookServiceTestDouble
    Set wb_srv_stub = New WorkbookServiceTestDouble
    
    Call TUtl.SetValue(wb_srv_stub.ExistsWorkbook_Values, True, "TestBook.xlsm")
    
    ' Act
    Dim actual_result As Boolean
    actual_result = wb_srv_stub.ExistsWorkbook("other_name")
    
    ' Assert
    Assert.Equals False, actual_result
End Sub

Public Sub Test_ExistsWorkbook_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    ' Key 未登録
    
    ' Act
    Dim actual_result As Boolean
    actual_result = wb_double.ExistsWorkbook("UnknownBook.xlsx")
    
    ' Assert - 登録がないなら False
    Assert.IsFalse actual_result
End Sub

' ----------------------------------------------------------------------------
' GetAllWorkbook のテスト
' ----------------------------------------------------------------------------
Public Sub Test_GetAllWorkbook_ReturnsSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    Dim expected_arr(0 To 2) As String
    expected_arr(0) = "BookA"
    expected_arr(1) = "BookB"
    expected_arr(2) = "BookC"
    
    ' 事前にクラスの配列プロパティへ設定
    Call TUtl.SetValue(wb_double.GetAllWorkbook_Values, expected_arr, Nothing)
    
    ' Act
    Dim actual_arr() As String
    actual_arr = wb_double.GetAllWorkbook
    
    ' Assert
    Assert.Equals "BookA", actual_arr(0)
    Assert.Equals "BookB", actual_arr(1)
    Assert.Equals "BookC", actual_arr(2)
End Sub

' ----------------------------------------------------------------------------
' OpenWorkbook のテスト
' ----------------------------------------------------------------------------
Public Sub Test_OpenWorkbook_RecordsInDictionaryAndReturnsBookN(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    Dim file_path As String
    file_path = "C:\Data\MyWorkbook.xlsx"
    ' 初期 BookCount=1 の想定
    
    ' Act
    Dim actual_name As String
    actual_name = wb_double.OpenWorkbook(file_path)
    
    ' Assert
    ' 1) 戻り値は "Book1" (BookCount=1)
    Assert.Equals "Book1", actual_name
    
    ' 2) BookCount が 2 に増えている
    Assert.EqualsNumeric 2, wb_double.BookCount
    
    ' 3) 辞書に ("Book1") が登録されているか
    '    ここでは file_path がキー引数, 値="Book1" で保存される
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_double.OpenWorkbook_Results, file_path)
    Assert.Equals "Book1", stored_val
End Sub

' ----------------------------------------------------------------------------
' SaveWorkbook のテスト
' ----------------------------------------------------------------------------
Public Sub Test_SaveWorkbook_SetsDictionaryAndReturnsLeaf(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "Book1"
    Dim file_path As String
    file_path = "C:\Temp\Sub\Output.xlsx"
    
    ' Act
    Dim actual_leaf As String
    actual_leaf = wb_double.SaveWorkbook(wb_name, file_path, True)
    
    ' Assert
    ' 1) 戻り値 = "Output.xlsx" の葉
    Assert.Equals "Output.xlsx", actual_leaf
    
    ' 2) 辞書に格納されているか
    '    引数 (wb_name, file_path, True) をキー、値 = "Output.xlsx"
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_double.SaveWorkbook_Results, wb_name, file_path, True)
    Assert.Equals "Output.xlsx", stored_val
End Sub

' ----------------------------------------------------------------------------
' IsSaved のテスト
' ----------------------------------------------------------------------------
Public Sub Test_IsSaved_RegisteredValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "BookX"
    
    ' 登録
    Call TUtl.SetValue(wb_double.IsSaved_Values, True, wb_name)
    
    ' Act
    Dim actual_saved As Boolean
    actual_saved = wb_double.IsSaved(wb_name)
    
    ' Assert
    Assert.IsTrue actual_saved
End Sub

Public Sub Test_IsSaved_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    ' Act
    Dim actual_saved As Boolean
    actual_saved = wb_double.IsSaved("NoSuchBook")
    
    ' Assert
    Assert.IsFalse actual_saved
End Sub

' ----------------------------------------------------------------------------
' CloseWorkbook のテスト
' ----------------------------------------------------------------------------
Public Sub Test_CloseWorkbook_WithFalse_RecordsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "TempBook"
    
    ' Act
    wb_double.CloseWorkbook wb_name, False
    
    ' Assert
    ' ここでは引数 (wb_name, False) がキー、値=True
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_double.ColoseWorkbook_Results, wb_name, False)
    Assert.Equals True, stored_val
End Sub

Public Sub Test_CloseWorkbook_NoForceArg_RecordsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_double As WorkbookServiceTestDouble
    Set wb_double = New WorkbookServiceTestDouble
    
    ' Act (Force 省略 → 既定 True)
    wb_double.CloseWorkbook "NoArgBook"
    
    ' Assert
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_double.ColoseWorkbook_Results, "NoArgBook", True)
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' ExistsWorksheet のテスト
' ----------------------------------------------------------------------------
Public Sub Test_ExistsWorksheet_WithSetValue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    Dim ws_name As String
    ws_name = "SheetA"
    Dim wb_name As String
    wb_name = "MyBook.xlsm"
    
    ' 登録 (Sheet, Book) をキーに True
    Call TUtl.SetValue(wb_stub.ExistsWorksheet_Values, True, ws_name, wb_name)
    
    ' Act
    Dim actual_result As Boolean
    actual_result = wb_stub.ExistsWorksheet(ws_name, wb_name)
    
    ' Assert
    Assert.Equals True, actual_result
End Sub

Public Sub Test_ExistsWorksheet_NotRegistered_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' Act
    Dim actual_result As Boolean
    actual_result = wb_stub.ExistsWorksheet("NoSuchSheet", "NoSuchBook")
    
    ' Assert
    Assert.IsFalse actual_result
End Sub

Public Sub Test_ExistsWorksheet_PartialMatch_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' 同じシート名で別のブックとか、同じブック名で別のシート名とかで合わないケース
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' 登録
    Call TUtl.SetValue(wb_stub.ExistsWorksheet_Values, True, "SheetX", "Book1.xlsm")
    
    ' Act
    Dim result_sameSheet_otherBook As Boolean
    result_sameSheet_otherBook = wb_stub.ExistsWorksheet("SheetX", "Book2.xlsm")
    
    Dim result_otherSheet_sameBook As Boolean
    result_otherSheet_sameBook = wb_stub.ExistsWorksheet("SheetY", "Book1.xlsm")
    
    ' Assert
    Assert.IsFalse result_sameSheet_otherBook
    Assert.IsFalse result_otherSheet_sameBook
End Sub

' ----------------------------------------------------------------------------
' GetAllWorksheet のテスト
' ----------------------------------------------------------------------------
Public Sub Test_GetAllWorksheet_WithSetValueArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "SomeBook.xlsm"
    
    Dim expected_ws(0 To 2) As String
    expected_ws(0) = "SheetA"
    expected_ws(1) = "SheetB"
    expected_ws(2) = "SheetC"
    
    ' 登録 (Book) -> 配列
    Call TUtl.SetValue(wb_stub.GetAllWorksheet_Values, expected_ws, wb_name)
    
    ' Act
    Dim actual_arr() As String
    actual_arr = wb_stub.GetAllWorksheet(wb_name)
    
    ' Assert
    Assert.Equals "SheetA", actual_arr(0)
    Assert.Equals "SheetB", actual_arr(1)
    Assert.Equals "SheetC", actual_arr(2)
End Sub

' ----------------------------------------------------------------------------
' AddWorksheet のテスト
' ----------------------------------------------------------------------------
Public Sub Test_AddWorksheet_WithoutName_ReturnsSheetX(ByVal Assert As UnitTestAssert)
    ' 省略時は "Sheet" & AddSheetCount
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    Dim wb_name As String
    wb_name = "DataBook.xlsm"
    
    ' Act
    Dim actual_sheet As String
    actual_sheet = wb_stub.AddWorksheet(, wb_name)
    
    ' Assert
    ' 初期 AddSheetCount=1 => 戻り値は "Sheet1"
    Assert.Equals "Sheet1", actual_sheet
    Assert.EqualsNumeric 2, wb_stub.AddSheetCount   ' カウントが+1される
    
    ' Dictionary に(Sheet="", wb_name, SheetIndex=0, Before=False)をキー, 値="Sheet1"
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_stub.AddWorksheet_Results, "", wb_name, 0&, False)
    Assert.Equals "Sheet1", stored_val
End Sub

Public Sub Test_AddWorksheet_WithName_ReturnsSameName(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' Act
    Dim actual_sheet As String
    actual_sheet = wb_stub.AddWorksheet("CustomSheet", "MyBook.xlsm", 2, True)
    
    ' Assert
    ' 戻り値はそのまま "CustomSheet"
    Assert.Equals "CustomSheet", actual_sheet
    ' AddSheetCount は変わらない(省略時にのみ++想定)
    Assert.EqualsNumeric 1, wb_stub.AddSheetCount
    
    ' Dictionary 登録確認
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_stub.AddWorksheet_Results, "CustomSheet", "MyBook.xlsm", 2&, True)
    Assert.Equals "CustomSheet", stored_val
End Sub

' ----------------------------------------------------------------------------
' RemoveWorksheet のテスト
' ----------------------------------------------------------------------------
Public Sub Test_RemoveWorksheet_RecordsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' Act
    wb_stub.RemoveWorksheet "SheetToDelete", "Book99.xlsx"
    
    ' Assert
    ' 辞書に ( "SheetToDelete","Book99.xlsx" ) をキー, 値=True
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_stub.RemoveWorksheet_Results, "SheetToDelete", "Book99.xlsx")
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' CopyWorksheet のテスト
' ----------------------------------------------------------------------------
Public Sub Test_CopyWorksheet_NoDestName_ReturnsSheetX(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' Act
    Dim new_sheet As String
    new_sheet = wb_stub.CopyWorksheet("SrcSheet", "SrcBook.xlsm") ' DestinationWorksheetName省略
    
    ' Assert
    ' 初期 CopySheetCount=1 => "Sheet1"
    Assert.Equals "Sheet1", new_sheet
    Assert.EqualsNumeric 2, wb_stub.CopySheetCount  ' +1 されたか
    
    ' Dictionary には (SourceWorksheetName="SrcSheet", SourceBook="SrcBook.xlsm",DestName="",DestBook="",SheetIndex=0,Before=False) → 値="Sheet1"
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_stub.CopyWorksheet_Results, "SrcSheet", "SrcBook.xlsm", "", "", 0&, False)
    Assert.Equals "Sheet1", stored_val
End Sub

Public Sub Test_CopyWorksheet_WithAllArgs_ReturnsDestName(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim wb_stub As WorkbookServiceTestDouble
    Set wb_stub = New WorkbookServiceTestDouble
    
    ' Act
    Dim copied_name As String
    copied_name = wb_stub.CopyWorksheet( _
                        "OriginalSheet", _
                        "SourceBook.xlsm", _
                        "NewSheetName", _
                        "DestBook.xlsm", _
                        5, _
                        True)
    
    ' Assert
    ' 戻り値は "NewSheetName"
    Assert.Equals "NewSheetName", copied_name
    ' CopySheetCountはインクリメントしない
    Assert.EqualsNumeric 1, wb_stub.CopySheetCount
    
    ' Dictionary にはキー ( "OriginalSheet","SourceBook.xlsm","NewSheetName","DestBook.xlsm",5,True ) → 値="NewSheetName"
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(wb_stub.CopyWorksheet_Results, "OriginalSheet", "SourceBook.xlsm", "NewSheetName", "DestBook.xlsm", 5&, True)
    Assert.Equals "NewSheetName", stored_val
End Sub

