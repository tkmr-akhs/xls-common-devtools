Attribute VB_Name = "Lib_UnitTest"
Option Explicit

' #############################################################################
'!
'! @brief
'! ユニット テストのエントリ ポイントです。
'!
' #############################################################################

Const C_SHEET_NAME As String = "UNIT_TEST_SHEET"
Const C_NEW_BOOK As Boolean = False
Const C_COL_MOD As Long = 1
Const C_COL_SUB As Long = 2
Const C_COL_OKNG As Long = 3
Const C_COL_DESC As Long = 4
Const C_COL_END As Long = 4
Const C_COL_BTN As Long = 5
Const C_COL_MBTN As Long = 6
Const C_SUB_MAIN As String = "UnitTestMain"
Const C_COLOR_RESET_BG As Long = &H404040
Const C_COLOR_RESET_FG As Long = &HC0C0C0


'* ユニット テストのエントリ ポイントです。
'*
'* @details
'* ユニット テストのエントリ ポイントです。
'* このプロジェクト内の全モジュールを走査し、Test_ で始まる UnitTestAssert 型を引数とするサブ プロシージャを見つけて実行します。
'*
'* ユニット テスト例:
'* @code
'* Sub Test_SomeFunction(Assert As UnitTestAssert)
'*     ' Arrange
'*     Dim some_obj As SomeObject: t some_obj = New SomeObject
'*
'*     ' Act
'*     Dim actual_value As Long
'*     actual_value = some_obj.SomeFunction()
'*
'*     ' Assert
'*     Assert.Equals 0, actual_value)
'* End Sub
'* @endcode
Public Sub UnitTestMain()
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    Call app_state.DisableUpdates(StopEvents:=False)

    ' テスト番号の取得
    Dim test_idx As Long: test_idx = 0

    Dim caller_str As String

    On Error Resume Next
    caller_str = CStr(Application.Caller)
    If Err.Number <> 0 Then
        Err.Clear
        caller_str = ""
    End If
    On Error GoTo ON_ERROR

    If IsNumeric(caller_str) Then
        test_idx = CLng(caller_str)
    Else
        test_idx = 0
    End If

    ' テスト結果の出力先を準備
    Dim result_sheet As Worksheet
    Set result_sheet = pPrepareResultSheet(test_idx)

    ' テスト実行
    If test_idx = 0 Then
        Call pRunAllTest(result_sheet)
    Else
        Call pRunTestCore(result_sheet, result_sheet.Cells(test_idx, C_COL_MOD), result_sheet.Cells(test_idx, C_COL_SUB), test_idx)
    End If

    Call app_state.Restore
    On Error GoTo 0
    Exit Sub

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub

Private Function pBuildWorkbookMacroName(ByVal MacroName As String) As String
    pBuildWorkbookMacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & MacroName
End Function

Private Function pPrepareResultSheet(ByVal TestIndex As Long) As Worksheet
    Dim result_sheet As Worksheet
    If C_NEW_BOOK Then
        ' 新しいブックに出力
        Dim result_book As Workbook
        Set result_book = Workbooks.Add
        
        Set result_sheet = result_book.Worksheets(1)
        Call AddButton(result_sheet, 1, C_COL_MBTN, "すべて実行", pBuildWorkbookMacroName(C_SUB_MAIN), "Button_UnitTestMain")
    Else
        ' ThisWorkbook に出力
        If TestIndex = 0 Then
            ' 全テスト実行の場合
            
            ' シート作成を試行
            On Error Resume Next
            Set result_sheet = ThisWorkbook.Worksheets(C_SHEET_NAME)
            On Error GoTo 0
            If result_sheet Is Nothing Then
                ' 新規シート
                Set result_sheet = ThisWorkbook.Worksheets.Add()
                result_sheet.Name = C_SHEET_NAME
            Else
                ' 既存シート
                result_sheet.Columns(C_COL_OKNG).Interior.Color = C_COLOR_RESET_BG
                result_sheet.Columns(C_COL_OKNG).Font.Color = C_COLOR_RESET_FG
                Call result_sheet.Cells.ClearContents
                Call ClearButton(result_sheet)
            End If
            Call AddButton(result_sheet, 1, C_COL_MBTN, "すべて実行", pBuildWorkbookMacroName(C_SUB_MAIN), "Button_UnitTestMain")
        Else
            Set result_sheet = ThisWorkbook.Worksheets(C_SHEET_NAME)
            result_sheet.Cells(TestIndex, C_COL_OKNG).Interior.Color = C_COLOR_RESET_BG
            result_sheet.Cells(TestIndex, C_COL_OKNG).Font.Color = C_COLOR_RESET_FG
            Call result_sheet.Cells(TestIndex, C_COL_OKNG).ClearContents
            Call result_sheet.Cells(TestIndex, C_COL_DESC).ClearContents
        End If
    End If
    
    result_sheet.Cells(1, C_COL_MOD).Value = "Category"
    result_sheet.Cells(1, C_COL_SUB).Value = "Test Item"
    result_sheet.Cells(1, C_COL_OKNG).Value = "Result"
    result_sheet.Cells(1, C_COL_DESC).Value = "Description"
    Call result_sheet.Range(RangeAddress(StartColumn:=1, FinishColumn:=C_COL_END)).AutoFilter
    
    Set pPrepareResultSheet = result_sheet
End Function

Private Sub pRunAllTest(ByVal ResultSheet As Worksheet)
    ' VBIDE のプロジェクト オブジェクトを取得
    Dim vb_proj As Variant 'VBIDE.VBProject
    Set vb_proj = ThisWorkbook.VBProject

    ' テスト サブ プロシージャを抽出するための正規表現の準備
    Dim sub_re As RegExp
    Set sub_re = New RegExp
    sub_re.Pattern = "(?:^Sub|^[^']*Public.*\sSub)\s(Test_[^\s(]+)\s*\([^,]*\s+As\s+UnitTestAssert.*\).*$"
    
    ' プロジェクト オブジェクトのコンポーネントすべてについて処理
    Dim row_idx As Long: row_idx = 2
    Dim vb_comp As Variant 'VBIDE.VBComponent
    For Each vb_comp In vb_proj.VBComponents
        ' モジュール名の取得
        Dim mod_name As String
        On Error Resume Next
        mod_name = vb_comp.Name
        If Err.Number <> 0 Then
            Debug.Print "<" & row_idx & "> [&H" & Hex(Err.Number) & "] " & Err.Source & " | " & Err.Description
            Err.Clear
            On Error GoTo 0
        Else
            On Error GoTo 0
            
            'Debug.Print "Search " & mod_name
            
            ' コンポーネントのコード モジュールを得る
            Dim vb_comp_code As Variant 'VBIDE.CodeModule
            Set vb_comp_code = vb_comp.CodeModule
            
            ' コード モジュールのすべての行を処理する
            Dim line_idx As Long
            For line_idx = 1 To vb_comp_code.CountOfLines
                Dim code_line As String
                code_line = vb_comp_code.Lines(line_idx, 1)
                
                ' 行が正規表現にマッチするかチェック
                Dim match_result As MatchCollection
                Set match_result = sub_re.Execute(code_line)
                
                If 0 < match_result.Count Then
                    'Debug.Print "Found " & TestName
                        
                    Dim sub_name As String
                    sub_name = match_result.Item(0).SubMatches(0)
                    
                    ' テストを実行する
                    Call pRunTestCore(ResultSheet, mod_name, sub_name, row_idx)
                    
                    ' ボタンを追加する
                    Call AddButton(ResultSheet, row_idx, C_COL_BTN, "再実行", pBuildWorkbookMacroName(C_SUB_MAIN), row_idx)
                    
                    ' 行を進める
                    row_idx = row_idx + 1
                End If
            Next line_idx
        End If
    Next vb_comp
End Sub

Private Sub pRunTestCore(ByVal ResultSheet As Worksheet, ByVal TestModName As String, ByVal TestSubName As String, ByVal RowIndex As Long)
    ' Assert オブジェクトを準備する
    Dim assert_obj As UnitTestAssert: Set assert_obj = New UnitTestAssert
    
    ' テスト サブ プロシージャを実行する
    Application.Run pBuildWorkbookMacroName(TestModName & "." & TestSubName), assert_obj
    
    ' 実行結果を書き出す
    Call pWriteResult(ResultSheet, RowIndex, TestModName, TestSubName, assert_obj)
End Sub

Private Sub pWriteResult(ByVal ResultSheet As Worksheet, ByVal RowIndex As Long, ByVal ModuleName As String, ByVal TestSubName As String, ByVal AssertObject As UnitTestAssert)
    ResultSheet.Cells(RowIndex, C_COL_MOD).Value = ModuleName
    ResultSheet.Cells(RowIndex, C_COL_SUB).Value = TestSubName
        ResultSheet.Cells(RowIndex, C_COL_DESC).Value = AssertObject.ResultMessage
    If AssertObject.IsFailed Then
        ' failed
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Value = "NG"
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Interior.Color = RGB(255, 128, 128)
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Font.Color = RGB(64, 64, 64)
    Else
        ' passed
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Value = "OK"
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Interior.Color = RGB(128, 255, 128)
        ResultSheet.Cells(RowIndex, C_COL_OKNG).Font.Color = RGB(64, 64, 64)
    End If
End Sub

