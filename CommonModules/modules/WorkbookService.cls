VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorkbookService

Private Const C_VBEXT_CT_DOCUMENT As Long = 100

' #############################################################################
'!
'! @brief
'! Excel ファイルのオープン、クローズ、書き込みなどを管理するサービスの実装です。
'!
' #############################################################################

'* ThisWorkbook のブック名を得ます。
'*
'* @return ThisWorkbook のブック名。
'*
'* @details
'* 呼び出し元が ThisWorkbook を直接参照せずに、自ブック名を取得するために使用します。
Public Function GetThisWorkbookName() As String
    GetThisWorkbookName = ThisWorkbook.Name
End Function
Private Function IWorkbookService_GetThisWorkbookName() As String
    IWorkbookService_GetThisWorkbookName = GetThisWorkbookName()
End Function

'* ThisWorkbook を基準にしたディレクトリ パスを得ます。
'*
'* @return 絶対パスのディレクトリ パス。
'*
'* @details
'* ThisWorkbook が保存済みの場合は、その保存先ディレクトリを返します。
'* ThisWorkbook が未保存の場合は、CurDir の結果を絶対パスにして返します。
Public Function GetThisWorkbookDirectoryPath() As String
    Dim dir_path As String
    
    dir_path = ThisWorkbook.Path
    If dir_path = "" Then
        dir_path = CurDir()
    End If
    
    GetThisWorkbookDirectoryPath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(dir_path)
End Function
Private Function IWorkbookService_GetThisWorkbookDirectoryPath() As String
    IWorkbookService_GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath()
End Function

' ワークブックが存在するかを確認します。
'
' @param Book ワークブック名。
' @return ワークブックが存在するか。
'
' @details
' ワークブック名が存在するかを確認します。
Public Function ExistsWorkbook(ByVal Book As String) As Boolean
    Dim target_book As Workbook
    On Error Resume Next
        Set target_book = Workbooks(Book)
    Err.Clear
    On Error GoTo 0
    
    If target_book Is Nothing Then
        ExistsWorkbook = False
    Else
        ExistsWorkbook = True
    End If
End Function
Private Function IWorkbookService_ExistsWorkbook(ByVal Book As String) As Boolean
    IWorkbookService_ExistsWorkbook = ExistsWorkbook(Book)
End Function

' ThisWorkbook 以外のワークブックの一覧を得ます。
'
' @return ワークブックの一覧。
'
' @details
' ThisWorkbook を除くワークブック名の一覧を、文字列の配列で得ます。
Public Function GetAllWorkbook() As String()
    Dim result_arr() As String
    
    Dim this_book As String
    this_book = ThisWorkbook.Name
    
    Dim book_item As Workbook
    For Each book_item In Workbooks
        If Not book_item Is ThisWorkbook Then Call pAddItemToStringArray(result_arr, book_item.Name)
    Next book_item
    GetAllWorkbook = result_arr
End Function
Private Function IWorkbookService_GetAllWorkbook() As String()
    IWorkbookService_GetAllWorkbook = GetAllWorkbook()
End Function

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

' ワークシートが存在するかを確認します。
'
' @param Sheet ワークシート名。
' @param Book ワークブック名。省略した場合は ThisWorkbook 内を確認します。
' @return ワークシートが存在するか。
'
' @details
' ワークブック内にワークシートが存在するかを確認します。
' 指定された名前のワークブックがない場合もエラーにはならず False が返ります。
Public Function ExistsWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        On Error Resume Next
            Set target_book = Workbooks(Book)
        Err.Clear
        On Error GoTo 0
    
        If target_book Is Nothing Then
            ExistsWorksheet = False
            Exit Function
        End If
    End If
    
    Dim target_sheet As Worksheet
    On Error Resume Next
        Set target_sheet = target_book.Worksheets(Sheet)
    Err.Clear
    On Error GoTo 0
    
    If target_sheet Is Nothing Then
        ExistsWorksheet = False
    Else
        ExistsWorksheet = True
    End If
End Function
Private Function IWorkbookService_ExistsWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    IWorkbookService_ExistsWorksheet = ExistsWorksheet(Sheet, Book:=Book)
End Function

' ワークシートの一覧を得ます。
'
' @param Book ワークブック名。省略した場合は ThisWorkbook 内を確認します。
' @return ワークシートの一覧。
'
' @details
' ワークブック内のワークシート名の一覧を、文字列の配列で得ます。
Public Function GetAllWorksheet(Optional ByVal Book As String = "") As String()
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Workbooks(Book)
    End If
    
    Dim result_arr() As String
    
    Dim sheet_item As Worksheet
    For Each sheet_item In target_book.Worksheets
        Call pAddItemToStringArray(result_arr, sheet_item.Name)
    Next sheet_item
    
    GetAllWorksheet = result_arr
End Function
Private Function IWorkbookService_GetAllWorksheet(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetAllWorksheet = GetAllWorksheet(Book:=Book)
End Function

' ワークブックを開きます。
'
' @param FilePath [省略可] ファイルパス。
' @return ブック名。
'
' @details
' ワークブックを開きます。ファイルパスを省略すると新規作成となります。
' 元ファイルはロックしません。
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    Dim current_active_workbook As Workbook
    On Error Resume Next
    Set current_active_workbook = ActiveWorkbook
    Err.Clear
    On Error GoTo ON_ERROR

    'ブックを開く
    Dim workbook_obj As Workbook
    If FilePath = "" Then
        Set workbook_obj = Workbooks.Add()
    Else
        Dim file_path As String
        file_path = pGetAbsolutePath(FilePath)
        Set workbook_obj = Workbooks.Add(Template:=file_path)
    End If
    ActiveWindow.Visible = False

    ' ブックを開いたことによって変わったアプリケーション設定を元に戻す
    If Not current_active_workbook Is Nothing Then Call current_active_workbook.Activate
    Call app_state.Restore

    OpenWorkbook = workbook_obj.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If Not current_active_workbook Is Nothing Then Call current_active_workbook.Activate
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_OpenWorkbook(Optional ByVal FilePath As String = "") As String
    IWorkbookService_OpenWorkbook = OpenWorkbook(FilePath:=FilePath)
End Function

Private Function pGetAbsolutePath(ByVal PathLikeString As String) As String
    pGetAbsolutePath = GetAbsolutePathFromParent(GetThisWorkbookDirectoryPath(), PathLikeString)
End Function

' ワークブックを保存します。
'
' @param Book 対象のブック名。
' @param FilePath 保存先ファイル パス。
' @param Force [省略可] 強制的に保存するか否か。
' @return 保存後のブック名。
'
' @details
' ワークブックを保存します。Force が False のときに、すでにファイルが存在している場合、確認メッセージが表示されます。
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    Dim target_book As Workbook
    Set target_book = Workbooks(Book)

    Dim file_path As String
    file_path = pGetAbsolutePath(FilePath)

    ' アクティブ ブックと保存対象ウィンドウの状態を控える
    Dim current_active_workbook As Workbook
    On Error Resume Next
    Set current_active_workbook = ActiveWorkbook
    Err.Clear
    On Error GoTo ON_ERROR

    Dim target_window_visible As Boolean
    Dim has_target_window_visible As Boolean
    target_window_visible = target_book.Windows(1).Visible
    has_target_window_visible = True

    Application.Calculation = xlCalculationAutomatic
    If Force Then Application.DisplayAlerts = False
    If Not target_window_visible Then target_book.Windows(1).Visible = True
    Call target_book.Activate
    Call target_book.SaveAs(file_path, FileFormat:=GetExcelFileFormat(file_path))
    target_book.Windows(1).Visible = target_window_visible

    ' 保存によって変わったアクティブ ブックとアプリケーション設定を元に戻す
    If Not current_active_workbook Is Nothing Then
        If current_active_workbook.Name <> target_book.Name Or target_window_visible Then
            Call current_active_workbook.Activate
        End If
    End If
    Call app_state.Restore

    ' 保存後のブック名を返す。
    SaveWorkbook = target_book.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If has_target_window_visible Then target_book.Windows(1).Visible = target_window_visible
    If Not current_active_workbook Is Nothing Then
        If target_book Is Nothing Then
            Call current_active_workbook.Activate
        ElseIf current_active_workbook.Name <> target_book.Name Or target_window_visible Then
            Call current_active_workbook.Activate
        End If
    End If
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    IWorkbookService_SaveWorkbook = SaveWorkbook(Book, FilePath, Force:=Force)
End Function

' ワークブックが最終保存後に変更されていないか確認します。
'
' @param Book 対象のブック名。
' @return 最終保存後に変更されていないか否か。
'
' @details
' Excel の Workbook.Saved と同じ意味で、保存先パスの有無は判定しません。
Public Function IsSaved(ByVal Book As String) As Boolean
    IsSaved = Workbooks(Book).Saved
End Function
Private Function IWorkbookService_IsSaved(ByVal Book As String) As Boolean
    IWorkbookService_IsSaved = IsSaved(Book)
End Function

' ワークブックが保存先パスを持つか確認します。
'
' @param Book 対象のブック名。
' @return 保存先パスを持つか否か。
'
' @details
' 新規未保存ブックのように保存先パスを持たない場合は False を返します。変更有無は判定しません。
Public Function HasPath(ByVal Book As String) As Boolean
    HasPath = (Workbooks(Book).Path <> "")
End Function
Private Function IWorkbookService_HasPath(ByVal Book As String) As Boolean
    IWorkbookService_HasPath = HasPath(Book)
End Function

' ワークブックを閉じます。
'
' @param Book 対象のブック名。
' @param Force [省略可] 強制的に閉じるか否か。True の場合は確認なしで変更を破棄して閉じます。
'
' @details
' ワークブックを閉じます。Force が False のときには、確認メッセージが表示されます。
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    If Force Then
        Call Workbooks(Book).Close(SaveChanges:=False)
    Else
        Call Workbooks(Book).Close
    End If
End Sub
Private Sub IWorkbookService_CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call CloseWorkbook(Book, Force:=Force)
End Sub


' ワークブック全体を検索します
'
' @param What 検索する文字列。
' @param RangeBounds 検索対象の範囲。
' @param LookInValues [省略可] 値を検索対象とするか。デフォルトは True です。False にした場合は数式を検索対象とします。
' @param LookAtWhole [省略可] セル内容が完全に同一であるものを検索するか。デフォルトは True です。True にした場合は完全一致で検索します。False としたときは部分一致で検索します。
' @param MatchCase [省略可] 大文字と小文字を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
' @param MatchByte [省略可] 全角と半角を区別するか。デフォルトは True です。True にした場合は区別して検索します。False としたときは区別せずに検索します。
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    Dim sheet_srv As WorksheetService
    Set sheet_srv = New WorksheetService
    
    Dim book_name As String
    If Book = "" Then
        book_name = ThisWorkbook.Name
    Else
        book_name = Book
    End If
    
    Dim sheet_arr() As String
    sheet_arr = GetAllWorksheet(book_name)
    
    Dim result_arr() As WorksheetRangeBounds
    If (Not (Not sheet_arr)) <> 0 Then
        Dim sheet_item As Variant 'String
        For Each sheet_item In sheet_arr
            Dim range_item As WorksheetRangeBounds
            Set range_item = New_RangeBounds(Sheet:=sheet_item, Book:=book_name)
            Dim tmp_arr() As WorksheetRangeBounds
            tmp_arr = sheet_srv.Find(What, range_item, LookInValues:=LookInValues, LookAtWhole:=LookAtWhole, MatchCase:=MatchCase, MatchByte:=MatchByte)
            Call pJoinRangeBoundsArray(result_arr, tmp_arr)
        Next sheet_item
    End If
    
    Find = result_arr
End Function
Private Function IWorkbookService_Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    IWorkbookService_Find = Find( _
            What, Book:=Book, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

Private Sub pJoinRangeBoundsArray(ByRef ResultArray() As WorksheetRangeBounds, ByRef TmpArray() As WorksheetRangeBounds)
    Dim item_idx As Long
    Dim prev_ubound As Long
    If (Not (Not ResultArray)) = 0 Then
        If (Not (Not TmpArray)) = 0 Then
            ' 両方未初期化
            Exit Sub
        Else
            ' TmpArray のみ初期化済み
            ResultArray = TmpArray
        End If
    Else
        If (Not (Not TmpArray)) = 0 Then
            ' ResultArray のみ初期化済み
            Exit Sub
        Else
            ' 両方初期化済み
            prev_ubound = UBound(ResultArray)
            ReDim Preserve ResultArray(0 To (prev_ubound + UBound(TmpArray) + 1))
            For item_idx = 0 To UBound(TmpArray)
                Set ResultArray(prev_ubound + 1 + item_idx) = TmpArray(item_idx)
            Next item_idx
        End If
    End If
End Sub


' ワークシートを追加します。
'
' @param Sheet [省略可] 追加するワークシート名。省略したときは自動で命名されます。
' @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
' @param SheetIndex [省略可] ワークシート番号。省略したときはアクティブなシートが指定されたものとします。
' @param Before [省略可] 前に追加するか。省略したときは指定されたワークシート番号の後ろに追加されます。
' @return 追加されたワークシートの名前。
'
' @details
' ワークシートを追加します。
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    On Error GoTo ON_ERROR

    ' 処理対象となるワークブックの取得
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' アクティブを控える
    Dim current_active_book As Workbook
    Dim current_target_active_sheet As Worksheet
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_target_active_sheet = target_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' ワークシート名のチェック
    Call pRaiseIfInvalidWorksheetName(Sheet, "Sheet")

    Dim added_sheet As Worksheet
    On Error Resume Next
    Set added_sheet = target_book.Worksheets(Sheet)
    Err.Clear
    On Error GoTo ON_ERROR
    If Not added_sheet Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService", "追加しようとしたシート名が既に存在します。(" & Sheet & "@" & Book & ")"

    ' 追加の目標となるワークシートの番号を取得
    Dim target_sheet_idx As Long
    target_sheet_idx = pGetTargetSheetIndex(target_book, SheetIndex)

    ' 追加の目標となるワークシートの取得
    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(target_sheet_idx)

    ' ワークシートの追加
    If Before Then
        Set added_sheet = target_book.Worksheets.Add(Before:=target_sheet)
    Else
        Set added_sheet = target_book.Worksheets.Add(After:=target_sheet)
    End If

    ' ワークシート名が指定されていたらワークシート名を変更
    If Sheet <> "" Then added_sheet.Name = Sheet

    Call pRestoreActiveState(current_active_book, current_target_active_sheet)

    AddWorksheet = added_sheet.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    Call pRestoreActiveState(current_active_book, current_target_active_sheet)

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    IWorkbookService_AddWorksheet = AddWorksheet( _
            Sheet:=Sheet, Book:=Book, _
            SheetIndex:=SheetIndex, Before:=Before)
End Function

Private Function pGetTargetSheetIndex(ByVal TargetBook As Workbook, ByVal SheetIndex As Long) As Long
    If SheetIndex = 0 Then
        ' シート番号が 0 なら、アクティブなシート
        pGetTargetSheetIndex = TargetBook.ActiveSheet.Index
    Else
        If SheetIndex < 0 Then
            ' シート番号が 0 未満なら、末尾から。-1 は末尾、-2 は末尾から 2 番目です。
            pGetTargetSheetIndex = MaxLng(TargetBook.Worksheets.Count + 1 + SheetIndex, 1)
        Else
            ' シート番号を超えるなら、先頭から
            pGetTargetSheetIndex = MinLng(SheetIndex, TargetBook.Worksheets.Count)
        End If
    End If
End Function

Private Sub pRaiseIfInvalidWorksheetName(ByVal WorksheetName As String, ByVal ArgumentName As String)
    If WorksheetName = "" Then Exit Sub

    If Len(WorksheetName) > 31 Then
        Err.Raise vbObjectError + 1, "Class WorkbookService", ArgumentName & " は 31 文字以下のシート名を指定してください。(" & WorksheetName & ")"
    End If

    Dim invalid_chars As Variant
    invalid_chars = Array(":", "\", "/", "?", "*", "[", "]")

    Dim invalid_char As Variant
    For Each invalid_char In invalid_chars
        If InStr(1, WorksheetName, CStr(invalid_char), vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 1, "Class WorkbookService", ArgumentName & " にシート名として使用できない文字が含まれています。(" & CStr(invalid_char) & ")"
        End If
    Next invalid_char
End Sub

Private Sub pRestoreActiveState(ByVal ActiveBook As Workbook, ByVal TargetBookActiveSheet As Worksheet)
    On Error Resume Next
    If Not TargetBookActiveSheet Is Nothing Then Call TargetBookActiveSheet.Activate
    If Not ActiveBook Is Nothing Then Call ActiveBook.Activate
    Err.Clear
    On Error GoTo 0
End Sub

' ワークシートを削除します。
'
' @param Sheet [省略可] 削除するワークシート名。省略したときはアクティブなシートが指定されたものとします。
' @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'
' @details
' ワークシートを削除します。
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Dim app_state As ApplicationScreenUpdateManager
    Set app_state = New ApplicationScreenUpdateManager

    On Error GoTo ON_ERROR

    ' 削除対象ワークシートがあるワークブックの取得
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' アクティブを控える
    Dim current_active_book As Workbook
    Dim current_target_active_sheet As Worksheet
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_target_active_sheet = target_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' 削除対象ワークシートの取得
    Dim target_sheet As Worksheet
    If Sheet = "" Then
        Set target_sheet = target_book.ActiveSheet
    Else
        Set target_sheet = target_book.Worksheets(Sheet)
    End If

    ' 削除対象ワークシートがあるワークブックに、削除対象以外のワークシートがない場合、空のワークシートを追加
    If target_book.Worksheets.Count <= 1 Then target_book.Worksheets.Add

    ' 対象ワークシートを削除
    Application.DisplayAlerts = False
    Call target_sheet.Delete
    Call pRestoreActiveState(current_active_book, current_target_active_sheet)
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
    Call pRestoreActiveState(current_active_book, current_target_active_sheet)
    Call app_state.Restore
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Sub
Private Sub IWorkbookService_RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call RemoveWorksheet(Sheet:=Sheet, Book:=Book)
End Sub

' ワークシートを複製します。
'
' @param SourceWorksheetName [省略可] 複製するワークシート名。省略したときはアクティブなシートが指定されたものとします。
' @param SourceWorkbookName [省略可] 複製するワークシートがあるワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
' @param DestinationWorksheetName [省略可] 複製されたワークシートにつける名前。省略したときは自動で命名されます。
' @param DestinationWorkbookName [省略可] 複製先のワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
' @param SheetIndex [省略可] ワークシート番号。省略したときは複製先のワークブックのアクティブなシートが指定されたものとします。
' @param Before [省略可] 複製したシートを対象の前に追加するか。省略したときは指定されたワークシート番号の後ろに追加されます。
' @return 追加されたワークシートの名前。
'
' @details
' ワークシートを複製します。
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String

    On Error GoTo ON_ERROR

    ' 複製元ワークシートがあるワークブックの取得
    Dim src_book As Workbook
    If SourceWorkbookName = "" Then
        Set src_book = ThisWorkbook
    Else
        Set src_book = Application.Workbooks(SourceWorkbookName)
    End If

    ' 複製元となるワークシートの取得
    Dim src_sheet As Worksheet
    If SourceWorksheetName = "" Then
        Set src_sheet = src_book.ActiveSheet
    Else
        Set src_sheet = src_book.Worksheets(SourceWorksheetName)
    End If

    ' 複製先ワークブックの取得
    Dim dst_book As Workbook
    If DestinationWorkbookName = "" Then
        Set dst_book = ThisWorkbook
    Else
        Set dst_book = Application.Workbooks(DestinationWorkbookName)
    End If

    ' ワークシート名のチェック
    Call pRaiseIfInvalidWorksheetName(DestinationWorksheetName, "DestinationWorksheetName")

    Dim added_sheet As Worksheet
    On Error Resume Next
    Set added_sheet = dst_book.Worksheets(DestinationWorksheetName)
    Err.Clear
    On Error GoTo ON_ERROR
    If Not added_sheet Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService", "複製先のシート名が既に存在します。(" & DestinationWorksheetName & "@" & DestinationWorkbookName & ")"

    ' アクティブを控える
    Dim current_active_book As Workbook
    Dim current_destination_active_sheet As Worksheet
    On Error Resume Next
    Set current_active_book = ActiveWorkbook
    Set current_destination_active_sheet = dst_book.ActiveSheet
    Err.Clear
    On Error GoTo ON_ERROR

    ' 追加の目標となるワークシートの番号を取得
    Dim target_sheet_idx As Long
    target_sheet_idx = pGetTargetSheetIndex(dst_book, SheetIndex)

    ' 追加の目標となるワークシートの取得
    Dim target_sheet As Worksheet
    Set target_sheet = dst_book.Worksheets(target_sheet_idx)

    ' ワークシートの追加
    If Before Then
        Call src_sheet.Copy(Before:=target_sheet)
    Else
        Call src_sheet.Copy(After:=target_sheet)
    End If

    ' 複製されたシートの取得
    Set added_sheet = dst_book.ActiveSheet

    ' ワークシート名が指定されていたらワークシート名を変更
    If DestinationWorksheetName <> "" Then added_sheet.Name = DestinationWorksheetName

    ' アクティブを戻す
    If Not current_destination_active_sheet Is Nothing Then Call current_destination_active_sheet.Activate
    If Not current_active_book Is Nothing Then Call current_active_book.Activate

    CopyWorksheet = added_sheet.Name
    On Error GoTo 0
    Exit Function

ON_ERROR:
    Dim err_num As Long: err_num = Err.Number
    Dim err_source As String: err_source = Err.Source
    Dim err_desc As String: err_desc = Err.Description
    Dim err_help_file As String: err_help_file = Err.HelpFile
    Dim err_help_context As Long: err_help_context = Err.HelpContext

    On Error Resume Next
    If Not current_destination_active_sheet Is Nothing Then Call current_destination_active_sheet.Activate
    If Not current_active_book Is Nothing Then Call current_active_book.Activate
    On Error GoTo 0

    Err.Raise err_num, err_source, err_desc, err_help_file, err_help_context
End Function
Private Function IWorkbookService_CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    IWorkbookService_CopyWorksheet = CopyWorksheet( _
            SourceWorksheetName:=SourceWorksheetName, _
            SourceWorkbookName:=SourceWorkbookName, _
            DestinationWorksheetName:=DestinationWorksheetName, _
            DestinationWorkbookName:=DestinationWorkbookName, _
            SheetIndex:=SheetIndex, _
            Before:=Before)
End Function

'* VBA コンポーネントを削除します。
'*
'* @param ComponentNames 削除対象のコンポーネント名、またはその配列。
'* @param ComponentType [省略可] 削除対象の VBIDE コンポーネント種別。0 の場合は種類を限定しません。
'* @param Book [省略可] ワークブック名。省略したときは ThisWorkbook のワークブック名が使用されます。
'*
'* @details
'* 対象ブックの VBProject から、指定された名前の VBA コンポーネントを削除します。
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If
    
    Dim vb_comps As Object
    Set vb_comps = target_book.VBProject.VBComponents
    
    Dim component_name_set As Object
    Set component_name_set = pBuildComponentNameSet(ComponentNames)
    
    Call pValidateVBComponentsForRemove(vb_comps, component_name_set, ComponentType)
    
    Dim component_name As Variant
    For Each component_name In component_name_set.Keys
        Call vb_comps.Remove(vb_comps.Item(CStr(component_name)))
    Next component_name
End Sub
Private Sub IWorkbookService_RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call RemoveVBComponents(ComponentNames, ComponentType:=ComponentType, Book:=Book)
End Sub

Private Function pBuildComponentNameSet(ByVal ComponentNames As Variant) As Object
    Dim component_name_set As Object
    Set component_name_set = CreateObject("Scripting.Dictionary")
    component_name_set.CompareMode = vbBinaryCompare
    
    If IsArray(ComponentNames) Then
        If Not pIsArrayInitialized(ComponentNames) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames には 1 件以上のコンポーネント名を指定してください。"
        
        Dim name_item As Variant
        For Each name_item In ComponentNames
            Call pAddComponentName(component_name_set, name_item)
        Next name_item
    Else
        Call pAddComponentName(component_name_set, ComponentNames)
    End If
    
    If component_name_set.Count = 0 Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames には 1 件以上のコンポーネント名を指定してください。"
    
    Set pBuildComponentNameSet = component_name_set
End Function

Private Function pIsArrayInitialized(ByVal TargetArray As Variant) As Boolean
    On Error Resume Next
    Dim lower_bound As Long
    lower_bound = LBound(TargetArray)
    Dim upper_bound As Long
    upper_bound = UBound(TargetArray)
    pIsArrayInitialized = (Err.Number = 0 And lower_bound <= upper_bound)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub pAddComponentName(ByVal ComponentNameSet As Object, ByVal ComponentNameValue As Variant)
    If IsArray(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames には文字列のコンポーネント名を指定してください。"
    If IsObject(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames には文字列のコンポーネント名を指定してください。"
    If IsError(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames にエラー値は指定できません。"
    If IsNull(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames に Null は指定できません。"
    If IsEmpty(ComponentNameValue) Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames に Empty は指定できません。"
    If VarType(ComponentNameValue) <> vbString Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames には文字列のコンポーネント名を指定してください。"
    
    Dim component_name As String
    component_name = CStr(ComponentNameValue)
    If component_name = "" Then Err.Raise vbObjectError + 1, "Class WorkbookService", "ComponentNames に空文字列は指定できません。"
    
    If Not ComponentNameSet.Exists(component_name) Then Call ComponentNameSet.Add(component_name, component_name)
End Sub

Private Sub pValidateVBComponentsForRemove(ByVal VBComponents As Object, ByVal ComponentNameSet As Object, ByVal ComponentType As Long)
    Dim component_name As Variant
    For Each component_name In ComponentNameSet.Keys
        Dim vb_comp As Object
        Set vb_comp = pFindVBComponent(VBComponents, CStr(component_name))
        
        If vb_comp Is Nothing Then Err.Raise vbObjectError + 1, "Class WorkbookService", "削除対象の VBA コンポーネントが存在しません。(" & CStr(component_name) & ")"
        If ComponentType <> 0 And vb_comp.Type <> ComponentType Then Err.Raise vbObjectError + 1, "Class WorkbookService", "削除対象の VBA コンポーネント種別が一致しません。(" & CStr(component_name) & ")"
        If vb_comp.Type = C_VBEXT_CT_DOCUMENT Then Err.Raise vbObjectError + 1, "Class WorkbookService", "Document モジュールは削除できません。(" & CStr(component_name) & ")"
    Next component_name
End Sub

Private Function pFindVBComponent(ByVal VBComponents As Object, ByVal ComponentName As String) As Object
    On Error Resume Next
    Set pFindVBComponent = VBComponents.Item(ComponentName)
    If Err.Number <> 0 Then
        Set pFindVBComponent = Nothing
        Err.Clear
    End If
    On Error GoTo 0
End Function
'
' ワークシートをアクティブにします。
'
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    ' 対象ワークシートがあるワークブックの取得
    Dim target_book As Workbook
    If Book = "" Then
        Set target_book = ThisWorkbook
    Else
        Set target_book = Application.Workbooks(Book)
    End If

    ' 対象ワークシートの取得
    Dim target_sheet As Worksheet
    Set target_sheet = target_book.Worksheets(Sheet)

    Call target_book.Activate
    Call target_sheet.Activate
End Sub
Private Sub IWorkbookService_ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call ActivateWorksheet(Sheet, Book:=Book)
End Sub

