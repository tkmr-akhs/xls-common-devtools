Attribute VB_Name = "Lib_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 基本的な関数などをまとめた標準モジュールです。
'! 他のツールとも共用されるため、このツールで使用しないものも含まれます。
'!
' #############################################################################

'* WorkgbookService。ユニットテキスト時にはテスト ダブルに置き換えてください。
Public WbSrv As IWorkbookService

'* WorkgsheetService。ユニットテキスト時にはテスト ダブルに置き換えてください。
Public WsSrv As IWorksheetService


'* 未設定の日付を表す
Public Const G_DATE_NULL As Date = #12/31/1899#

'* 有効な日付の最小値
Public Const G_DATE_MINIMUM As Date = #3/1/1900#

'* 行番号の最大値
Public Const G_ROW_MAX As Long = 1048576

'* 列番号の最大値
Public Const G_COL_MAX As Long = 16384

'* 行番号および列番号の省略値
Public Const G_OMIT_CELL_INDEX As Long = -2147483648#

'* ファイルシステムのパス区切り文字
Public Const G_FS_PATH_SEP As String = "\"

Private Const C_LONG_MAX As Long = 2147483647#
Private Const C_LONG_MIN As Long = -2147483648#
Private Const C_A1_TOKEN_CELL As Long = 1
Private Const C_A1_TOKEN_ROW As Long = 2
Private Const C_A1_TOKEN_COLUMN As Long = 3

' #############################################################################
'
' 独自ワークシート関数
'
' #############################################################################

''* 新しい Excel で存在するワークシート関数 TEXTSPLIT の独自実装。
''*
''* @param Expression 分割するテキスト。
''* @param ColumnDelimiter 列の区切り文字。(一般的な CSV では 「,」)
''* @param RowDelimiter 行の区切り文字。(一般的な CSV では 改行記号) (省略可能)
''* @param IgnoreEmpty 連続する区切り記号を無視かどうか。無視するには True を指定します。 既定値は False で、PadWith で埋められます。
''* @param MatchMode (未実装の機能)
''* @param PadWith 連続する区切り記号の場合の値。既定値は #N/A。
''* @return 分割した結果の配列。
''*
''* @details
''* TEXTSPLIT がないバージョンの Excel のための実装です。詳細については Microsoft の TEXTSPLIT の説明を参照してください。
''*
''* @see https://support.microsoft.com/ja-jp/office/textsplit-%E9%96%A2%E6%95%B0-b1ca414e-4c21-4ca0-b1b7-bdecace8a6e7
'Function TEXTSPLIT( _
'        ByVal Expression As String, ByVal ColumnDelimiter As String, Optional ByVal RowDelimiter As String, _
'        Optional ByVal IgnoreEmpty As Boolean = False, Optional ByVal MatchMode As Integer = 0, Optional ByVal PadWith As Variant = Nothing) As Variant()
'
'    If MatchMode <> 0 Then
'        ' MatchMode 1 (IgnoreCase) は未実装
'        TEXTSPLIT = CVErr(xlErrValue)
'        Exit Function
'    End If
'
'    If IsObject(PadWith) Then
'        If PadWith Is Nothing Then
'            PadWith = CVErr(xlErrNA)
'        End If
'    End If
'
'    Dim rows_arr() As String
'    rows_arr = Split(Expression, RowDelimiter)
'
'    Dim max_col As Long
'    Dim rows_list As ObjectList
'    Set rows_list = New ObjectList
'    Dim row_item As Variant 'String
'    For Each row_item In rows_arr
'        If row_item <> "" Or Not IgnoreEmpty Then
'            Dim cols_arr() As String
'            cols_arr = Split(row_item, ColumnDelimiter)
'
'            Dim cols_list As ObjectList
'            Set cols_list = New ObjectList
'
'            Dim col_item As Variant 'String
'            For Each col_item In cols_arr
'                If col_item <> "" Or Not IgnoreEmpty Then
'                    Call cols_list.Add(col_item)
'                End If
'            Next col_item
'
'            If 0 < cols_list.Count Or Not IgnoreEmpty Then
'                Call rows_list.Add(cols_list)
'                max_col = MaxLng(max_col, cols_list.Count)
'            End If
'        End If
'    Next row_item
'
'    Dim result() As Variant
'    ReDim result(0 To rows_list.Count - 1, 0 To max_col - 1)
'
'    Dim row_idx As Long
'    Dim col_idx As Long
'    For row_idx = 0 To UBound(result, 1) Step 1
'        For col_idx = 0 To UBound(result, 2) Step 1
'            Set cols_list = rows_list.Item(row_idx)
'            If col_idx < cols_list.Count Then
'                result(row_idx, col_idx) = cols_list.Item(col_idx)
'            Else
'                result(row_idx, col_idx) = PadWith
'            End If
'        Next col_idx
'    Next row_idx
'
'    TEXTSPLIT = result
'End Function

''* 新しい Excel で存在するワークシート関数 TEXTJOIN の独自実装。
''*
''* @param Delimiter 区切り文字。
''* @param IgnoreEmpty 空文字列を無視かどうか。True の場合、空文字列は無視されます。False の場合、無視されず区切り文字が連続します。
''* @param Expression1 1 番目の文字列
''* @param Expressions 2 番目以降の文字列。(可変長引数)
''*
''* @return 結合した文字列。
''*
''* @details
''* TEXTJOIN がないバージョンの Excel のための実装です。詳細については Microsoft の TEXTJOIN の説明を参照してください。
''*
''* @see https://support.microsoft.com/ja-jp/office/textjoin-%E9%96%A2%E6%95%B0-357b449a-ec91-49d0-80c3-0e8fc845691c
'Function TEXTJOIN(ByVal Delimiter As String, ByVal IgnoreEmpty As Boolean, ByVal Expression1 As Variant, ParamArray Expressions() As Variant) As Variant
'    Dim result As String
'    Dim is_first As Boolean
'
'    is_first = True
'
'    On Error Resume Next
'    Call pTextJoinCore(result, is_first, Delimiter, Expression1, IgnoreEmpty)
'    If Err.Number <> 0 Then
'        TEXTJOIN = CVErr(xlErrNA)
'        Exit Function
'    End If
'
'    Dim param_item As Variant
'    For Each param_item In Expressions
'        Call pTextJoinCore(result, is_first, Delimiter, param_item, IgnoreEmpty)
'        If Err.Number <> 0 Then
'            TEXTJOIN = CVErr(xlErrNA)
'            Exit Function
'        End If
'    Next param_item
'
'    TEXTJOIN = result
'End Function

'Sub pTextJoinCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If TypeOf Expression Is Range Or IsArray(Expression) Then
'        ' Expression が Range か配列の場合
'        Dim cell_item As Range
'        For Each cell_item In Expression
'            Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, cell_item.Value, IgnoreEmpty)
'        Next cell_item
''    ElseIf IsArray(Expression) Then
''        ' Expression が配列の場合
''        If LBound(Expression) <= UBound(Expression) Then
''            Dim idx As Long
''            For idx = LBound(Expression) To UBound(Expression)
''                Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression(idx), IgnoreEmpty)
''            Next idx
''        ElseIf Not IgnoreEmpty Then
''            ResultString = ResultString & Delimiter
''        End If
'    Else
'        ' Range や配列でない場合
'        Call pTextJoinItemCore(ResultString, IsFirst, Delimiter, Expression, IgnoreEmpty)
'    End If
'End Sub

'Private Sub pTextJoinItemCore(ByRef ResultString As String, ByRef IsFirst As Boolean, Delimiter As String, Expression As Variant, IgnoreEmpty As Boolean)
'    If IsError(Expression) Then Err.Raise xlErrNA, "TEXTJOIN", "結合対象にエラー値が含まれます"
'    If Expression <> "" Or Not IgnoreEmpty Then
'        If IsFirst Then
'            ResultString = Expression
'            IsFirst = False
'        Else
'            ResultString = ResultString & Delimiter & Expression
'        End If
'    End If
'End Sub

'* 文字列同士を比較し、差分の文字列を取得します。
'*
'* @param Expression1 元の文字列。
'* @param Expression2 変更後の文字列。
'* @param ExtractType 抽出タイプ。0 を指定すると、元の文字列から削除された文字列を取得します。1 を指定すると、元の文字列に追加された文字列を取得します。
'*
'* 文字列同士を比較し、差分の文字を前から順に結合した文字列を取得します。
'* たとえば、Expression1 が「aaa [bbb ccc] ddd」、Expression2 が「aaa bbb ddd eee」で、ExtractType が 0 の場合は、「[ccc] 」が返ります。
Public Function DIFFSTR( _
    ByVal Expression1 As String, _
    ByVal Expression2 As String, _
    Optional ByVal ExtractType As Integer = 0) As Variant  ' String

    '=== 1) 空文字列が混在するケースの簡易処理 ==================
    If Expression1 = "" Or Expression2 = "" Then
        If Expression1 = "" Then
            If ExtractType = 0 Then
                DIFFSTR = ""
            Else
                DIFFSTR = Expression2
            End If
        Else
            ' Expression2 が空文字列
            If ExtractType = 0 Then
                DIFFSTR = Expression1
            Else
                DIFFSTR = ""
            End If
        End If
        Exit Function
    End If
    
    '=== 2) 文字列を文字配列に変換 ==============================
    Dim expr_1_arr() As String
    Dim expr_2_arr() As String
    expr_1_arr = ConvertStringToCharArray(Expression1)
    expr_2_arr = ConvertStringToCharArray(Expression2)
    
    '=== 3) DiffStringArray を使用して差分を算出 ===============
    Dim diff_type_arr() As String
    Call DiffStringArray(expr_1_arr, expr_2_arr, diff_type_arr, EnableReplaceType:=False)
    
    '=== 4) 抽出タイプに応じて、差分文字列を再構築 ==============
    Dim result_str As String
    result_str = ""
    
    Dim i_idx As Long
    Dim last_idx As Long
    last_idx = UBound(expr_1_arr)  ' = UBound(expr_2_arr) = UBound(diff_type_arr)
    
    For i_idx = LBound(expr_1_arr) To last_idx
        Dim diff_tag As String
        diff_tag = diff_type_arr(i_idx)
        
        Select Case ExtractType
            Case 0
                ' 削除で旧文字列側が取り除かれた文字を集める
                If (diff_tag = "DEL") Then
                    result_str = result_str & expr_1_arr(i_idx)
                End If
            Case 1
                ' 追加で新文字列側に増えた文字を集める
                If (diff_tag = "ADD") Then
                    result_str = result_str & expr_2_arr(i_idx)
                End If
        End Select
    Next i_idx
    
    DIFFSTR = result_str
    
End Function

' #############################################################################
'
' 共通クラス モジュールのコンストラクターなど
'
' #############################################################################

'* 共通的なサービスを初期化します。
'*
'* @details
'* 共通的なサービスを初期化します。初期化対象は以下の通りです。
'*
'* * WorkbookService
'* * WorksheetService
Public Sub InitializeCommonService()
    If WbSrv Is Nothing Then Set WbSrv = New WorkbookService
    If WsSrv Is Nothing Then Set WsSrv = New WorksheetService
End Sub

'* WorksheetRangeBounds インスタンスを新規作成します。
'*
'* @param RowIndex [省略可] 行番号。
'* @param ColumnIndex [省略可] 列番号。
'* @param FinishRowIndex [省略可] 最後の行番号。
'* @param FinishColumnIndex [省略可] 最後の列番号。
'* @param SheetName [省略可] ワークシート名。
'* @param BookName [省略可] ワークブック名。
'* @return 新しい WorksheetRangeBounds インスタンス
'*
'* @details
'* WorksheetRangeBounds インスタンスを新規作成します。
'*
'* 使用例:
'* @code
'* ' 1 セル (B5)
'* Dim cell_bounds As WorksheetRangeBounds
'* Set cell_bounds = New_RangeBounds(Row:=5, Column:=2)
'*
'* ' 1 行全体 (5:5)
'* Dim row_bounds As WorksheetRangeBounds
'* Set row_bounds = New_RangeBounds(Row:=5)
'*
'* ' 2 列全体 (B:C)
'* Dim two_columns_bounds As WorksheetRangeBounds
'* Set two_columns_bounds = New_RangeBounds(Column:=2, FinishColumn:=3)
'*
'* ' 表全体 (B2:E20)
'* Dim table_bounds As WorksheetRangeBounds
'* Set table_bounds = New_RangeBounds(Row:=2, Column:=2, FinishRow:=20, FinishColumn:=5)
'* @endcode
Public Function New_RangeBounds( _
        Optional ByVal Row As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Column As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal Sheet As String = "Sheet1", _
        Optional ByVal Book As String = "") As WorksheetRangeBounds
    
    Dim result As WorksheetRangeBounds
    Set result = New WorksheetRangeBounds
    
    Call result.Initialize( _
            Row:=Row, _
            Column:=Column, _
            FinishRow:=FinishRow, _
            FinishColumn:=FinishColumn, _
            Sheet:=Sheet, _
            Book:=Book)
    
    Set New_RangeBounds = result
End Function

'* Excel アドレス文字列から WorksheetRangeBounds インスタンスを新規作成します。
'*
'* @param AddressString 初期化に使用する Excel アドレス文字列。
'* @return 新しい WorksheetRangeBounds インスタンス
'*
'* @details
'* New_ 系の処理は薄いファクトリに留め、実処理は WorksheetRangeBounds.InitializeFromAddress に委譲します。
Public Function New_RangeBoundsFromAddress(ByVal AddressString As String) As WorksheetRangeBounds
    Dim result As WorksheetRangeBounds
    Set result = New WorksheetRangeBounds
    Call result.InitializeFromAddress(AddressString)

    Set New_RangeBoundsFromAddress = result
End Function

'* UserInputSheet インスタンスを新規作成します。
'*
'* @param SheetName [省略可] ワークシート名。
'* @param BookName [省略可] ワークブック名。
'* @return 新しい UserInputSheet インスタンス
'*
'* @details
'* UserInputSheet インスタンスを新規作成します。
Public Function New_InputSheet(ByVal Sheet As String, Optional ByVal Book As String = "") As UserInputSheet
    Dim result As UserInputSheet
    Set result = New UserInputSheet
    Call result.Initialize(Sheet, Book:=Book)
    Set New_InputSheet = result
End Function

' #############################################################################
'
' エラー ハンドリング
'
' #############################################################################

'* エラー情報を処理し、必要に応じて情報を補完または蓄積するためのサブルーチン。
'*
'* @param ErrNumber [出力] エラー番号を格納する変数。既定値は 0。Append を True にしても、ErrNumber は上書きされます。(省略可能)
'* @param ErrSource [出力] エラーの発生源を格納する文字列変数。(省略可能)
'* @param ErrDescription [出力] エラーの説明を格納する文字列変数。(省略可能)
'* @param Supplementation エラーの説明に付加する補足情報。既定値は空文字列。
'* @param ErrClear エラーをクリアするかどうか。True の場合、エラー情報はクリアされます。既定値は True。
'* @param Append 複数のエラー情報を結合するかどうか。True の場合、既存の情報に追加されます。既定値は False。
'*
'* @details
'* このサブルーチンは、VBA の標準的なエラー情報をカスタマイズした形で処理します。
'* 特に、エラー番号、発生源、説明を取得し、補足情報を付加する機能があります。
'* 複数のエラーが発生した場合、Append パラメータを使用して情報を連結できます。
'*
'* 使用例:
'* @code
'* Dim ErrNum As Long
'* Dim ErrSrc As String
'* Dim ErrDesc As String
'*
'* On Error Resume Next
'*
'* ''//故意にエラーを発生させる例 (ゼロ除算)
'* Debug.Print 1 / 0
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "補足情報 1", Append:=True)
'*
'* ''//故意にエラーを発生させる例 (未定義サブ プロシージャ呼び出し)
'* Call NonExistentProcedure
'*
'* Call HandleError(ErrNum, ErrSrc, ErrDesc, "補足情報 2", Append:=True)
'*
'* On Error GoTo 0
'*
'* Err.Raise ErrNum, ErrSrc, ErrDesc
'* @endcode
Public Sub HandleError( _
        Optional ByRef ErrNumber As Long = 0, _
        Optional ByRef ErrSource As String = "", _
        Optional ByRef ErrDescription As String = "", _
        Optional ByVal Supplementation As String = "", _
        Optional ByVal ErrClear As Boolean = True, _
        Optional ByVal Append As Boolean = False)
    
    If Err.Number <> 0 Then
        ErrNumber = Err.Number
        If Append Then
            ErrSource = ErrSource & Err.Source & vbCrLf
            If Supplementation <> "" Then
                ErrDescription = ErrDescription & Err.Description & " [ " & Supplementation & " ]" & vbCrLf
            Else
                ErrDescription = ErrDescription & Err.Description & vbCrLf
            End If
        Else
            ErrSource = Err.Source
            If Supplementation <> "" Then
                ErrDescription = Err.Description & " [ " & Supplementation & " ]"
            Else
                ErrDescription = Err.Description
            End If
        End If
        
        If ErrClear Then
            Err.Clear
        End If
    End If
End Sub

' #############################################################################
'
' GUI 関連
'
' #############################################################################

'* ワークシートへボタンを作成します。
'*
'* @param TargetSheet 対象のワークシート。
'* @param RowIndex ボタンを追加する位置のセルの行番号。
'* @param ColumnIndex ボタンを追加する位置のセルの列番号。
'* @param Caption ボタンの表示文字列。
'* @param OnAction クリック時に実行するサブ プロシージャ。
'* @param OnAction [省略可]ボタンのコンポーネント名 (Name プロパティ)。
'* @return 追加したボタン オブジェクト
'*
'* @details
'* ワークシートへボタンを作成します。
Public Function AddButton( _
    ByVal TargetSheet As Worksheet, _
    ByVal RowIndex As Long, _
    ByVal ColumnIndex As Long, _
    ByVal Caption As String, _
    ByVal OnAction As String, _
    Optional ByVal Name As String = "") As Shape

    Dim tgt_rng As Range
    Set tgt_rng = TargetSheet.Cells(RowIndex, ColumnIndex)
    
    Dim shp_btn As Shape
    Set shp_btn = TargetSheet.Shapes.AddShape( _
                 Type:=msoShapeRectangle, _
                 Left:=tgt_rng.Left, _
                 Top:=tgt_rng.Top, _
                 Width:=tgt_rng.Width, _
                 Height:=tgt_rng.Height)
    
    If Name <> "" Then
        shp_btn.Name = Name
    End If
    
    shp_btn.OnAction = OnAction
    
    With shp_btn.TextFrame2
        .TextRange.Text = Caption
        
        With .TextRange.Font
            .NameComplexScript = "Yu Gothic UI Semibold"
            .NameFarEast = "Yu Gothic UI Semibold"
            .Name = "Yu Gothic UI Semibold"
            .Size = 9
            .Fill.ForeColor.RGB = RGB(64, 64, 64)
        End With
        
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
    End With
    
    shp_btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp_btn.TextFrame.VerticalAlignment = xlVAlignCenter
    
    shp_btn.Fill.ForeColor.RGB = RGB(192, 192, 192)
    shp_btn.Fill.Solid
    shp_btn.Line.ForeColor.RGB = RGB(64, 64, 64)
    shp_btn.Line.Weight = 1
    
    shp_btn.Shadow.Visible = msoFalse
    shp_btn.ThreeD.Visible = False
    
    Set AddButton = shp_btn
End Function

'* ワークシートのボタンを削除します。
'*
'* @param TargetSheet 対象のワークシート。
'* @param Name ボタンのコンポーネント名 (Name プロパティ)。
'*
'* @details
'* ワークシート上のボタンを、コンポーネント名を指定して削除します。
Public Sub DeleteButton(ByVal TargetSheet As Worksheet, ByVal Name As String)

    Dim shp As Shape
    Dim is_deleted As Boolean: is_deleted = False
    
    For Each shp In TargetSheet.Shapes
        If shp.Name = Name Then
            shp.Delete
            is_deleted = True
            Exit For
        End If
    Next shp
    
    If Not is_deleted Then
        Err.Raise vbObjectError + 1, "Sub DeleteButton", "ボタンが見つかりませんでした。(" & Name & ")"
    End If
    
End Sub

'* ワークシートのすべてのボタンを削除します。
'*
'* @param TargetSheet 対象のワークシート。
'*
'* @details
'* ワークシート上のボタンを、すべて削除します。
Public Sub ClearButton(ByVal TargetSheet As Worksheet)
    Dim idx As Long
    ' 後ろから削除しないとループ中に変更が起きた際に不具合が出ることがあるので
    ' For i = Shapes.Count To 1 Step -1 の形が安全です
    For idx = TargetSheet.Shapes.Count To 1 Step -1
        
        If TargetSheet.Shapes(idx).OnAction <> "" Then
            ' OnActionが設定されている → 「図形ボタン」と判断して削除
            TargetSheet.Shapes(idx).Delete
        End If
        
    Next idx
    
End Sub

' #############################################################################
'
' クリップボード関連
'
' #############################################################################

'* クリップボードへ文字列を送信します。
'*
'* @param SourceText クリップボードにコピーする文字列。
'*
'* @details
'* VBA の `Forms.TextBox.1` オブジェクトを使用して、指定された文字列をクリップボードにコピーします。
Public Sub SetClipboard(ByVal SourceText As String)
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        .Text = SourceText
        .SelStart = 0
        .SelLength = .TextLength
        .Copy
    End With
End Sub

'* クリップボードから文字列を取得します。
'*
'* @return クリップボードに格納されている文字列。クリップボードに文字列が含まれていない場合、空文字列を返します。
'*
'* @details
'* VBA の `Forms.TextBox.1` オブジェクトを使用して、クリップボードの内容を取得します。
Public Function GetClipboard() As String
    With CreateObject("Forms.TextBox.1")
        .MultiLine = True
        If .CanPaste = True Then .Paste
        GetClipboard = .Text
    End With
End Function

'* 可能な限り書式を除いた貼り付けを実行します。
'*
'* @details
'* クリップボードの内容を、以下の順序で貼り付けを試みます:
'* 1. 数式として貼り付け (`Paste:=xlPasteFormulas`)。
'* 2. テキストのみを貼り付け (`NoHTMLFormatting:=True`)。
'* 3. いずれも失敗した場合、標準の貼り付けを実行。
Public Sub PasteFormulas()
    On Error Resume Next
    
    Call Selection.PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    
    Call ActiveSheet.PasteSpecial(Format:="HTML", Link:=False, DisplayAsIcon:=False, NoHTMLFormatting:=True)
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    
    Call ActiveSheet.Paste
End Sub

' #############################################################################
'
' ファイル操作関連
'
' #############################################################################


'* パス内の特殊文字を全角文字に置換します。
'*
'* @param Path 置換対象のパス文字列
'* @return 特殊文字を全角文字に置換したパス文字列
'*
'* @details
'* ファイルシステムで使用できない特殊文字を全角文字に置換します。
'* 以下の文字が置換されます: `\`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`。
Public Function ReplaceSpecialCharacterOnFileSystemPath(ByVal Path As String) As Variant
    Dim result As String
    
    result = Path
    
    result = Replace(result, "\", "＼")
    result = Replace(result, "/", "／")
    result = Replace(result, ":", "：")
    result = Replace(result, "*", "＊")
    result = Replace(result, "?", "？")
    result = Replace(result, """", "″")
    result = Replace(result, "<", "＜")
    result = Replace(result, ">", "＞")
    result = Replace(result, "|", "｜")
    
    ReplaceSpecialCharacterOnFileSystemPath = result
End Function

'* パスを結合します。
'*
'* @param Path1 最初のパス文字列
'* @param Path2 結合する2番目のパス文字列
'* @param Paths 可変長の追加パス文字列
'* @return 結合されたパス文字列
'*
'* @details
'* 指定された複数の文字列をパス区切り文字で結合します。
'* 追加のパス文字列を可変長引数として渡すことができます。
Public Function JoinPath(ByVal Path1 As String, ByVal Path2 As String, ParamArray Paths() As Variant) As String
    Dim result As String
    
    result = pJoinPathCore(Path1, Path2)
    
    If UBound(Paths) = -1 Then
        JoinPath = result
        Exit Function
    End If
    
    Dim idx As Long
    For idx = LBound(Paths) To UBound(Paths)
        result = pJoinPathCore(result, Paths(idx))
    Next idx
    
    JoinPath = result
End Function

Private Function pJoinPathCore(ByVal Path1 As String, ByVal Path2 As String) As String
    
    If EndsWith(Path1, G_FS_PATH_SEP) Then
        If StartsWith(Path2, G_FS_PATH_SEP) Then
            pJoinPathCore = Left(Path1, Len(Path1) - 1) & Path2
        Else
            pJoinPathCore = Path1 & Path2
        End If
    Else
        If StartsWith(Path2, G_FS_PATH_SEP) Then
            pJoinPathCore = Path1 & Path2
        Else
            pJoinPathCore = Path1 & G_FS_PATH_SEP & Path2
        End If
    End If
End Function

'* パスの最後の部分を除いたパスを取得します。
'*
'* @param Path 入力パス文字列
'* @param IgnoreEndSep [省略可] 入力パス文字列の末尾の区切り文字を無視するか否か
'* @return 最後の部分を除いたパス文字列
'*
'* @details
'* 入力されたパス文字列の最後のパス区切り文字より前の部分を返します。
'* 例えば、`Path\to\File` の場合は `Path\to` を返します。
Public Function GetParentPath(ByVal Path As String, Optional ByVal IgnoreEndSep As Boolean = False) As String
    If EndsWith(Path, G_FS_PATH_SEP) Then
        Path = Left(Path, Len(Path) - 1)
        If Not IgnoreEndSep Then
            GetParentPath = Path
            Exit Function
        End If
    End If
    
    GetParentPath = Left(Path, InStrRev(Path, G_FS_PATH_SEP) - 1)
End Function

'* パスの最後の部分を取得します。
'*
'* @param Path 入力パス文字列
'* @param BaseName [省略可] ベース名を含めるか。規定値は True
'* @param Extension [省略可] 拡張子を含めるか。規定値は True
'* @param IgnoreEndSep [省略可] 入力パス文字列の末尾の区切り文字を無視するか否か
'* @return 最後の部分 (ファイル名またはフォルダ名)
'*
'* @details
'* 入力されたパス文字列の最後のパス区切り文字より後の部分を返します。
'* オプション引数 `Extension` を False にすると、拡張子を除去した結果を返します。
'* 例えば、`Path\to\File.txt` の場合は `File` を返します。
Public Function GetLeafFromPath(ByVal Path As String, Optional ByVal BaseName As Boolean = True, Optional ByVal Extension As Boolean = True, Optional ByVal IgnoreEndSep As Boolean = False) As String
    If Not BaseName And Not Extension Then Exit Function
    
    If EndsWith(Path, G_FS_PATH_SEP) Then
        If Not IgnoreEndSep Then
            GetLeafFromPath = ""
            Exit Function
        Else
            Path = Left(Path, Len(Path) - 1)
        End If
    End If
    
    Dim last_sep As Integer
    last_sep = InStrRev(Path, G_FS_PATH_SEP)
    
    Dim leaf_str As String
    If last_sep > 0 Then
        leaf_str = Mid(Path, last_sep + 1)
    Else
        leaf_str = Path
    End If
    
    If BaseName And Extension Then
        GetLeafFromPath = leaf_str
        Exit Function
    End If
    
    Dim last_period As Integer
    last_period = InStrRev(leaf_str, ".")
    
    Dim base_name As String, file_ext As String
    If 1 < last_period Then
        base_name = Left(leaf_str, last_period - 1)
        file_ext = Right(leaf_str, Len(leaf_str) - last_period + 1)
    Else
        base_name = leaf_str
        file_ext = ""
    End If
    
    If BaseName And Not Extension Then
        GetLeafFromPath = base_name
    ElseIf Not BaseName And Extension Then
        GetLeafFromPath = file_ext
    Else
        GetLeafFromPath = leaf_str
    End If
End Function

'* あるパス文字列が絶対パスかを判定します。
'*
'* @param TestPath 判定対象のパス文字列
'* @return 絶対パスの場合は True、それ以外は False
'*
'* @details
'* 入力されたパス文字列が絶対パスかどうかを判定します。
Public Function IsAbsolutePath(ByVal TestPath As String) As Boolean
    If Mid(TestPath, 2, 2) = ":\" Or Left(TestPath, 2) = "\\" Then
        IsAbsolutePath = True
    Else
        IsAbsolutePath = False
    End If
End Function

' #############################################################################
'
' 日付操作関連
'
' #############################################################################

'* ある日付が、未設定かどうかを判定します。
'*
'* @param TestDate 判定対象の日付
'* @return 未設定の日付の場合は True、それ以外は False
'*
'* @details
'* 指定された日付が `G_DATE_MINIMUM` 未満、または `G_DATE_NULL` と等しい場合に未設定とみなして True を返します。
'* `G_DATE_NULL` は未設定の日付を表し、`G_DATE_MINIMUM` は有効な日付の最小値を表します。
Public Function IsNullDate(ByVal TestDate As Date) As Boolean
    If TestDate < G_DATE_MINIMUM Then
        IsNullDate = True
    Else
        IsNullDate = False
    End If
End Function

' #############################################################################
'
' 各種変換
'
' #############################################################################

'* 文字列を文字の配列に変換します。
'*
'* @param Expression 文字列。
'* @return 文字の配列。
'*
'* @details
'* 文字列を文字の配列に変換します。
Public Function ConvertStringToCharArray(ByVal Expression As String) As String()
    Dim result_ubound As Long
    result_ubound = Len(Expression) - 1
    
    If result_ubound < 0 Then Exit Function
    
    Dim result() As String
    ReDim result(0 To result_ubound)
    
    Dim char_idx As Long
    For char_idx = 0 To result_ubound
        result(char_idx) = Mid(Expression, char_idx + 1, 1)
    Next char_idx
    
    ConvertStringToCharArray = result
End Function

'* String() 型の配列を Variant() 型の配列に変換します。
'*
'* @param StringArray 変換対象の String 型の配列
'* @return Variant 型の配列
'*
'* @details
'* String 型の配列を Variant 型に変換します。各要素の内容はそのまま維持されます。
Public Function ConvertArrayStringToVariant(ByRef StringArray() As String) As Variant()
    Dim result() As Variant
    Dim idx As Long
    
    ReDim result(LBound(StringArray) To UBound(StringArray))
    
    For idx = LBound(StringArray) To UBound(StringArray)
        result(idx) = StringArray(idx)
    Next idx
    
    ConvertArrayStringToVariant = result
End Function

'* Variant() 型の配列を String() 型の配列に変換します。
'*
'* @param VariantArray 変換対象の Variant 型の配列
'* @return String 型の配列
'*
'* @details
'* Variant 型の配列を String 型に変換します。各要素の内容はそのまま維持されます。
Public Function ConvertArrayVariantToString(ByRef VariantArray() As Variant) As String()
    Dim result() As String
    Dim idx As Long
    
    ReDim result(LBound(VariantArray) To UBound(VariantArray))
    
    For idx = LBound(VariantArray) To UBound(VariantArray)
        result(idx) = VariantArray(idx)
    Next idx
    
    ConvertArrayVariantToString = result
End Function

'* Boolean 型の値を文字列に変換します。
'*
'* @param BooleanValue 変換対象の Boolean 値
'* @param FlagOnString True の場合に返す文字列 (既定値は "■")
'* @param FlagOffString False の場合に返す文字列 (既定値は空文字列)
'* @return 変換後の文字列
'*
'* @details
'* Boolean 値を指定された文字列に変換します。
'* True の場合は `FlagOnString`、False の場合は `FlagOffString` を返します。
Public Function ConvertBooleanToString( _
        ByVal BooleanValue As String, _
        Optional ByVal FlagOnString As String = "■", _
        Optional ByVal FlagOffString As String = "") As String
    
    If BooleanValue Then
        ConvertBooleanToString = FlagOnString
    Else
        ConvertBooleanToString = FlagOffString
    End If
End Function

'* 文字列を Boolean 型の値に変換します。
'*
'* @param FlagValue 変換対象の文字列
'* @param FlagOnString True とみなす文字列 (既定値は "■")
'* @param FlagOffString False とみなす文字列 (既定値は空文字列)
'* @return 変換後の Boolean 値
'*
'* @details
'* 指定された文字列を基に Boolean 値を判定します。
'* `FlagOnString` の場合は True、`FlagOffString` の場合は False を返します。
'* それ以外の値の場合はエラーを発生させます。
Public Function ConvertStringToBoolean( _
        ByVal FlagValue As String, _
        Optional ByVal FlagOnString As String = "■", _
        Optional ByVal FlagOffString As String = "") As Boolean
    
    If FlagValue = FlagOnString Then
        ConvertStringToBoolean = True
    ElseIf FlagValue = FlagOffString Then
        ConvertStringToBoolean = False
    Else
        Err.Raise Number:=vbObjectError + 1, Source:="Sub CreateBackupFile", Description:="許容される FlagValue は「" & FlagOnString & "」か「" & FlagOffString & "」です。(" & FlagValue & ")"
    End If
End Function

' #############################################################################
'
' 配列関連
'
' #############################################################################

'* 2 次元配列を 1 次元配列に変換します。
'*
'* @param OriginalArray 2 次元配列。
'* @param ColumnDirection [省略可] 列方向 (縦方向、1 次元目) に読み取っていくか。デフォルトは False で行列方向 (横方向、2 次元目) に読み取っていきます。
'* @return 1 次元配列。
'*
'* @details
'* 2 次元配列を 0 ベースの 1 次元配列に変換します。
Public Function ConvertArray2dTo1d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()
    
    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="引数が配列ではありません。(" & TypeName(OriginalArray) & ")"
    End If

    Dim row_low_bnd As Long
    Dim row_up_bnd As Long
    Dim col_low_bnd As Long
    Dim col_up_bnd As Long
    Dim total_cnt As Long
    Dim res_arr() As Variant
    Dim arr_idx As Long
    Dim row_index As Long
    Dim col_index As Long

    row_low_bnd = LBound(OriginalArray, 1)
    row_up_bnd = UBound(OriginalArray, 1)
    col_low_bnd = LBound(OriginalArray, 2)
    col_up_bnd = UBound(OriginalArray, 2)

    total_cnt = (row_up_bnd - row_low_bnd + 1) * (col_up_bnd - col_low_bnd + 1)
    ReDim res_arr(0 To total_cnt - 1)

    arr_idx = 0
    If ColumnDirection Then
        For col_index = col_low_bnd To col_up_bnd
            For row_index = row_low_bnd To row_up_bnd
                res_arr(arr_idx) = OriginalArray(row_index, col_index)
                arr_idx = arr_idx + 1
            Next row_index
        Next col_index
    Else
        For row_index = row_low_bnd To row_up_bnd
            For col_index = col_low_bnd To col_up_bnd
                res_arr(arr_idx) = OriginalArray(row_index, col_index)
                arr_idx = arr_idx + 1
            Next col_index
        Next row_index
    End If

    ConvertArray2dTo1d = res_arr
End Function


'* 1 次元配列を 2 次元配列に変換します。
'*
'* @param OriginalArray 1 次元配列。
'* @param RowLBound [省略可] 行インデックス (1 次元目) の下限。デフォルトは 1 です。
'* @param ColumnLBound [省略可] 列インデックス (2 次元目) の下限。デフォルトは 1 です。
'* @param RowCount [省略可] 行 (1 次元目) の要素数。デフォルトは無制限です。
'* @param ColCount [省略可] 列 (2 次元目) の要素数。デフォルトは 1 です。
'* @param ColumnDirection [省略可] 列方向 (縦方向、1 次元目) に埋めていくか。デフォルトは False で行列方向 (横方向、2 次元目) に埋めていきます。
'* @return 2 次元配列。
'*
'* @details
'* 1 次元配列を 2 次元配列に変換します。
Public Function ConvertArray1dTo2d( _
        ByVal OriginalArray As Variant, _
        Optional ByVal RowLBound As Long = 1, _
        Optional ByVal ColLBound As Long = 1, _
        Optional ByVal RowCount As Long = -1, _
        Optional ByVal ColCount As Long = 1, _
        Optional ByVal ColumnDirection As Boolean = False) As Variant()
    
    If Not IsArray(OriginalArray) Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="引数が配列ではありません。(" & TypeName(OriginalArray) & ")"
    End If
    
    If RowCount <= 0 And ColCount <= 0 Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="行数と列数の両方が無制限として指定 (0 以下を指定) されました。"
    End If
    
    If 0 < RowCount And 0 < ColCount Then
        Err.Raise vbObjectError + 1, Source:="Function ConvertArray1dTo2d", Description:="行数と列数の両方の要素数が指定されました。"
    End If
    
    Dim orig_count As Long
    orig_count = UBound(OriginalArray) - LBound(OriginalArray) + 1
    
    Dim count_1 As Long, count_2 As Long
    If 0 < RowCount Then
        count_1 = RowCount
        count_2 = (orig_count + RowCount - 1) \ RowCount
    Else
        count_1 = (orig_count + ColCount - 1) \ ColCount
        count_2 = ColCount
    End If
    
    Dim max_idx_1 As Long, max_idx_2 As Long
    max_idx_1 = RowLBound + count_1 - 1
    max_idx_2 = ColLBound + count_2 - 1
    
    Dim result_arr() As Variant
    ReDim result_arr(RowLBound To max_idx_1, ColLBound To max_idx_2)
    
    Dim idx_1 As Long, idx_2 As Long
    idx_1 = RowLBound
    idx_2 = ColLBound
    
    Dim enum_obj As IEnumerator
    Set enum_obj = GetArrayEnumerator(OriginalArray)
    Do While enum_obj.MoveNext()
        If IsObject(enum_obj.Current) Then
            Set result_arr(idx_1, idx_2) = enum_obj.Current
        Else
            result_arr(idx_1, idx_2) = enum_obj.Current
        End If
        
        If ColumnDirection Then
            idx_1 = idx_1 + 1
            If max_idx_1 < idx_1 Then
                idx_1 = RowLBound
                idx_2 = idx_2 + 1
            End If
        Else
            idx_2 = idx_2 + 1
            If max_idx_2 < idx_2 Then
                idx_1 = idx_1 + 1
                idx_2 = ColLBound
            End If
        End If
    Loop
    
    ConvertArray1dTo2d = result_arr
End Function

'* 配列の添字の最大および最小を得ます。
'*
'* @param LBoundArray [出力] 各次元の最小の添字
'* @param LBoundArray [出力] 各次元の最大の添字
'* @param TargetArray 調査対象の配列
'*
'* @details
'* 配列の添字の最大および最小を得ます。
Public Sub GetArrayBounds(ByRef LBoundArray() As Long, ByRef UBoundArray() As Long, ByVal TargetArray As Variant)
    If Not IsArray(TargetArray) Then
        Err.Raise vbObjectError + 1, "Function GetArrayUbound", "配列ではありません。(" & TypeName(TargetArray) & ")"
    End If
    
    Dim result_l() As Long
    Dim result_u() As Long
    
    Dim dim_count As Integer
    dim_count = 0
    
    Do
        Dim lbound_num As Long
        Dim ubound_num As Long
        On Error Resume Next
            lbound_num = LBound(TargetArray, dim_count + 1)
            ubound_num = UBound(TargetArray, dim_count + 1)
            If Err.Number <> 0 Then Exit Do
        On Error GoTo 0
        
        ReDim Preserve result_l(0 To dim_count)
        result_l(dim_count) = lbound_num
        
        ReDim Preserve result_u(0 To dim_count)
        result_u(dim_count) = ubound_num
        
        dim_count = dim_count + 1
    Loop While True
    
    LBoundArray = result_l
    UBoundArray = result_u
End Sub


'* 配列に指定された値が含まれるかを判定します。
'*
'* @param TargetArray 判定対象の配列
'* @param CheckItem 判定対象の値
'* @return 配列に値が含まれる場合は True、それ以外は False
'*
'* @details
'* 配列に指定された値が含まれるかを判定します。
Public Function IsContainsIn(ByRef TargetArray As Variant, ByVal CheckItem As Variant) As Boolean
    Dim arr_item As Variant
    
    For Each arr_item In TargetArray
        If CheckItem = arr_item Then
            IsContainsIn = True
            Exit Function
        End If
    Next arr_item
    
    IsContainsIn = False
End Function

'* 配列をソートします。
'*
'* @param TargetArray ソート対象の配列
'* @param Descending 降順ソートを行う場合は True。既定値は False (昇順ソート)
'*
'* @details
'* 配列をソートします。
'*
'* @note
'* 配列の各要素が不等号演算で比較可能であることが前提です。
Public Sub SortArray(ByRef TargetArray As Variant, Optional ByVal Descending As Boolean = False)
    Call pSortArrayCore(TargetArray, LBound(TargetArray), UBound(TargetArray), Descending)
End Sub
Private Sub pSortArrayCore(ByRef TargetArray As Variant, ByVal MinIndex As Long, ByVal MaxIndex As Long, ByVal Descending As Boolean)
    If MinIndex >= MaxIndex Then Exit Sub
    Dim idx As Long
    Dim pos_lng As Long: pos_lng = MinIndex
    
    Call pSortSwap(TargetArray(MinIndex), TargetArray(Int((MinIndex + MaxIndex) \ 2)))
    For idx = MinIndex + 1 To MaxIndex
        If psortlessthan(TargetArray(idx), TargetArray(MinIndex), Descending) Then
            pos_lng = pos_lng + 1
            Call pSortSwap(TargetArray(idx), TargetArray(pos_lng))
        End If
    Next
    Call pSortSwap(TargetArray(MinIndex), TargetArray(pos_lng))
    
    Call pSortArrayCore(TargetArray, MinIndex, pos_lng - 1)
    Call pSortArrayCore(TargetArray, pos_lng + 1, MaxIndex)
End Sub
Private Sub pSortSwap(ByRef Item1 As Variant, ByRef Item2 As Variant)
    Dim tmp_var As Variant
    tmp_var = Item1
    Item1 = Item2
    Item2 = tmp_var
End Sub
Private Function pSortIsLessThan(ByVal Item1 As Variant, ByVal Item2 As Variant, ByVal Descending As Boolean) As Variant
    If Descending Then
        If Item1 < Item2 Then
            pSortIsLessThan = True
        Else
            pSortIsLessThan = False
        End If
    Else
        If Item2 < Item1 Then
            pSortIsLessThan = True
        Else
            pSortIsLessThan = False
        End If
    End If
End Function

'* 配列を連結します。
'*
'* @param Array1 最初の配列
'* @param Array2 2 番目の配列
'* @param OtherArrays 可変長の追加配列
'* @return 連結された配列
'*
'* @details
'* 複数の配列を結合して 1 つの配列を返します。
'* 引数が配列でない場合は長さ 1 の配列として扱われます。
Public Function ConcatArray(ByVal Array1 As Variant, ByVal Array2 As Variant, ParamArray OtherArrays() As Variant) As Variant()
    Dim result() As Variant
    Dim arr_length As Long
        
    arr_length = pGetArrayLengthCore(Array1)
    arr_length = arr_length + pGetArrayLengthCore(Array2)
    
    Dim param_idx As Long
    Dim param_item As Variant
    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            param_item = OtherArrays(param_idx)
            arr_length = arr_length + pGetArrayLengthCore(param_item)
        Next param_idx
    End If
    
    If arr_length = 0 Then
        ConcatArray = Array()
        Exit Function
    End If
    
    'If IsArray(Array1) And Not IsEmptyArray(Array1) Then
    '    ReDim result(LBound(Array1) To LBound(Array1) + arr_length - 1)
    'Else
        ReDim result(0 To arr_length - 1)
    'End If
    
    Dim result_idx As Long
    result_idx = LBound(result)
    
    Call pConcatArrayCore(result, result_idx, Array1)
    Call pConcatArrayCore(result, result_idx, Array2)
    
    If UBound(OtherArrays) <> -1 Then
        For param_idx = LBound(OtherArrays) To UBound(OtherArrays)
            param_item = OtherArrays(param_idx)
            Call pConcatArrayCore(result, result_idx, param_item)
        Next param_idx
    End If
    
    ConcatArray = result
End Function

Private Function pGetArrayLengthCore(ByVal TargetArray As Variant) As Long
    Dim result As Long
    
    If IsEmptyArray(TargetArray) Then
        result = 0
    ElseIf IsArray(TargetArray) Then
        result = UBound(TargetArray) - LBound(TargetArray) + 1
    Else
        result = 1
    End If
    
    pGetArrayLengthCore = result
End Function

Private Sub pConcatArrayCore(ByRef ResultArray() As Variant, ByRef CurrentIndex As Long, ByVal TargetArray As Variant)
    If IsEmptyArray(TargetArray) Then
        Exit Sub
    End If
    
    Dim idx As Long
    If IsArray(TargetArray) Then
        For idx = LBound(TargetArray) To UBound(TargetArray)
            ResultArray(CurrentIndex) = TargetArray(idx)
            CurrentIndex = CurrentIndex + 1
        Next idx
    Else
        ResultArray(CurrentIndex) = TargetArray
        CurrentIndex = CurrentIndex + 1
    End If
End Sub

'* 引数が空の配列かを判定します。
'*
'* @param TargetArray 判定対象の配列
'* @return 空の配列の場合は True、それ以外は False
'*
'* @details
'* 指定された引数が配列であり、要素が存在しない場合に True を返します。
'* 配列でない場合は False を返します。
'* 配列の境界値取得 (`UBound` / `LBound`) が失敗する場合は空の配列とみなします。
Public Function IsEmptyArray(ByVal TargetArray As Variant) As Boolean
    If Not IsArray(TargetArray) Then
        IsEmptyArray = False
        Exit Function
    End If
    
    Dim result As Boolean
    result = False
    
    Err.Clear
    On Error Resume Next
    result = UBound(TargetArray) < LBound(TargetArray)
    If Err.Number <> 0 Then
        result = True
        Err.Clear
    End If
    On Error GoTo 0
    
    IsEmptyArray = result
End Function

'* 配列の列挙子を取得します。
'*
'* @param TargetArray 列挙対象の配列
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 配列の列挙子 (IEnumerator オブジェクト)
'*
'* @details
'* 指定された配列を列挙するための `IEnumerator` オブジェクトを返します。
'* 列挙子を使用することで、配列の要素を簡単に反復処理できます。
'* 配列がオブジェクトの場合はそのまま参照を設定し、それ以外の場合は値を設定します。
'*
'* 使用例:
'* @code
'* Dim enum_obj As IEnumerator
'* Set enum_obj = GetArrayEnumerator(some_arr)
'* Do While enum_obj.MoveNext()
'*     Dim item_var as Variable
'*     item_var = enum_obj.Current
'*     'Debug.Print item_var
'* Loop
'* @endcode
Public Function GetArrayEnumerator(ByVal TargetArray As Variant, Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result As Enumerator
    Set result = New Enumerator
    Call result.Initialize(TargetArray, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetArrayEnumerator = result
End Function

' #############################################################################
'
' String 関連
'
' #############################################################################

'
'「01.ほげほげ」「02.ふがふが」のような形式の文字列を返します。
'
Public Function FormatIDName( _
        ByVal IDNumber As Integer, _
        ByVal Name As String, _
        Optional ByVal NumFormat As String = "00", _
        Optional ByVal sep As String = ".") As Variant
    
    FormatIDName = Format(IDNumber, NumFormat) & sep & Name
End Function

'* 複数キーを辞書で使用するために、キーを連結して返します。
'*
'* @param DictionaryKey1 最初のキー
'* @param DictionaryKeys 追加のキー (可変長引数)
'* @return 連結されたキー文字列
'*
'* @details
'* 指定された複数のキーを連結し、1 つの文字列として返します。
'* キーに使用される区切り文字はタブ文字 (`vbTab`) です。
Public Function GetMultiKey(ByVal DictionaryKey1 As Variant, ParamArray DictionaryKeys() As Variant) As String
    Dim result As String
    
    ' 一つ目のキーで初期化
    result = pGetMultiKeyEscape(CStr(DictionaryKey1))
    
    ' 他関数に ParamArray を引数で渡すために Variant 型に明示的に変換する
    Dim key_arr As Variant
    key_arr = DictionaryKeys
    
    Dim key_enum As IEnumerator
    Set key_enum = GetArrayEnumerator(key_arr)
    Do While key_enum.MoveNext()
        ' 残りのキーを連結
        result = result & vbTab & pGetMultiKeyEscape(CStr(key_enum.Current))
    Loop
    
    GetMultiKey = result
End Function
Private Function pGetMultiKeyEscape(ByVal DictionaryKey As String) As String
    pGetMultiKeyEscape = Replace(Replace(DictionaryKey, "\", "\\"), vbTab, "\t")
End Function

'* 文字列の配列の diff を取ります。
'*
'* @param OldArray [入出力] 古い文字列の配列。ReDim 可能である必要があります。
'* @param NewArray [入出力] 新しい文字列の配列。ReDim 可能である必要があります。
'* @param ChangeTypeArray [出力] 変更種別の配列。ReDim 可能である必要があります。渡される配列は、未初期化状態でも構いません。未初期化状態で渡された場合、LBound は OldArray と揃います。
'* @param EnableReplaceType [省略可] 削除と追加が連続する部分を置換として扱うか否か。デフォルトは True で、置換として扱います。
'* @param ReplaceCost [省略可] 削除と追加のコストを 1 としたとき、置換のコストをどうするか (0～2)。デフォルトは 1.5。
'*
'* @details
'* 配列 OldArray, NewArray の最小編集距離を算出し、
'* (Old行, New行, 変更種別) の並びをそれぞれの配列に上書きします。
'*
'* 変更種別は "" (一致), "DEL" (削除), "ADD" (追加), "MOD" (置換) のいずれかとなります。
'* 種別が一致の場合は OldArray、NewArray の両方の当該インデックスは元の文字列になります。
'* 削除の場合は、OldArray の当該インデックスは元の文字列が入り、NewArray の当該インデックスは空文字列となります。
'* 追加の場合は、NewArray の当該インデックスは新しい文字列が入り、OldArray の当該インデックスは空文字列となります。
'* 置換の場合は、OldArray の当該インデックスは元の文字列が、NewArray の当該インデックスは新しい文字列が入ります。
Public Sub DiffStringArray( _
        ByRef OldArray() As String, _
        ByRef NewArray() As String, _
        ByRef ChangeTypeArray() As String, _
        Optional ByVal EnableReplaceType As Boolean = True, _
        Optional ByVal ReplaceCost As Double = 1.5)
    
    Dim diff_old_coll As New Collection
    Dim diff_new_coll As New Collection
    Dim diff_op_coll  As New Collection
    
    If Not EnableReplaceType Then ReplaceCost = 3
    
    ' 再帰的に差分計算（Hirschberg アルゴリズム）
    Call pHirschbergDiffRecursive(OldArray, LBound(OldArray), UBound(OldArray), _
                              NewArray, LBound(NewArray), UBound(NewArray), _
                              diff_old_coll, diff_new_coll, diff_op_coll, ReplaceCost)
                              
    Dim num_items_val As Long, loop_idx_val As Long
    num_items_val = diff_old_coll.Count
    ReDim OldArray(0 To num_items_val - 1)
    ReDim NewArray(0 To num_items_val - 1)
    ReDim ChangeTypeArray(0 To num_items_val - 1)
    
    For loop_idx_val = 1 To num_items_val
        OldArray(loop_idx_val - 1) = diff_old_coll(loop_idx_val)
        NewArray(loop_idx_val - 1) = diff_new_coll(loop_idx_val)
        ChangeTypeArray(loop_idx_val - 1) = diff_op_coll(loop_idx_val)
    Next loop_idx_val
End Sub

Private Sub pHirschbergDiffRecursive(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)
    
    Dim SMALL_THRESHOLD As Long: SMALL_THRESHOLD = 1
    
    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1
    
    If old_count_val = 0 Then
        Dim loop_idx_new As Long
        For loop_idx_new = StartIdxNew To EndIdxNew
            DiffOldColl.Add ""
            DiffNewColl.Add NewArr(loop_idx_new)
            DiffOpColl.Add "ADD"
        Next loop_idx_new
        Exit Sub
    ElseIf new_count_val = 0 Then
        Dim loop_idx_old As Long
        For loop_idx_old = StartIdxOld To EndIdxOld
            DiffOldColl.Add OldArr(loop_idx_old)
            DiffNewColl.Add ""
            DiffOpColl.Add "DEL"
        Next loop_idx_old
        Exit Sub
    ElseIf old_count_val <= SMALL_THRESHOLD Or new_count_val <= SMALL_THRESHOLD Then
        Call pDiffSmall(OldArr, StartIdxOld, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
        Exit Sub
    End If
    
    Dim mid_idx_val As Long
    mid_idx_val = (StartIdxOld + EndIdxOld) \ 2
    
    Dim forward_costs_arr As Variant
    forward_costs_arr = pComputeForwardCost(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)
    
    Dim backward_costs_arr As Variant
    backward_costs_arr = pComputeBackwardCost(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew, EndIdxNew, ReplaceCost)
    
    Dim split_j_val As Long, loop_idx_new_val As Long
    Dim min_cost_val As Double
    min_cost_val = 1E+99
    Dim new_size_val As Long
    new_size_val = new_count_val
    Dim cur_cost_val As Double
    For loop_idx_new_val = 0 To new_size_val
        cur_cost_val = forward_costs_arr(loop_idx_new_val) + backward_costs_arr(new_size_val - loop_idx_new_val)
        If cur_cost_val < min_cost_val Then
            min_cost_val = cur_cost_val
            split_j_val = loop_idx_new_val
        End If
    Next loop_idx_new_val
    
    Call pHirschbergDiffRecursive(OldArr, StartIdxOld, mid_idx_val, NewArr, StartIdxNew, StartIdxNew + split_j_val - 1, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
    Call pHirschbergDiffRecursive(OldArr, mid_idx_val + 1, EndIdxOld, NewArr, StartIdxNew + split_j_val, EndIdxNew, _
                              DiffOldColl, DiffNewColl, DiffOpColl, ReplaceCost)
End Sub

Private Function pComputeForwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant
    
    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0
    
    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)
    
    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val
    
    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = StartIdxOld To EndIdxOld
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(StartIdxNew + inner_idx_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val
    
    pComputeForwardCost = cost_arr
End Function

Private Function pComputeBackwardCost(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, ByVal ReplaceCost As Double) As Variant
    
    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0
    
    Dim new_count_val As Long, loop_new_val As Long, loop_old_val As Long, inner_idx_val As Long
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim cost_arr() As Double
    ReDim cost_arr(0 To new_count_val)
    
    cost_arr(0) = 0
    For loop_new_val = 1 To new_count_val
        cost_arr(loop_new_val) = cost_arr(loop_new_val - 1) + COST_ADD
    Next loop_new_val
    
    Dim prev_cost_val As Double, temp_cost_val As Double, sub_cost_val As Double
    For loop_old_val = EndIdxOld To StartIdxOld Step -1
        prev_cost_val = cost_arr(0)
        cost_arr(0) = cost_arr(0) + COST_DEL
        For inner_idx_val = 1 To new_count_val
            temp_cost_val = cost_arr(inner_idx_val)
            If OldArr(loop_old_val) = NewArr(EndIdxNew - inner_idx_val + 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            cost_arr(inner_idx_val) = Application.Min(cost_arr(inner_idx_val) + COST_DEL, _
                                                       cost_arr(inner_idx_val - 1) + COST_ADD, _
                                                       prev_cost_val + sub_cost_val)
            prev_cost_val = temp_cost_val
        Next inner_idx_val
    Next loop_old_val
    
    pComputeBackwardCost = cost_arr
End Function

Private Sub pDiffSmall(ByRef OldArr() As String, ByVal StartIdxOld As Long, ByVal EndIdxOld As Long, _
        ByRef NewArr() As String, ByVal StartIdxNew As Long, ByVal EndIdxNew As Long, _
        ByRef DiffOldColl As Collection, ByRef DiffNewColl As Collection, ByRef DiffOpColl As Collection, ByVal ReplaceCost As Double)
    
    Dim COST_ADD As Double: COST_ADD = 1
    Dim COST_DEL As Double: COST_DEL = 1
    Dim COST_EQUAL As Double: COST_EQUAL = 0
    
    Dim old_count_val As Long, new_count_val As Long
    old_count_val = EndIdxOld - StartIdxOld + 1
    new_count_val = EndIdxNew - StartIdxNew + 1
    Dim dp_arr() As Double, op_arr() As String
    ReDim dp_arr(0 To old_count_val, 0 To new_count_val)
    ReDim op_arr(0 To old_count_val, 0 To new_count_val)
    
    Dim loop_old_val As Long, loop_new_val As Long
    dp_arr(0, 0) = 0
    op_arr(0, 0) = "DONE"
    For loop_old_val = 1 To old_count_val
        dp_arr(loop_old_val, 0) = dp_arr(loop_old_val - 1, 0) + COST_DEL
        op_arr(loop_old_val, 0) = "DEL"
    Next loop_old_val
    For loop_new_val = 1 To new_count_val
        dp_arr(0, loop_new_val) = dp_arr(0, loop_new_val - 1) + COST_ADD
        op_arr(0, loop_new_val) = "ADD"
    Next loop_new_val
    
    Dim inner_old_val As Long, inner_new_val As Long
    Dim sub_cost_val As Double, del_cost_val As Double, add_cost_val As Double, mod_cost_val As Double
    For inner_old_val = 1 To old_count_val
        For inner_new_val = 1 To new_count_val
            If OldArr(StartIdxOld + inner_old_val - 1) = NewArr(StartIdxNew + inner_new_val - 1) Then
                sub_cost_val = COST_EQUAL
            Else
                sub_cost_val = ReplaceCost
            End If
            del_cost_val = dp_arr(inner_old_val - 1, inner_new_val) + COST_DEL
            add_cost_val = dp_arr(inner_old_val, inner_new_val - 1) + COST_ADD
            mod_cost_val = dp_arr(inner_old_val - 1, inner_new_val - 1) + sub_cost_val
            dp_arr(inner_old_val, inner_new_val) = Application.Min(del_cost_val, add_cost_val, mod_cost_val)
            If dp_arr(inner_old_val, inner_new_val) = mod_cost_val Then
                op_arr(inner_old_val, inner_new_val) = IIf(sub_cost_val = COST_EQUAL, "", "MOD")
            ElseIf dp_arr(inner_old_val, inner_new_val) = del_cost_val Then
                op_arr(inner_old_val, inner_new_val) = "DEL"
            Else
                op_arr(inner_old_val, inner_new_val) = "ADD"
            End If
        Next inner_new_val
    Next inner_old_val
    
    Dim pos_old_val As Long, pos_new_val As Long
    pos_old_val = old_count_val: pos_new_val = new_count_val
    Dim temp_old_arr() As String, temp_new_arr() As String, temp_op_arr() As String
    ReDim temp_old_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_new_arr(0 To old_count_val + new_count_val - 1)
    ReDim temp_op_arr(0 To old_count_val + new_count_val - 1)
    Dim step_idx_val As Long: step_idx_val = 0
    Do While pos_old_val > 0 Or pos_new_val > 0
        Dim cur_op_val As String
        cur_op_val = op_arr(pos_old_val, pos_new_val)
        Select Case cur_op_val
            Case ""
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = ""
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "MOD"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "MOD"
                pos_old_val = pos_old_val - 1: pos_new_val = pos_new_val - 1
            Case "DEL"
                temp_old_arr(step_idx_val) = OldArr(StartIdxOld + pos_old_val - 1)
                temp_new_arr(step_idx_val) = ""
                temp_op_arr(step_idx_val) = "DEL"
                pos_old_val = pos_old_val - 1
            Case "ADD"
                temp_old_arr(step_idx_val) = ""
                temp_new_arr(step_idx_val) = NewArr(StartIdxNew + pos_new_val - 1)
                temp_op_arr(step_idx_val) = "ADD"
                pos_new_val = pos_new_val - 1
        End Select
        step_idx_val = step_idx_val + 1
    Loop
    
    Dim num_steps_val As Long: num_steps_val = step_idx_val
    Dim rev_old_arr() As String, rev_new_arr() As String, rev_op_arr() As String
    ReDim rev_old_arr(0 To num_steps_val - 1)
    ReDim rev_new_arr(0 To num_steps_val - 1)
    ReDim rev_op_arr(0 To num_steps_val - 1)
    For loop_old_val = 0 To num_steps_val - 1
        rev_old_arr(loop_old_val) = temp_old_arr(num_steps_val - 1 - loop_old_val)
        rev_new_arr(loop_old_val) = temp_new_arr(num_steps_val - 1 - loop_old_val)
        rev_op_arr(loop_old_val) = temp_op_arr(num_steps_val - 1 - loop_old_val)
    Next loop_old_val
    
    Dim loop_idx_final_val As Long
    For loop_idx_final_val = 0 To num_steps_val - 1
        DiffOldColl.Add rev_old_arr(loop_idx_final_val)
        DiffNewColl.Add rev_new_arr(loop_idx_final_val)
        DiffOpColl.Add rev_op_arr(loop_idx_final_val)
    Next loop_idx_final_val
End Sub

'* String の配列が空かどうかを判定します。
'*
'* @param StringArray 判定対象の String 型配列
'* @return 空の場合は True、それ以外は False
'*
'* @details
'* 配列が初期化されていない場合、要素が 1 未満の場合、または唯一の要素が空文字列の場合に True を返します。
Public Function IsEmptyStringArray(ByRef StringArray() As String) As Boolean
    If (Not StringArray) = -1 Then
        ' 未初期化
        IsEmptyStringArray = True
    ElseIf LBound(StringArray) < UBound(StringArray) Then
        ' 長さが 1 以上
        IsEmptyStringArray = False
    ElseIf UBound(StringArray) < LBound(StringArray) Then
        ' 長さが 1 未満
        IsEmptyStringArray = True
    ElseIf StringArray(LBound(StringArray)) = "" Then
        ' 長さが 1 で、内容が空文字列
        IsEmptyStringArray = True
    Else
        ' それ以外
        IsEmptyStringArray = False
    End If
End Function

'* 文字列に対して複数条件で置換を実行します。
'*
'* @param Expression 対象の文字列
'* @param Find1 最初の置換対象文字列
'* @param Replace1 最初の置換後文字列 (文字列または配列)
'* @param FindReplacePairs 可変長の置換対象文字列と置換後文字列のペア
'* @return 置換後の文字列配列
'*
'* @details
'* 指定された条件で文字列を置換し、すべての結果を配列で返します。
'* 置換後文字列が配列の場合、それぞれの組み合わせを結果に含めます。
Public Function ReplaceMulti(ByVal Expression As String, ByVal Find1 As String, ByVal Replace1 As Variant, ParamArray FindReplacePairs() As Variant) As String()
    Dim result_list As ObjectList
    Set result_list = New ObjectList
    
    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        If (UBound(FindReplacePairs) - LBound(FindReplacePairs) + 1) Mod 2 <> 0 Then
            Err.Raise vbObjectError + 1, "Function ReplaceMulti", "可変長文字列が奇数個です。置換対象と置換後文字列の対応が取れていません。"
            Exit Function
        End If
    End If
    
    Call result_list.Add(Expression)
    Call pReplaceMultiCore(result_list, Find1, Replace1)
    
    If LBound(FindReplacePairs) <= UBound(FindReplacePairs) Then
        Dim idx As Long
        For idx = LBound(FindReplacePairs) To UBound(FindReplacePairs) - 1 Step 2
            Call pReplaceMultiCore(result_list, FindReplacePairs(idx), FindReplacePairs(idx + 1))
        Next idx
    End If
    
    ReplaceMulti = result_list.ConvertToStringArray()
End Function

Private Sub pReplaceMultiCore(ByRef ResultList As ObjectList, ByVal FindString As String, ByVal Replaces As Variant)
    Dim result As ObjectList
    Set result = New ObjectList
    
    Dim enum_obj As IEnumerator
    Set enum_obj = ResultList.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim current_item As String
        current_item = enum_obj.Current
        If IsArray(Replaces) Then
            Dim replace_item As Variant
            For Each replace_item In Replaces
                Call result.Add(Replace(current_item, FindString, replace_item))
            Next replace_item
        Else
            Call result.Add(Replace(current_item, FindString, Replaces))
        End If
    Loop
    
    Set ResultList = result
End Sub

'* 改行記号をエスケープします。
'*
'* @param Expression 対象の文字列
'* @param EscSeqChar エスケープ記号 (既定値は `\`)
'* @return 改行記号をエスケープした文字列
'*
'* @detailsｌ
'* 改行記号 (`vbCr`, `vbLf`) を指定されたエスケープ記号で置換します。
Public Function EscapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "") As String
    If EscSeqChar = "" Then EscSeqChar = "\"
    
    Dim result As String
    result = Replace(Expression, EscSeqChar, EscSeqChar & EscSeqChar)
    result = Replace(result, vbCr, EscSeqChar & "r")
    result = Replace(result, vbLf, EscSeqChar & "n")
    EscapeLineSeparator = result
End Function

'* 改行記号のエスケープを解除します。
'*
'* @param Expression 対象の文字列
'* @param EscSeqChar エスケープ記号 (既定値は `\`)
'* @return エスケープ解除後の文字列
'*
'* @details
'* エスケープされた改行記号 (`\n`, `\r`) を元の改行記号 (`vbCr`, `vbLf`) に戻します。
Public Function UnescapeLineSeparator(ByVal Expression As String, Optional ByVal EscSeqChar As String = "\") As String
    Dim result As String
    Dim idx As Long
    Dim str_len As Long
    Dim cur_char As String
    
    str_len = Len(Expression)
    idx = 1
    result = ""
    
    Do While idx <= str_len
        cur_char = Mid$(Expression, idx, 1)
        
        If cur_char = EscSeqChar Then
            ' エスケープシーケンスの開始
            If idx < str_len Then
                Dim next_char As String
                next_char = Mid$(Expression, idx + 1, 1)
                
                Select Case next_char
                    Case "n"
                        result = result & vbLf
                        idx = idx + 2
                    Case "r"
                        result = result & vbCr
                        idx = idx + 2
                    Case Else
                        result = result & next_char
                        idx = idx + 2
                End Select
            Else
                ' 文字列の最後にエスケープ文字がある場合
                Err.Raise vbObjectError + 512, "Function UnescapeLineSeparator", "文字列の最後にエスケープ文字があります。"
            End If
        Else
            ' エスケープ文字でない場合はそのまま追加
            result = result & cur_char
            idx = idx + 1
        End If
    Loop
    
    UnescapeLineSeparator = result
End Function

'* 改行記号で文字列を分割して配列を返します。
'*
'* @param StringList 対象の文字列
'* @return 改行記号で分割された文字列配列
'*
'* @details
'* 改行記号に基づいて文字列を分割し、配列で返します。
Public Function SplitByLineSeparator(ByVal StringList As String) As String()
    SplitByLineSeparator = Split(pUnifyLineSeparatorCore(StringList), vbLf)
End Function

'* 改行記号を統一します。
'*
'* @param Expression 対象の文字列
'* @param LineSep 統一後の改行記号 (既定値は `vbLf`)
'* @return 統一後の文字列
'*
'* @details
'* 改行記号を指定された記号 (`vbLf`, `vbCr`, `vbCrLf`) に置換します。
Public Function UnifyLineSeparator(ByVal Expression As String, Optional ByVal LineSep As String = vbLf) As String
    If LineSep <> vbLf And LineSep <> vbCr And LineSep <> vbCrLf Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function UnifyLineSeparator", Description:="改行記号ではありません。(" & LineSep & ")"
    End If
    
    UnifyLineSeparator = Replace(pUnifyLineSeparatorCore(Expression), vbLf, LineSep)
End Function

Private Function pUnifyLineSeparatorCore(ByVal Expression As String) As String
    pUnifyLineSeparatorCore = Replace(Replace(Expression, vbCrLf, vbLf), vbCr, vbLf)
End Function

'* String が格納された ObjectList を、Delimiter で区切って一つの String にします。
'*
'* @param SourceList 文字列を格納した ObjectList
'* @param Delimiter 区切り文字 (既定値は半角スペース)
'* @return 区切り文字で連結された文字列
'*
'* @details
'* 指定された文字列リストを区切り文字で連結し、単一の文字列として返します。
Public Function JoinStringList(ByVal SourceList As ObjectList, Optional ByVal Delimiter As String = " ") As String
    Dim result As String
    
    If SourceList.Count = 0 Then Exit Function
    
    result = SourceList.Item(0)
    
    Dim idx As Long
    For idx = 1 To SourceList.Count - 1
        result = result & Delimiter & SourceList.Item(idx)
    Next idx
    
    JoinStringList = result
End Function

'* String が格納された ObjectSet を、Delimiter で区切って一つの String にします。
'*
'* @param SourceSet 文字列を格納した ObjectSet
'* @param Delimiter 区切り文字 (既定値は半角スペース)
'* @return 区切り文字で連結された文字列
'*
'* @details
'* 指定された文字列セットを区切り文字で連結し、単一の文字列として返します。
Public Function JoinStringSet(ByVal SourceSet As ObjectSet, Optional ByVal Delimiter As String = " ") As String
    Dim result As String
    
    If SourceSet.Count = 0 Then Exit Function
    
    result = SourceSet.Item(0)
    
    Dim idx As Long
    For idx = 1 To SourceSet.Count - 1
        result = result & Delimiter & SourceSet.Item(idx)
    Next idx
    
    JoinStringSet = result
End Function

'* メッセージ文字列をページに分割して表示する MsgBox です。
'*
'* @param MessageString 表示するメッセージ文字列
'* @param Title MsgBox のタイトル (オプション)
'*
'* @details
'* 指定されたメッセージをページに分割し、ページごとに MsgBox で表示します。
Public Sub MsgBoxPage(ByVal MessageString As String, Optional ByVal Title As String = "")
    Dim msg_pages() As String
    Dim msg_page As Variant 'String
    Dim total_num As Long
    Dim page_num As Long
    
    msg_pages = SplitMessage(MessageString)
    
    total_num = UBound(msg_pages) - LBound(msg_pages) + 1
    
    page_num = 1
    For Each msg_page In msg_pages
        MsgBox msg_page, Title:=Title & "(" & page_num & "/" & total_num & ")"
        page_num = page_num + 1
    Next msg_page
End Sub

'* メッセージ文字列をページに分割します。
'*
'* @param MessageString 対象の文字列
'* @param PageSize ページの最大サイズ (バイト数) (既定値は 1023)
'* @return ページに分割された文字列の配列
'*
'* @details
'* 指定されたメッセージ文字列を改行単位で分割し、指定されたバイト数の範囲でページ化します。
Public Function SplitMessage(ByVal MessageString As String, Optional ByVal PageSize As Long = 1023) As Variant
    Dim msgs_list As ObjectList
    Dim msg_str As String
    Dim lines_arr() As String
    Dim line_str As Variant 'String
    Dim taken_str As String
    Dim rem_str As String
    Dim result() As String
    Dim idx As Long
    
    Set msgs_list = New ObjectList
    
    If MessageString = "" Then
        ReDim result(0 To 0)
        result(0) = ""
        SplitMessage = result
        Exit Function
    End If
    
    lines_arr() = SplitByLineSeparator(MessageString)
    
    For Each line_str In lines_arr
        If msg_str = "" And LenB(StrConv(line_str, vbFromUnicode)) <= PageSize _
                Or LenB(StrConv(msg_str, vbFromUnicode)) + 2 + LenB(StrConv(line_str, vbFromUnicode)) <= PageSize Then ' 2 は CRLF 分
            ' この行を追加しても、ページ サイズ以内の場合
            ' メッセージに行を追加する。
            If msg_str = "" Then
                msg_str = line_str ' 先頭行
            Else
                msg_str = msg_str & vbCrLf & line_str
            End If
        Else
            ' この行を追加すると、ページ サイズを超える場合
            If msg_str <> "" Then
                ' メッセージが空でないなら、現在のメッセージをそのままページとして追加する。
                Call msgs_list.Add(msg_str)
            End If
            
            Do While LenB(StrConv(line_str, vbFromUnicode)) > PageSize
                ' この行がページサイズを超えている間、ページサイズ分を切り出して追加。
                Call pTakeString(taken_str, rem_str, line_str, PageSize)
                Call msgs_list.Add(taken_str)
                line_str = rem_str ' ページとして追加しなかった行の残りを現在の行とする
            Loop
            
            msg_str = line_str
        End If
    Next line_str
    
    If LenB(StrConv(msg_str, vbFromUnicode)) > 0 Then Call msgs_list.Add(msg_str)
    
    ReDim result(0 To msgs_list.Count - 1)
    For idx = LBound(result) To UBound(result)
        result(idx) = msgs_list.Item(idx)
    Next idx
    
    SplitMessage = result
End Function

Private Sub pTakeString( _
        ByRef TakenString As String, _
        ByRef RemainingString As String, _
        ByVal Expression As String, _
        ByVal LengthByte As Integer)
    
    Dim char As String
    Dim idx As Long
    Dim result As String
    Dim total_bytes As Long
    Dim current_bytes As Long

    For idx = 1 To Len(Expression)
        char = Mid(Expression, idx, 1)
        current_bytes = LenB(StrConv(char, vbFromUnicode))
        
        If total_bytes + current_bytes > LengthByte Then Exit For
        
        result = result + char
        total_bytes = total_bytes + current_bytes
    Next idx
    
    TakenString = result
    RemainingString = Mid(Expression, idx)
End Sub

'* 文字列の前後の半角空白、タブ、改行記号を削除します。
'*
'* @param Expression 対象の文字列
'* @param IgnoreHead 先頭の空白を無視する場合は True (既定値は False)
'* @param IgnoreTail 末尾の空白を無視する場合は True (既定値は False)
'* @param RemoveFullWidthSpace 全角空白も削除する場合は True (既定値は False)
'* @return 前後の空白を削除した文字列
'*
'* @details
'* 指定された条件に基づいて、文字列の前後から空白文字を削除します。
Public Function Strip( _
        ByVal Expression As String, _
        Optional ByVal IgnoreHead As Boolean = False, _
        Optional ByVal IgnoreTail As Boolean = False, _
        Optional ByVal RemoveFullWidthSpace As Boolean = False) As String
    
    Dim head_pos As Integer
    Dim tail_pos As Integer
    Dim test_char As String
    
    head_pos = 1
    tail_pos = Len(Expression)
    
    If tail_pos = 0 Then
        Strip = Expression
        Exit Function
    End If
    
    If Not IgnoreHead Then
        ' 先頭の空白、タブ、改行を削除
        Do While head_pos <= tail_pos
            test_char = Mid(Expression, head_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If
                          
            head_pos = head_pos + 1
        Loop
    End If
    
    If Not IgnoreTail Then
        ' 末尾の空白、タブ、改行を削除
        Do While tail_pos >= head_pos
            test_char = Mid(Expression, tail_pos, 1)
            If Not pIsWhitespace(test_char, RemoveFullWidthSpace) Then
                Exit Do
            End If
                
            tail_pos = tail_pos - 1
        Loop
    End If
    
    ' 最終的な文字列を返す
    Strip = Mid(Expression, head_pos, tail_pos - head_pos + 1)
End Function

Private Function pIsWhitespace(ByVal Character As String, ByVal IncludeFullWidthSpace As Boolean) As Boolean
    Dim result As Boolean
    
    If Character = " " Then
        result = True
    ElseIf Character = vbTab Then
        result = True
    ElseIf Character = vbCrLf Then
        result = True
    ElseIf Character = vbCr Then
        result = True
    ElseIf Character = vbLf Then
        result = True
    ElseIf (Character = "　" And IncludeFullWidthSpace) Then
        result = True
    Else
        result = False
    End If
    
    pIsWhitespace = result
End Function

'* 文字列が指定の文字列で始まるかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param SearchString 検索対象の文字列
'* @return 指定の文字列で始まる場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で始まるかを判定します。
Public Function StartsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If Len(Expression) < 1 Or Len(Expression) < Len(SearchString) Then
        StartsWith = False
        Exit Function
    End If
    
    If Left(Expression, Len(SearchString)) = SearchString Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

'* 文字列が指定の文字列で終わるかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param SearchString 検索対象の文字列
'* @return 指定の文字列で終わる場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で終わるかを判定します。
Public Function EndsWith(ByVal Expression As String, ByVal SearchString As String) As Boolean
    If Len(Expression) < 1 Or Len(Expression) < Len(SearchString) Then
        EndsWith = False
        Exit Function
    End If
    
    If Right(Expression, Len(SearchString)) = SearchString Then
        EndsWith = True
    Else
        EndsWith = False
    End If
End Function

'* 文字列が指定の文字列で括られているかを判定します。
'*
'* @param Expression 判定対象の文字列
'* @param QuoteString 開始文字列
'* @param EndString 終了文字列 (既定値は `QuoteString`) 終了文字列が省略された場合、開始文字列と同じものを使用します。
'* @return 指定の文字列で括られている場合は True、それ以外は False
'*
'* @details
'* 判定対象の文字列が指定の文字列で括られているかを判定します。
Public Function IsQuotedWith(ByVal Expression As String, ByVal QuoteString As String, Optional ByVal EndString As String = "") As Boolean
    If EndString = "" Then
        EndString = QuoteString
    End If
    
    If Len(Expression) < 2 Or Len(Expression) < (Len(QuoteString) + Len(EndString)) Then
        IsQuotedWith = False
        Exit Function
    End If
    
    If Left(Expression, Len(QuoteString)) = QuoteString And Right(Expression, Len(EndString)) = EndString Then
        IsQuotedWith = True
    Else
        IsQuotedWith = False
    End If
End Function

' #############################################################################
'
' 数値関連
'
' #############################################################################

'* 引数として渡された Long 型の値のうち、最大のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最大の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MaxLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result As Long
    
    result = Number1
    If result < Number2 Then result = Number2
    
    If UBound(Numbers) <> -1 Then
        Dim idx As Long
        For idx = LBound(Numbers) To UBound(Numbers)
            If result < Numbers(idx) Then result = Numbers(idx)
        Next idx
    End If
    
    MaxLng = result
End Function

'* 引数として渡された Double 型の値のうち、最大のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最大の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MaxDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result As Double
    
    result = Number1
    If result < Number2 Then result = Number2
    
    If UBound(Numbers) <> -1 Then
        Dim idx As Long
        For idx = LBound(Numbers) To UBound(Numbers)
            If result < Numbers(idx) Then result = Numbers(idx)
        Next idx
    End If
    
    MaxDbl = result
End Function

'* 引数として渡された Long 型の値のうち、最小のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最小の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MinLng(ByVal Number1 As Long, ByVal Number2 As Long, ParamArray Numbers() As Variant) As Long
    Dim result As Long
    
    result = Number1
    If Number2 < result Then result = Number2
    
    If UBound(Numbers) <> -1 Then
        Dim idx As Long
        For idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(idx) < result Then result = Numbers(idx)
        Next idx
    End If
    
    MinLng = result
End Function

'* 引数として渡された Double 型の値のうち、最小のものを返します。
'*
'* @param Number1 最初の数値
'* @param Number2 2 番目の数値
'* @param Numbers その他の数値 (可変長引数)
'* @return 最小の数値
'*
'* @details
'* 最低 2 つの引数を指定し、それに加えて可変長引数として任意の個数の数値を指定できます。
Public Function MinDbl(ByVal Number1 As Double, ByVal Number2 As Double, ParamArray Numbers() As Variant) As Double
    Dim result As Double
    
    result = Number1
    If Number2 < result Then result = Number2
    
    If UBound(Numbers) <> -1 Then
        Dim idx As Long
        For idx = LBound(Numbers) To UBound(Numbers)
            If Numbers(idx) < result Then result = Numbers(idx)
        Next idx
    End If
    
    MinDbl = result
End Function

' #############################################################################
'
' Integer, Long 関連
'
' #############################################################################

'* Long 型の値を 2 進数表記の文字列に変換します。
'*
'* @param LongValue 変換対象の Long 型の値
'* @return 2 進数表記の文字列
'*
'* @details
'* 指定された Long 型の値を 2 進数表記の文字列に変換します。
Public Function LongToBin(ByVal LongValue As Long) As String
    Dim high_bit As String
    Dim long_value As Long
    Dim idx As Integer
    Dim result As String
    
    If 0 <= LongValue Then
        high_bit = "0"
        long_value = LongValue
    Else
        high_bit = "1"
        long_value = LongValue And &H7FFFFFFF
    End If
    
    For idx = 1 To 31
        result = (long_value Mod 2) & result
        long_value = long_value \ 2
    Next idx
    
    result = high_bit & result
    
    LongToBin = result
End Function

'Sub Test_BitShift()
'    Dim test_values(0 To 7) As Long
'    Dim arr_idx As Integer
'    'For arr_idx = 0 To 3
'    '    test_values(arr_idx) = 32767& * Rnd
'    '    test_values(arr_idx) = test_values(arr_idx) * (2 ^ 16) + 32767& * Rnd
'    '    If Rnd < 0.5 Then test_values(arr_idx) = -test_values(arr_idx) - 1
'    'next
'
'    test_values(0) = &H80010003
'    test_values(1) = &H80010002
'    test_values(2) = &HE0010003
'    test_values(3) = &HE0010002
'    test_values(4) = &H40010003
'    test_values(5) = &H40010002
'    test_values(6) = &H70010003
'    test_values(7) = &H70010002
'
'    arr_idx = 0 ' 0 ～ 7 の範囲で変更
'
'    Dim test_value As Long
'    test_value = test_values(arr_idx)
'
'    'Debug.Print "----------------------"
'    'Debug.Print "left"
'    'Debug.Print test_value & "(" & LongToBin(test_value) & ")"
'
'    Dim idx As Long
'    For idx = 33 To -33 Step -1
'        'Debug.Print LongToBin(BitLeft(test_value, idx)) & " | " & idx
'    Next
'    'Debug.Print "right (arithmetic)"
'    For idx = 0 To 33
'        'Debug.Print LongToBin(BitRight(test_value, idx, Arithmetic:=True)) & " | " & idx
'    Next
'End Sub

'* ビット左シフト演算を行う関数。
'*
'* @param TargetValue 左シフトする対象の値。
'* @param ShiftCount シフトするビット数。
'* @return 左シフトされた結果の値。
'*
'* @details
'* 引数として指定された Long 値に対して、左ビットシフト演算を行います。
'* 負のシフト数が指定された場合は、論理右シフトとして処理します。
Public Function BitLeft(ByVal TargetValue As Long, ByVal ShiftCount As Long) As Long
    ' シフト数 0 の場合は、そのまま返す
    If ShiftCount = 0 Then
        BitLeft = TargetValue
        Exit Function
    End If
    
    ' シフト数が 32 以上の場合は、すべて消える
    If 31 < ShiftCount Then
        BitLeft = &H0&
        Exit Function
    End If
    
    ' シフト数が負の場合は、論理右シフトとして処理
    If ShiftCount < 0 Then
        BitLeft = BitRight(TargetValue, -ShiftCount)
        Exit Function
    End If
    
    ' ShiftCount が 31 のときは桁あふれするので先に処理
    If ShiftCount = 31 Then
        ' ShiftCount が 31 のときは、TargetValue の最も右のビットが符合ビット (最も左) になる。
        ' その他は 0 になる。
        If TargetValue Mod 2 = 0 Then
            ' 偶数のときは最も右のビットは 0 なので、符号ビットは 0
            BitLeft = &H0&
        Else
            ' 奇数のときは最も右のビットは 1 なので、符号ビットは 1
            BitLeft = &H80000000
        End If
        Exit Function
    End If
    
    ' 以降、シフト数が 1 ～ 30 の計算
    
    Dim result As Long
   
   ' 桁あふれ対策のマスクを準備
   ' 他言語でいう「&HFFFFFFFF >> (ShiftCount + 1)」と同じ値を得るための処理
   ' 消えるはずの上位ビットと、処理後に符合ビットになるビットを 0 にするために使用する。
    Dim mask_value As Long
    mask_value = (2& ^ (31 - ShiftCount)) - 1
    
    ' シフト処理
    ' 他言語でいう「TargetValue << ShiftCount」を 2 の累乗を掛けることで模擬
    result = (TargetValue And mask_value) * (2 ^ ShiftCount)
    
    ' 最上位ビット取得用マスクを準備
    ' 他言語でいう「&H1& << (32 - ShiftCount)」と同じ値を得るための処理
    ' 処理後に符号ビットになるはずのビットを得るために使用する。
    Dim high_bit_mask As Long
    high_bit_mask = &H1& * (2& ^ (31 - ShiftCount))
    
    ' 符号ビット (最上位ビット) の処理
    ' 0 でないなら符号ビットを立てる。
    If (TargetValue And high_bit_mask) <> 0 Then
        result = result Or &H80000000
    End If
    
    BitLeft = result
End Function

'* ビット右シフト演算を行う関数。
'*
'* @param TargetValue 右シフトする対象の値。
'* @param ShiftCount シフトするビット数。
'* @param Arithmetic 算術シフトを行う場合はTrue。デフォルトはFalse。
'* @return 右シフトされた結果の値。
'*
'* @details
'* 引数として指定された Long 値に対して、右ビットシフト演算を行います。
'* 負のシフト数が指定された場合は、左シフトとして処理します。
'* 算術シフトが要求された場合、左側が符号ビットで埋められます。
Public Function BitRight(ByVal TargetValue As Long, ByVal ShiftCount As Long, Optional ByVal Arithmetic As Boolean = False) As Long
    ' シフト数 0 の場合は、そのまま返す
    If ShiftCount = 0 Then
        BitRight = TargetValue
        Exit Function
    End If
    
    ' シフト数が 32 以上の場合
    If 31 < ShiftCount Then
        If Arithmetic And TargetValue < 0 Then
            ' すべて符号ビットで埋められる
            BitRight = &HFFFFFFFF
        Else
            ' すべて消える
            BitRight = 0
        End If
        Exit Function
    End If
    
    ' シフト数が負の場合は、左シフトとして処理
    If ShiftCount < 0 Then
        BitRight = BitLeft(TargetValue, -ShiftCount)
        Exit Function
    End If
    
    ' ShiftCount が 31 のときは桁あふれするので先に処理
    If ShiftCount = 31 Then
        If TargetValue < 0 Then
            If Arithmetic Then
                ' 31 桁は符号ビット (1) で埋められ、32 桁目は符号ビットそのもの (1)
                BitRight = &HFFFFFFFF
            Else
                ' 31 桁は 0 で埋められ、32 桁目は符号ビットそのもの (1)
                BitRight = &H1&
            End If
        Else
            ' 31 桁は符号ビット (0) で埋められ、32 桁目は符号ビットそのもの (0)
            ' 31 桁は 0 で埋められ、32 桁目は符号ビットそのもの (0)
            ' いずれにせよ、すべて 0
            BitRight = &H0&
        End If
        Exit Function
    End If
    
    ' 以降、シフト数が 1 ～ 30 の計算
    
    Dim result As Long
    ' まずは符号ビットを倒す
    result = TargetValue And &H7FFFFFFF
    
    ' 1 個シフト
    ' 整数の範囲内で 2 で 1 回割ることで、右シフトを模擬
    result = result \ 2
    
    ' 符号ビットを 1 個シフト
    ' 符号ビットが 1 (つまり負) のときに、第 2 ビット目を立てることで模擬
    If TargetValue < 0 Then
        result = result Or &H40000000
    End If
    
    ' 残りのシフト
    ' 整数の範囲内で、2 の (ShiftCount - 1) 乗 で割ることで、右シフトを模擬
    ' (シフト数が 1 のときは無駄な処理になるが、1 以外の時に不要な比較をするのとどちらが良いか…)
    result = result \ (2 ^ (ShiftCount - 1))
    
    If Arithmetic And TargetValue < 0 Then
        ' 他言語でいう「&HFFFFFFFF << (32 - ShiftCount - 1)」と同じ値を得るための処理
        Dim sign_mask As Long
        sign_mask = Not ((2& ^ (32 - ShiftCount)) - 1)
        
        ' 論理シフトのせいで 0 になっている部分を 1 で埋める
        result = result Or sign_mask
    End If
    
    BitRight = result
End Function

'* Long 型の値を符号なし整数として比較する関数。
'*
'* @param ValueA 比較する最初の値。
'* @param ValueB 比較する2番目の値。
'* @return 値の比較結果。A = B の場合は 0、A < B の場合は -1、A > B の場合は 1。
'*
'* @details
'* 値 A と値 B を符号なし整数として比較します。
Public Function CompareAsUnsignLong(ByVal ValueA As Long, ByVal ValueB As Long) As Integer
    If ValueA = ValueB Then
        CompareAsUnsignLong = 0
    Else
        If 0 <= ValueA Then
            If 0 <= ValueB Then
                ' 両方正の時は普通に比較する。
                If ValueA < ValueB Then
                    CompareAsUnsignLong = -1
                Else
                    CompareAsUnsignLong = 1
                End If
            Else
                ' A が正、B が負のときは A < B
                CompareAsUnsignLong = -1
            End If
        Else
            If 0 <= ValueB Then
                ' A が負、B が正のときは A > B
                CompareAsUnsignLong = 1
            Else
                ' 両方負の時は普通に比較する。
                If ValueA < ValueB Then
                    CompareAsUnsignLong = -1
                Else
                    CompareAsUnsignLong = 1
                End If
            End If
        End If
    End If
End Function

'* 符号なし整数として A < B の関係を判定します。
'*
'* @param ValueA 比較する最初の値
'* @param ValueB 比較する 2 番目の値
'* @return A が B より小さい場合は True、それ以外は False
'*
'* @details
'* 2 つの Long 型の値を符号なし整数として比較し、A < B の場合に True を返します。
'* この関数の判定結果は、`CompareAsUnsignLong` 関数が負の値を返す場合と一致します。
Public Function IsLessThanUnsignLong(ByVal ValueA As Long, ByVal ValueB As Long) As Boolean
    If CompareAsUnsignLong(ValueA, ValueB) < 0 Then
        IsLessThanUnsignLong = True
    Else
        IsLessThanUnsignLong = False
    End If
End Function

'* 2 個の Long 型の値を符号なし整数として加算します。
'*
'* @param ValueA 加算対象の最初の値
'* @param ValueB 加算対象の 2 番目の値
'* @return 符号なしとして扱った加算結果
'*
'* @details
'* 2 個の Long 型の値を符号なし整数として扱い、加算を行います。
'* 加算結果が `FFFFFFFF` を超える場合はエラーとなります。
Public Function AddUnsignLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    Dim a_high As Long
    Dim a_low As Long
    Call pSeparateUnsignLong(a_high, a_low, ValueA)
    
    Dim b_high As Long
    Dim b_low As Long
    Call pSeparateUnsignLong(b_high, b_low, ValueB)
    
    Dim carry_bit_l As Long
    Dim add_result_l As Long
    Call pAddUnsignLongCore(add_result_l, carry_bit_l, a_low, b_low)
    
    Dim carry_bit_ha As Long
    Dim add_result_ha As Long
    Call pAddUnsignLongCore(add_result_ha, carry_bit_ha, a_high, carry_bit_l)
    
    If 0 < carry_bit_ha Then Err.Raise vbObjectError + 1, "Function AddUnsignLong", "加算の結果が FFFFFFFF を超えます(" & Hex(ValueA) & " + " & Hex(ValueB) & ")"
    
    Dim carry_bit_h As Long
    Dim add_result_h As Long
    Call pAddUnsignLongCore(add_result_h, carry_bit_h, add_result_ha, b_high)
    If 0 < carry_bit_h Then Err.Raise vbObjectError + 1, "Function AddUnsignLong", "加算の結果が FFFFFFFF を超えます(" & Hex(ValueA) & " + " & Hex(ValueB) & ")"
    
    Dim high_bit As Boolean
    If 32767 < add_result_h Then
        high_bit = True
        add_result_h = add_result_h And &H7FFF&
    Else
        high_bit = False
    End If
    
    add_result_h = add_result_h * 65536
    
    Dim result As Long
    result = add_result_h + add_result_l
    
    If high_bit Then
        result = result Or &H80000000
    End If
    
    AddUnsignLong = result
End Function

Private Sub pSeparateUnsignLong(ByRef HighPart As Long, ByRef LowPart As Long, ByVal TargetValue As Long)
    Dim high_bit As Boolean
    If TargetValue < 0 Then
        high_bit = True
        TargetValue = TargetValue And &H7FFFFFFF
    Else
        high_bit = False
    End If
    
    LowPart = TargetValue And &HFFFF&
    
    TargetValue = TargetValue \ 65536
    
    If high_bit Then
        HighPart = TargetValue Or &H8000&
    Else
        HighPart = TargetValue
    End If
End Sub

Private Sub pAddUnsignLongCore(ByRef ResultValue As Long, ByRef CarryBit As Long, ByVal ValueA As Long, ByVal ValueB As Long)
    Dim result As Long
    result = ValueA + ValueB
    If 65535 < result Then
        ResultValue = result - 65536
        CarryBit = 1
    Else
        ResultValue = result
        CarryBit = 0
    End If
End Sub

'* 2 個の Long 型の値を符号なし整数として減算します。
'*
'* @param ValueA 減算される値
'* @param ValueB 減算する値
'* @return 符号なしとして扱った減算結果
'*
'* @details
'* 2 個の Long 型の値を符号なし整数として扱い、減算を行います。
'* 減算結果が負になる場合 (`A < B`) はエラーとなります。
Public Function SubtractUnsignLong(ByVal ValueA As Long, ByVal ValueB As Long) As Long
    If ValueA = ValueB Then
        SubtractUnsignLong = 0
        Exit Function
    End If
    
    If IsLessThanUnsignLong(ValueA, ValueB) Then
        Err.Raise vbObjectError + 1, "Function SubtractUnsignLong", "第 1 引数より第 2 引数のほうが大きく、計算結果が負になります。(a: " & Hex(ValueA) & ", b: " & Hex(ValueB) & ")"
        Exit Function
    End If
    
    Dim result As Long
    Dim lng_1 As Long
    Dim lng_2 As Long
    
    If 0 <= ValueA Then
        If 0 <= ValueB Then
            ' 両方正の時は普通に計算する。
            result = ValueA - ValueB
        Else
            ' A が正、B が負のとき
            'Debug.Print "Function SubtractUnsignLong: 想定外 (A が正、B が負)"
        End If
    Else
        If 0 <= ValueB Then
            ' A が負、B が正のとき
            On Error Resume Next
            lng_1 = ValueA - C_LONG_MIN
            lng_2 = C_LONG_MAX - ValueB
            result = AddUnsignLong(lng_1, lng_2)
            result = AddUnsignLong(result, 1)
        Else
            ' 両方負の時は普通に計算する。
            result = ValueA - ValueB
        End If
    End If
    
    SubtractUnsignLong = CLng(result)
End Function

'* 指定された値が整数 (Integer 型) に変換可能かを判定します。
'*
'* @param Value 判定対象の値
'* @return 整数に変換可能な場合は True、それ以外は False
'*
'* @details
'* 指定された値が数値であり、`Integer` 型に変換可能な範囲内である場合に True を返します。
'* 値が数値でない場合や範囲外の場合は False を返します。
Public Function IsInteger(ByVal Value As Variant) As Boolean
    If IsNumeric(Value) Then
        If CDbl(Value) = CInt(Value) Then
            IsInteger = True
            Exit Function
        End If
    End If
    
    IsInteger = False
End Function

'* 指定された値が長整数 (Long 型) に変換可能かを判定します。
'*
'* @param Value 判定対象の値
'* @return 長整数に変換可能な場合は True、それ以外は False
'*
'* @details
'* 指定された値が数値であり、`Long` 型に変換可能な範囲内である場合に True を返します。
'* 値が数値でない場合や範囲外の場合は False を返します。
Public Function IsLong(ByVal Value As Variant) As Boolean
    If IsNumeric(Value) Then
        If CDbl(Value) = CLng(Value) Then
            IsLong = True
            Exit Function
        End If
    End If
    
    IsLong = False
End Function

' #############################################################################
'
' Excel 関連
'
' #############################################################################

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel ファイル フォーマット判定
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* ファイル パスの拡張子から、ファイルフォーマットを得ます。
'*
'* @param FileNameOrPath ファイルのパス文字列。
'* @return ファイル フォーマット。判定できなかった場合は、xlOpenXMLWorkbook となります。
'*
'* @details
'* ファイル パスの拡張子から、ファイルフォーマットを得ます。
'* 拡張子の大小文字は区別しません。
Public Function GetExcelFileFormat(ByVal FileNameOrPath As String) As Long
    Select Case LCase$(GetLeafFromPath(FileNameOrPath, BaseName:=False, Extension:=True))
     Case ".xlsm"
        GetExcelFileFormat = xlOpenXMLWorkbookMacroEnabled
     Case ".xltm"
        GetExcelFileFormat = xlOpenXMLTemplateMacroEnabled
     Case ".xls"
        GetExcelFileFormat = xlExcel8
     Case ".xla"
        GetExcelFileFormat = xlAddIn8
     Case ".xlam"
        GetExcelFileFormat = xlOpenXMLAddIn
     Case ".xlsb"
        GetExcelFileFormat = xlExcel12
     Case ".xlt"
        GetExcelFileFormat = xlTemplate8
     Case ".xltx"
        GetExcelFileFormat = xlOpenXMLTemplate
     Case ".xlw"
        GetExcelFileFormat = xlExcel4Workbook
     Case ".csv"
        GetExcelFileFormat = xlCSV
     Case ".txt"
        GetExcelFileFormat = xlCurrentPlatformText
     Case ".dbf"
        GetExcelFileFormat = xlDBF4
     Case ".dif"
        GetExcelFileFormat = xlDIF
     Case ".htm"
        GetExcelFileFormat = xlHtml
     Case ".html"
        GetExcelFileFormat = xlHtml
     Case ".ods"
        GetExcelFileFormat = xlOpenDocumentSpreadsheet
     Case ".slk"
        GetExcelFileFormat = xlSYLK
     Case Else
        GetExcelFileFormat = xlOpenXMLWorkbook
    End Select
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Excel アドレス関連
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'* Range アドレスが複数選択 (例: A1, B2:C3, D4) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 複数選択されている場合は True、それ以外は False
'*
'* @details
'* Range アドレスが複数選択 (例: A1, B2:C3, D4) かどうかをチェックします。
Public Function IsMultiRange(ByVal AddressString As String) As Boolean
    Dim folder_path As String
    Dim book_name As String
    Dim sheet_name As String
    Dim cell_address As String
    Call SplitExcelAddress(folder_path, book_name, sheet_name, cell_address, AddressString)

    IsMultiRange = 0 < InStr(cell_address, ",")
End Function

'* Range アドレスが Area (単一セルを除く連続した複数セル範囲) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return Area の場合は True、それ以外は False
'*
'* @details
'* この共通モジュールでは Area を「単一セルを除く、連続した複数セル範囲」として扱います。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsArea = range_bounds.IsArea
End Function

'* Range アドレスが Cell (単一セル範囲) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 単一セル範囲の場合は True、それ以外は False
'*
'* @details
'* Range アドレスが Cell (例: A1) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsCell(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsCell = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsCell = range_bounds.IsCell
End Function

'* Range アドレスが行全体 (例: 1:2) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 行全体である場合は True、それ以外は False
'*
'* @details
'* Range アドレスが行全体 (例: 1:2) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsEntireRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireRow = range_bounds.IsEntireRow
End Function

'* Range アドレスが列全体 (例: A:B) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 列全体である場合は True、それ以外は False
'*
'* @details
'* Range アドレスが列全体 (例: A:B) かどうかをチェックします。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsEntireColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsEntireColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsEntireColumn = range_bounds.IsEntireColumn
End Function

'* Range アドレスが 1 行形状 (例: A1 や A1:B1 や 1:1) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 行形状の場合は True、それ以外は False
'*
'* @details
'* 単一セルも 1 行形状として True になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneRow(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRow = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRow = range_bounds.IsOneRow
End Function

'* Range アドレスが 1 列形状 (例: A1 や A1:A2 や A:A) かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 列形状の場合は True、それ以外は False
'*
'* @details
'* 単一セルも 1 列形状として True になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneColumn(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumn = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumn = range_bounds.IsOneColumn
End Function

'* Range アドレスが 1 行だけで構成される Area かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 行だけで構成される Area の場合は True、それ以外は False
'*
'* @details
'* A1:B1 や 1:1 は True、A1 は Cell のため False になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneRowArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneRowArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneRowArea = range_bounds.IsOneRowArea
End Function

'* Range アドレスが 1 列だけで構成される Area かどうかをチェックします。
'*
'* @param AddressString 判定対象の Excel アドレス文字列
'* @return 1 列だけで構成される Area の場合は True、それ以外は False
'*
'* @details
'* A1:A2 や A:A は True、A1 は Cell のため False になります。
'* 非矩形の複数選択範囲は False を返します。
Public Function IsOneColumnArea(ByVal AddressString As String) As Boolean
    If IsMultiRange(AddressString) Then
        IsOneColumnArea = False
        Exit Function
    End If

    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBoundsFromAddress(AddressString)

    IsOneColumnArea = range_bounds.IsOneColumnArea
End Function

'* パラメータを指定して範囲のアドレス表記文字列を得ます。
'*
'* @param StartRow 開始行番号 (省略時は未指定)
'* @param StartColumn 開始列番号 (省略時は未指定)
'* @param FinishRow 終了行番号 (省略時は未指定)
'* @param FinishColumn 終了列番号 (省略時は未指定)
'* @param IsAbsoluteStartRow 開始行を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteStartColumn 開始列を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteFinishRow 終了行を絶対参照で指定する場合は True (既定値は False)
'* @param IsAbsoluteFinishColumn 終了列を絶対参照で指定する場合は True (既定値は False)
'* @param ReferenceRow 相対参照の基準行番号 (既定値は 1)
'* @param ReferenceColumn 相対参照の基準列番号 (既定値は 1)
'* @param AddressType アドレス形式 ("A1" または "R1C1") (既定値は "A1")
'* @param SheetName シート名 (省略時は未指定)
'* @param BookName ブック名 (省略時は未指定)
'* @return 指定された条件に基づいて生成された範囲のアドレス文字列
'*
'* @details
'* 行、列、絶対参照・相対参照などのパラメータを指定して、Excel の範囲アドレスを文字列で生成します。
'*
'* 行指定が省略された場合、列範囲として扱われます。列指定が省略された場合、行範囲として扱われます。
'* 行指定と列指定の両方を省略することはできません。
Public Function RangeAddress( _
        Optional ByVal StartRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal StartColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishRow As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal FinishColumn As Long = G_OMIT_CELL_INDEX, _
        Optional ByVal IsAbsoluteStartRow As Boolean = False, _
        Optional ByVal IsAbsoluteStartColumn As Boolean = False, _
        Optional ByVal IsAbsoluteFinishRow As Boolean = False, _
        Optional ByVal IsAbsoluteFinishColumn As Boolean = False, _
        Optional ByVal ReferenceRow As Long = 0, _
        Optional ByVal ReferenceColumn As Long = 0, _
        Optional ByVal AddressType As String = "A1", _
        Optional ByVal SheetName As String = "", _
        Optional ByVal BookName As String = "") As Variant
    
    ' 相対参照の参照元確認
    If (IsAbsoluteStartRow Or IsAbsoluteFinishRow) And ReferenceRow < 1 And AddressType = "A1" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="基準となる行インデックスは正である必要があります。(" & ReferenceRow & ")"
        Exit Function
    End If
    
    If (IsAbsoluteStartColumn Or IsAbsoluteFinishColumn) And ReferenceColumn < 1 And AddressType = "A1" Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="基準となる列インデックスは正である必要があります。(" & ReferenceColumn & ")"
        Exit Function
    End If
    
    ' 行指定および列指定の確認
    If StartRow = G_OMIT_CELL_INDEX And StartColumn = G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始行と開始列の両方を省略することはできません。"
        Exit Function
    End If
    
    If StartRow = G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始行のみを省略することはできません。(finish_row: " & FinishRow & ")"
        Exit Function
    ElseIf StartRow <> G_OMIT_CELL_INDEX And FinishRow = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartRow = IsAbsoluteFinishRow Then
            ' 絶対アドレス指定が一致していたら、FinishRow を補完する。
            FinishRow = StartRow
        ElseIf Not IsAbsoluteFinishRow Then
            ' IsAbsoluteFinishRow が初期値 (False) のままなら、FinishRow と IsAbsoluteFinishRow を補完する。
            FinishRow = StartRow
            IsAbsoluteFinishRow = IsAbsoluteStartRow
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="終了行を省略することはできません。(start_row: " & StartRow & ")"
            Exit Function
        End If
    End If
    
    If StartColumn = G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX Then
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="開始列のみを省略することはできません。(finish_col: " & FinishColumn & ")"
        Exit Function
    ElseIf StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = G_OMIT_CELL_INDEX Then
        If IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
            ' 絶対アドレス指定が一致していたら、FinishColumn を補完する。
            FinishColumn = StartColumn
        ElseIf Not IsAbsoluteFinishColumn Then
            ' IsAbsoluteFinishColumn が初期値 (False) のままなら、FinishColumn と IsAbsoluteFinishColumn を補完する。
            FinishColumn = StartColumn
            IsAbsoluteFinishColumn = IsAbsoluteStartColumn
        Else
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="終了列を省略することはできません。(start_col: " & StartColumn & ")"
            Exit Function
        End If
    End If
    
    Dim is_row_range As Boolean
    If StartColumn = G_OMIT_CELL_INDEX Then
        is_row_range = True
    Else
        is_row_range = False
    End If
    
    Dim is_col_range As Boolean
    If StartRow = G_OMIT_CELL_INDEX Then
        is_col_range = True
    Else
        is_col_range = False
    End If
    
    Dim is_cell_range As Boolean
    If StartRow <> G_OMIT_CELL_INDEX And FinishRow = StartRow And IsAbsoluteStartRow = IsAbsoluteFinishRow _
            And StartColumn <> G_OMIT_CELL_INDEX And FinishColumn = StartColumn And IsAbsoluteStartColumn = IsAbsoluteFinishColumn Then
        is_cell_range = True
    Else
        is_cell_range = False
    End If
    
    Dim result As String
    result = ExcelBookAndSheetAddress(BookName, SheetName)
    
    Select Case AddressType
     Case "A1"
        If is_cell_range Then
            result = result & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow)
        Else
            If is_row_range Then
                result = result & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & ":" & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            ElseIf is_col_range Then
                result = result & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteStartColumn, ReferenceColumn)
            Else
                result = result & pA1columnAddressCore(StartColumn, IsAbsoluteStartColumn, ReferenceColumn) & pA1RowAddressCore(StartRow, IsAbsoluteStartRow, ReferenceRow) & _
                        ":" & pA1columnAddressCore(FinishColumn, IsAbsoluteFinishColumn, ReferenceColumn) & pA1RowAddressCore(FinishRow, IsAbsoluteFinishRow, ReferenceRow)
            End If
        End If
     Case "R1C1"
        If is_cell_range Then
            result = result & "R" & pR1C1AddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1AddressCore(StartColumn, IsAbsoluteStartColumn)
        Else
            If is_row_range Then
                result = result & "R" & pR1C1AddressCore(StartRow, IsAbsoluteStartRow) & "C:R" & pR1C1AddressCore(FinishRow, IsAbsoluteFinishRow) & "C"
            ElseIf is_col_range Then
                result = result & "RC" & pR1C1AddressCore(StartColumn, IsAbsoluteStartColumn) & ":RC" & pR1C1AddressCore(FinishColumn, IsAbsoluteFinishColumn)
            Else
                result = result & "R" & pR1C1AddressCore(StartRow, IsAbsoluteStartRow) & "C" & pR1C1AddressCore(StartColumn, IsAbsoluteStartColumn) & _
                        ":" & "R" & pR1C1AddressCore(FinishRow, IsAbsoluteFinishRow) & "C" & pR1C1AddressCore(FinishColumn, IsAbsoluteFinishColumn)
            End If
        End If
     Case Else
        Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="AddressType は A1 か R1C1 です。(" & AddressType & ")"
        Exit Function
    End Select
    
    RangeAddress = result
End Function

Private Function pA1columnAddressCore(ByVal ColumnIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceColumn As Long) As String
    If IsAbsolute Then
        If ColumnIndex < 1 Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & ColumnIndex & ")"
        End If
        
        pA1columnAddressCore = "$" & ExcelA1ColumnAddress(ColumnIndex)
    Else
        If ReferenceColumn + ColumnIndex < 1 Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & ReferenceColumn & " + " & ColumnIndex & " = " & (ReferenceColumn + ColumnIndex) & ")"
        End If
        
        pA1columnAddressCore = ExcelA1ColumnAddress(ReferenceColumn + ColumnIndex)
    End If
End Function

Private Function pA1RowAddressCore(ByVal RowIndex As Long, ByVal IsAbsolute As Boolean, ByVal ReferenceRow As Long) As String
    If IsAbsolute Then
        If RowIndex < 1 Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="列インデックスが範囲外です。(" & RowIndex & ")"
        End If
        
        pA1RowAddressCore = "$" & CStr(RowIndex)
    Else
        If ReferenceRow + RowIndex < 1 Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="行インデックスが範囲外です。(" & ReferenceRow & " + " & RowIndex & " = " & (ReferenceRow + RowIndex) & ")"
        End If
        
        pA1RowAddressCore = CStr(ReferenceRow + RowIndex)
    End If
End Function

Private Function pR1C1AddressCore(ByVal IndexNumber As Long, ByVal IsAbsolute As Boolean) As String
    If IsAbsolute Then
        If IndexNumber < 1 Then
            Err.Raise Number:=vbObjectError + 1, Source:="Function RangeAddress", Description:="インデックスが範囲外です。(" & IndexNumber & ")"
        End If
        
        pR1C1AddressCore = CStr(IndexNumber)
    Else
        If IndexNumber <> 0 Then
            pR1C1AddressCore = "[" & CStr(IndexNumber) & "]"
        Else
            pR1C1AddressCore = ""
        End If
    End If
End Function

'* ブック名とシート名を Excel アドレス形式にした文字列を得ます。
'*
'* @param BookName ブック名 (省略時は空文字列)
'* @param SheetName シート名 (省略時は空文字列)
'* @return 指定されたブック名およびシート名を含む Excel アドレス形式の文字列
'*
'* @details
'* ブック名とシート名を組み合わせて、Excel で使用可能なアドレス形式の文字列を生成します。
Public Function ExcelBookAndSheetAddress(Optional ByVal BookName As String = "", Optional ByVal SheetName As String = "") As String
    If pIsNeedQuoteBookAndSheetAddressCore(BookName, SheetName) Then
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "'[" & pEscapeBookAndSheetAddressCore(BookName) & "]" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = "'" & pEscapeBookAndSheetAddressCore(SheetName) & "'!"
        Else
            'Debug.Print "Function ExcelBookAndSheetAddress: 想定外"
        End If
    Else
        If BookName <> "" Then
            ExcelBookAndSheetAddress = "[" & BookName & "]" & SheetName & "!"
        ElseIf SheetName <> "" Then
            ExcelBookAndSheetAddress = SheetName & "!"
        Else
            ExcelBookAndSheetAddress = ""
        End If
    End If
End Function

Private Function pIsNeedQuoteBookAndSheetAddressCore(ByRef BookName As String, ByRef SheetName As String) As String
    If BookName = "" And SheetName = "" Then
        pIsNeedQuoteBookAndSheetAddressCore = False
    ElseIf 0 < InStr(BookName, " ") Or 0 < InStr(BookName, "'") Or 0 < InStr(BookName, "!") Or 0 < InStr(BookName, "[") Or 0 < InStr(BookName, "]") Or 0 < InStr(BookName, "(") Or 0 < InStr(BookName, ")") _
            Or 0 < InStr(SheetName, " ") Or 0 < InStr(SheetName, "'") Or 0 < InStr(SheetName, "!") Then
        pIsNeedQuoteBookAndSheetAddressCore = True
    Else
        pIsNeedQuoteBookAndSheetAddressCore = False
    End If
End Function

Private Function pEscapeBookAndSheetAddressCore(ByVal BookOrSheetName As String) As String
    pEscapeBookAndSheetAddressCore = Replace(Replace(Replace(BookOrSheetName, "'", "''"), "[", "("), "]", ")")
End Function

'* 列番号を A1 形式の列名 (例: A, B, ... Z, AA, AB...) に変換します。
'*
'* @param ColumnIndex 列番号 (1 以上の値)
'* @return 列番号に対応する A1 形式の列名
'*
'* @details
'* 指定された列番号を基に、Excel の A1 形式の列名を生成します。
'* 列番号が負または 0 の場合、エラーになります。
Public Function ExcelA1ColumnAddress(ByVal ColumnIndex As Long) As String
    Dim result As String
    Dim temp_num As Long
    
    Do While ColumnIndex > 0
        temp_num = (ColumnIndex - 1) Mod 26
        result = Chr(temp_num + 65) & result
        ColumnIndex = (ColumnIndex - 1) \ 26
    Loop
    
    ExcelA1ColumnAddress = result
End Function

'* Excel のアドレス表記から各情報を取り出します。
'*
'* @param FolderPath [出力] 結果として取得されるフォルダパス
'* @param BookName [出力] 結果として取得されるブック名
'* @param SheetName [出力] 結果として取得されるシート名
'* @param CellAddress [出力] 結果として取得されるセルアドレス
'* @param AddressString 分解対象となる Excel アドレス文字列
'*
'* @details
'* 指定された Excel アドレス文字列を分解し、フォルダパス、ブック名、シート名、セルアドレスに分けて出力します。
'* 形式が正しくない場合はエラーにします。
Public Sub SplitExcelAddress(ByRef FolderPath As String, ByRef BookName As String, ByRef SheetName As String, ByRef CellAddress As String, ByVal AddressString As String)
    FolderPath = ""
    BookName = ""
    SheetName = ""
    CellAddress = ""

    If AddressString = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim addr_parts() As String
    addr_parts = Split(AddressString, "!")
    If 1 < UBound(addr_parts) Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If

    Dim location_part As String
    If UBound(addr_parts) = 0 Then
        CellAddress = addr_parts(0)
        If 0 < InStr(CellAddress, "[") Or 0 < InStr(CellAddress, "]") Or 0 < InStr(CellAddress, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If
    Else
        location_part = addr_parts(0)
        CellAddress = addr_parts(1)

        If location_part = "" Or CellAddress = "" Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        If StartsWith(location_part, "'") Then
            If Not EndsWith(location_part, "'") Or Len(location_part) < 2 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            location_part = Mid(location_part, 2, Len(location_part) - 2)
            location_part = Replace(location_part, "''", "'")
        ElseIf 0 < InStr(location_part, "'") Then
            Call pRaiseInvalidExcelAddress(AddressString)
        End If

        Dim close_book_idx As Long
        close_book_idx = InStrRev(location_part, "]")
        If 0 < close_book_idx Then
            Dim open_book_idx As Long
            open_book_idx = InStrRev(Left(location_part, close_book_idx - 1), "[")
            If open_book_idx < 1 Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If

            FolderPath = Left(location_part, open_book_idx - 1)
            BookName = Mid(location_part, open_book_idx + 1, close_book_idx - open_book_idx - 1)
            SheetName = Mid(location_part, close_book_idx + 1)
            If BookName = "" Or SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        Else
            If 0 < InStr(location_part, "[") Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
            SheetName = location_part
            If SheetName = "" Then
                Call pRaiseInvalidExcelAddress(AddressString)
            End If
        End If
    End If

    If CellAddress = "" Then
        Call pRaiseInvalidExcelAddress(AddressString)
    End If
End Sub

Private Sub pRaiseInvalidExcelAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitExcelAddress", Description:="Excel アドレス文字列の形式が正しくありません。(" & AddressString & ")"
End Sub

'* A1 形式の単一矩形範囲アドレスを、開始・終了インデックスへ分解します。
'*
'* @param StartRow [出力] 開始行番号。列範囲の場合は G_OMIT_CELL_INDEX。
'* @param StartColumn [出力] 開始列番号。行範囲の場合は G_OMIT_CELL_INDEX。
'* @param FinishRow [出力] 終了行番号。列範囲の場合は G_OMIT_CELL_INDEX。
'* @param FinishColumn [出力] 終了列番号。行範囲の場合は G_OMIT_CELL_INDEX。
'* @param AddressString 分解対象の A1 形式アドレス。ブック名・シート名は含めない。
'*
'* @details
'* A1、A1:B2、1:3、A:C、$A$1:$B$2 を扱います。
'* 複数範囲、R1C1 形式、ブック名・シート名付きアドレス、不完全なアドレスはエラーにします。
Public Sub SplitA1RangeAddress( _
        ByRef StartRow As Long, _
        ByRef StartColumn As Long, _
        ByRef FinishRow As Long, _
        ByRef FinishColumn As Long, _
        ByVal AddressString As String)

    StartRow = G_OMIT_CELL_INDEX
    StartColumn = G_OMIT_CELL_INDEX
    FinishRow = G_OMIT_CELL_INDEX
    FinishColumn = G_OMIT_CELL_INDEX

    Dim normalized_address As String
    normalized_address = Trim(AddressString)
    If normalized_address = "" Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If 0 < InStr(normalized_address, ",") Or 0 < InStr(normalized_address, "!") _
            Or 0 < InStr(normalized_address, "[") Or 0 < InStr(normalized_address, "]") _
            Or 0 < InStr(normalized_address, "'") Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    Dim address_parts() As String
    address_parts = Split(normalized_address, ":")

    Dim start_type As Long
    Dim finish_type As Long
    If UBound(address_parts) = 0 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        If start_type <> C_A1_TOKEN_CELL Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If
        FinishRow = StartRow
        FinishColumn = StartColumn
    ElseIf UBound(address_parts) = 1 Then
        Call pSplitA1AddressToken(StartRow, StartColumn, start_type, address_parts(0), AddressString)
        Call pSplitA1AddressToken(FinishRow, FinishColumn, finish_type, address_parts(1), AddressString)
        If start_type <> finish_type Then
            Call pRaiseInvalidA1RangeAddress(AddressString)
        End If

        If start_type = C_A1_TOKEN_ROW Then
            StartColumn = G_OMIT_CELL_INDEX
            FinishColumn = G_OMIT_CELL_INDEX
        ElseIf start_type = C_A1_TOKEN_COLUMN Then
            StartRow = G_OMIT_CELL_INDEX
            FinishRow = G_OMIT_CELL_INDEX
        End If
    Else
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If

    If StartRow <> G_OMIT_CELL_INDEX And FinishRow <> G_OMIT_CELL_INDEX And FinishRow < StartRow Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
    If StartColumn <> G_OMIT_CELL_INDEX And FinishColumn <> G_OMIT_CELL_INDEX And FinishColumn < StartColumn Then
        Call pRaiseInvalidA1RangeAddress(AddressString)
    End If
End Sub

Private Sub pSplitA1AddressToken( _
        ByRef RowIndex As Long, _
        ByRef ColumnIndex As Long, _
        ByRef TokenType As Long, _
        ByVal AddressToken As String, _
        ByVal OriginalAddressString As String)

    RowIndex = G_OMIT_CELL_INDEX
    ColumnIndex = G_OMIT_CELL_INDEX
    TokenType = 0

    Dim normalized_token As String
    normalized_token = UCase(Replace(AddressToken, "$", ""))
    If normalized_token = "" Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim col_text As String
    Dim row_text As String
    Dim found_digit As Boolean
    Dim char_idx As Long
    For char_idx = 1 To Len(normalized_token)
        Dim char_code As Long
        char_code = Asc(Mid(normalized_token, char_idx, 1))

        If Asc("A") <= char_code And char_code <= Asc("Z") Then
            If found_digit Then
                Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
            End If
            col_text = col_text & Chr(char_code)
        ElseIf Asc("0") <= char_code And char_code <= Asc("9") Then
            found_digit = True
            row_text = row_text & Chr(char_code)
        Else
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    Next

    If col_text <> "" Then
        ColumnIndex = pA1ColumnIndex(col_text, OriginalAddressString)
    End If
    If row_text <> "" Then
        RowIndex = pA1RowIndex(row_text, OriginalAddressString)
    End If

    If col_text <> "" And row_text <> "" Then
        TokenType = C_A1_TOKEN_CELL
    ElseIf row_text <> "" Then
        TokenType = C_A1_TOKEN_ROW
    ElseIf col_text <> "" Then
        TokenType = C_A1_TOKEN_COLUMN
    Else
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
End Sub

Private Function pA1ColumnIndex(ByVal ColumnAddress As String, ByVal OriginalAddressString As String) As Long
    Dim result As Long
    Dim char_idx As Long
    For char_idx = 1 To Len(ColumnAddress)
        Dim char_code As Long
        char_code = Asc(Mid(ColumnAddress, char_idx, 1))
        result = result * 26 + char_code - Asc("A") + 1
        If G_COL_MAX < result Then
            Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
        End If
    Next

    If result < 1 Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1ColumnIndex = result
End Function

Private Function pA1RowIndex(ByVal RowAddress As String, ByVal OriginalAddressString As String) As Long
    If Len(CStr(G_ROW_MAX)) < Len(RowAddress) Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If

    Dim result As Long
    result = CLng(RowAddress)
    If result < 1 Or G_ROW_MAX < result Then
        Call pRaiseInvalidA1RangeAddress(OriginalAddressString)
    End If
    pA1RowIndex = result
End Function

Private Sub pRaiseInvalidA1RangeAddress(ByVal AddressString As String)
    Err.Raise Number:=vbObjectError + 1, Source:="Sub SplitA1RangeAddress", Description:="A1 形式の単一矩形範囲アドレスではありません。(" & AddressString & ")"
End Sub

'* Range の Text プロパティを格納した ObjectList を返します。
'*
'* @param TargetRange 対象となる範囲
'* @param IgnoreEmpty 空のセルを無視する場合は True (既定値は False)
'* @return 対象範囲内のセルの Value プロパティを文字列化して格納した ObjectList
'*
'* @details
'* 指定された範囲の各セルの Value プロパティを文字列化して ObjectList に格納し、返します。
'* IgnoreEmpty が True の場合、空のセルはリストに含まれません。
Public Function ConvertRangeToStringList(ByVal TargetRange As WorksheetRangeBounds, Optional ByVal IgnoreEmpty As Boolean = False) As ObjectList
    Dim result As ObjectList
    Set result = New ObjectList

    Dim enum_obj As IEnumerator
    Set enum_obj = TargetRange.GetEnumerator(EnumerateType:="Cells")
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current

        Dim cell_text As String
        Call WsSrv.ReadCell(cell_item, cell_text)

        If Not IgnoreEmpty Or cell_text <> "" Then
            Call result.Add(cell_text)
        End If
    Loop

    Set ConvertRangeToStringList = result
End Function


