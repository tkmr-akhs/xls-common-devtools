VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserInputSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ユーザー インターフェイス用シートにアクセスするためのクラスです。
'!
' #############################################################################

Private pIsInitialized As Boolean
Private pRangeBounds As WorksheetRangeBounds

'* WorksheetService のオブジェクト。ユニット テストのときにはテスト ダブルに置き換えます。
Public WsSrv As IWorksheetService

'* WorkbookService のオブジェクト。ユニット テストのときにはテスト ダブルに置き換えます。
Public WbSrv As IWorkbookService

'* ワークシート名
Public Property Get WorksheetName() As String
    Call pCheckInit
    WorksheetName = pRangeBounds.WorksheetName
End Property

'* ワークブック名
Public Property Get WorkbookName() As String
    Call pCheckInit
    WorkbookName = pRangeBounds.WorkbookName
End Property

Private Sub Class_Initialize()
    If WsSrv Is Nothing Then Set WsSrv = New WorksheetService
    If WbSrv Is Nothing Then Set WbSrv = New WorkbookService
End Sub

Private Sub pCheckInit()
    If Not pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet", "未初期化状態です。"
End Sub

'* 各プロパティを設定します。
'*
'* @param Sheet UI シートの名前。
'* @param Book UI シートがあるワークブックの名前。省略した場合は ThisWorkbook の名前となります。
'*
'* @details
'* 各プロパティを設定します。
Public Sub Initialize(ByVal Sheet As String, Optional ByVal Book As String = "")
    If pIsInitialized Then Err.Raise vbObjectError + 1, "Class UserInputSheet", "すでに初期化済みです。"
    
    Dim book_name As String
    
    If Book = "" Then
        book_name = WbSrv.GetThisWorkbookName()
    Else
        book_name = Book
    End If
    
    If Not WbSrv.ExistsWorksheet(Sheet, Book:=book_name) Then Err.Raise vbObjectError + 1, "Class UserInputSheet", "対象シートが見つかりません。(" & Sheet & "@" & book_name & ")"
    
    Set pRangeBounds = New_RangeBounds(Sheet:=Sheet, Book:=book_name)
    pIsInitialized = True
End Sub

'* 指定された項目名に対応する項目値の範囲を取得します。
'*
'* @param FirstItemName 第一項目名
'* @param SecondItemName 第二項目名 (省略可能)
'* @return 対応する項目値の範囲 (Range オブジェクト)
'*
'* @details
'* 第一項目名の位置を基準に、行または列の範囲を取得します。
'* 第二項目名が省略されている場合、第一項目名の行全体の範囲を返します。
'* 第二項目名が指定されている場合、その列に対応する範囲を返します。
Public Function GetItemRange( _
        ByVal FirstItemName As String, _
        Optional ByVal SecondItemName As String = "") As WorksheetRangeBounds
    
    Call pCheckInit
    
    Dim used_range As WorksheetRangeBounds
    Set used_range = WsSrv.GetUsedRangeBounds(pRangeBounds)
    
    ' 1 列目を第一項目名の列とし、残りと分割する
    Dim first_col As WorksheetRangeBounds
    Set first_col = used_range.TransformAbsolute(FinishColumn:=used_range.Column)
    
    Dim remain_range As WorksheetRangeBounds
    Set remain_range = used_range.TransformAbsolute(Column:=used_range.Column + 1)
    
    ' 第一項目名の列を名前で検索する
    Dim found_cells() As WorksheetRangeBounds
    found_cells = WsSrv.Find(FirstItemName, first_col)
    
    ' もし空の配列が返ってきたら、見つからなかったということなので Nothingで終了
    If IsEmptyArray(found_cells) Then
        Set GetItemRange = Nothing
        Exit Function
    End If
    
    ' 見つかった第一項目の右横の行
    Dim second_row As WorksheetRangeBounds
    Set second_row = remain_range.TransformAbsolute(Row:=found_cells(0).Row, FinishRow:=found_cells(0).Row)
    
    ' 切り詰める
    Set second_row = WsSrv.GetUsedRangeBounds(second_row)
        
    ' 第二項目名が指定されていないときは、second_row は値
    If SecondItemName = "" Then
        If second_row.IsEmpty Then
            ' 空範囲の時は、1 セルにして返す。
            Set GetItemRange = second_row.TransformAbsolute(FinishRow:=second_row.Row, FinishColumn:=second_row.Column)
        Else
            Set GetItemRange = second_row
        End If
        Exit Function
    End If
    
    ' 第二項目名を探す。
    found_cells = WsSrv.Find(SecondItemName, second_row)
    
    ' もし空の配列が返ってきたら、見つからなかったということなので Nothing で終了
    If IsEmptyArray(found_cells) Then
        Set GetItemRange = Nothing
        Exit Function
    End If
    
    ' 次の第一項目を見つけ、第二項目の値の行数を算出
    Dim first_remain_col As WorksheetRangeBounds
    Set first_remain_col = first_col.TransformAbsolute(Row:=found_cells(0).Row + 1)
    Dim finish_row_idx As Long
    finish_row_idx = found_cells(0).Row
    Dim enum_obj As IEnumerator
    Set enum_obj = first_remain_col.GetEnumerator()
    Do While enum_obj.MoveNext()
        Dim cell_item As WorksheetRangeBounds
        Set cell_item = enum_obj.Current
        If WsSrv.IsEmptyCell(cell_item, IgnoreEmptyString:=True) Then
            finish_row_idx = finish_row_idx + 1
        Else
            Exit Do
        End If
    Loop
    finish_row_idx = MinLng(finish_row_idx, used_range.FinishRow)
    
    ' 結果の範囲を生成
    Dim result As WorksheetRangeBounds
    Set result = remain_range.TransformAbsolute(Row:=found_cells(0).Row + 1, Column:=found_cells(0).Column, FinishRow:=finish_row_idx, FinishColumn:=found_cells(0).Column)
    Set result = WsSrv.GetUsedRangeBounds(result, IgnoreEmpty:=True)
    Set GetItemRange = result
End Function

