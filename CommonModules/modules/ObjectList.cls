VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private Const C_CHUNK_SIZE As Integer = 16

' #############################################################################
'!
'! @brief
'! ダイナミックな型とサイズのオブジェクトリストを管理するクラスです。
'! このクラスは、オブジェクトおよびプリミティブ型の動的配列を提供し、
'! 型の一貫性、オブジェクト操作、および拡張機能をサポートします。
'!
' #############################################################################

Private pDepth As Integer
Private pRootArray As ArrayObject
Private pCount As Long
Private pTypeName As String
Private pIsObject As Boolean
Private pIsComparable As Boolean
Private pIsStringable As Boolean
Private pIsEquatable As Boolean
Private pIsDuplicateCheckable As Boolean

'* リストの現在の件数。
Public Property Get Count() As Long
    Count = pCount
End Property

'* リスト内の要素の型。
Public Property Get ItemTypeName() As String
    ItemTypeName = pTypeName
End Property

'* 指定されたインデックスの要素を返します。
'*
'* @param ItemIndex 要素のインデックス。
'* @return 指定されたインデックスにある要素。
'*
'* @details
'* 指定されたインデックスの要素を返します。
Public Function Item(ByVal ItemIndex As Long) As Variant
    Call pCheckIndex(ItemIndex)

    Dim index_arr() As Integer
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, ItemIndex)

    Dim target_arr As ArrayObject
    Set target_arr = pRootArray

    Dim idx As Integer
    For idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = target_arr.Item(index_arr(idx))
    Next idx

    If pIsObject Then
        Set Item = target_arr.Item(index_arr(UBound(index_arr)))
    Else
        Item = target_arr.Item(index_arr(UBound(index_arr)))
    End If
End Function

Private Sub Class_Initialize()
    Call pInitializeCore
End Sub

Private Sub pInitializeCore()
    pDepth = 0
    pCount = 0
    pTypeName = ""
    pIsObject = False
    pIsComparable = False
    pIsEquatable = False
    pIsDuplicateCheckable = False
    pIsStringable = False
    Set pRootArray = New ArrayObject
    Call pRootArray.ReDimArray(0, C_CHUNK_SIZE - 1)
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Debug.Print "Terminate: ObjectList(0x" & Hex(ObjPtr(Me)) & ")"

    Dim idx As Long
    For idx = 0 To pRootArray.UpperBound - 1
        Call pRootArray.Update(idx, Empty)
    Next idx
End Sub

'* 指定されたインデックスの要素を更新します。
'*
'* @param ItemIndex 更新する要素のインデックス。
'* @param ItemObject 新しい要素。
'*
'* @details
'* 指定されたインデックスの要素を更新します。
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call pCheckIndex(ItemIndex)
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Call pUpdateCore(ItemIndex, ItemObject)
End Sub

Private Sub pUpdateCore(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Dim index_arr() As Integer
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, ItemIndex)

    Dim target_arr As ArrayObject
    Set target_arr = pRootArray

    Dim idx As Integer
    For idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = target_arr.Item(index_arr(idx))
    Next idx

    Call target_arr.Update(index_arr(UBound(index_arr)), ItemObject)
End Sub

'* リストを配列として返します。
'*
'* @return 要素の配列。
'*
'* @details
'* リストを配列として返します。
Public Function ConvertToArray() As Variant()
    If pCount = 0 Then
        Exit Function
    End If

    Dim result() As Variant
    Dim idx As Long

    ReDim result(0 To pCount - 1)

    For idx = LBound(result) To UBound(result)
        If pIsObject Then
            Set result(idx) = Item(idx)
        Else
            result(idx) = Item(idx)
        End If
    Next idx

    ConvertToArray = result
End Function

'* リストを文字列の配列として返します。
'*
'* @return 文字列の配列。
'*
'* @details
'* リストを文字列の配列として返します。
Public Function ConvertToStringArray() As String()
    If pCount = 0 Then
        Exit Function
    End If

    Dim result() As String
    Dim idx As Long
    Dim stringable_obj As IStringable

    ReDim result(0 To pCount - 1)

    For idx = LBound(result) To UBound(result)
        If pIsObject Then
            If Item(idx) Is Nothing Then
                result(idx) = "Nothing"
            Else
                If TypeOf Item(idx) Is IStringable Then
                    Set stringable_obj = Item(idx)
                    result(idx) = stringable_obj.ToString()
                Else
                    result(idx) = TypeName(Item(idx))
                End If
            End If
        Else
            result(idx) = Item(idx)
        End If
    Next idx

    ConvertToStringArray = result
End Function

'* 新しい要素をリストに追加します。
'*
'* @param ItemObject 追加する要素。
'*
'* @details
'* 新しい要素をリストに追加します。
Public Sub Add(ByVal ItemObject As Variant)
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If

    Dim new_arr As ArrayObject
    Dim index_arr() As Integer
    Dim remain_index As Long
    Call pGetIndexArray(index_arr, remain_index, pCount)
    Call pExpandIfNeed(remain_index)

    Dim target_arr As ArrayObject
    If 0 < remain_index Then
        Set target_arr = pGetChildArray(pRootArray, remain_index)
    Else
        Set target_arr = pRootArray
    End If

    Dim idx As Integer
    For idx = LBound(index_arr) To UBound(index_arr) - 1
        Set target_arr = pGetChildArray(target_arr, index_arr(idx))
    Next idx

    Call target_arr.Update(index_arr(UBound(index_arr)), ItemObject)

    pCount = pCount + 1
End Sub

'* 配列の各要素をリストに追加します。
'*
'* @param ItemArray 追加する要素の配列。
'*
'* @details
'* 配列の各要素をリストに追加します。
Public Sub AddArray(ByVal ItemArray As Variant)
    If Not IsArray(ItemArray) Then Err.Raise vbObjectError + 1, "Class ObjectList", "引数が配列ではありません"
    If IsEmptyArray(ItemArray) Then Exit Sub

    Dim item_obj As Variant

    For Each item_obj In ItemArray
        Call Add(item_obj)
    Next item_obj
End Sub

'* 他のリストの要素をこのリストに取り込みます。
'*
'* @param OtherList 取り込む要素を持つリスト。
'*
'* @details
'* 他のリストの要素をこのリストに取り込みます。
Public Sub AddOther(ByVal OtherList As ObjectList)
    Dim idx As Long

    For idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(idx))
    Next idx
End Sub

'* 他のセットの要素をこのリストに取り込みます。
'*
'* @param OtherSet 取り込む要素を持つセット。
'*
'* @details
'* 他のセットの要素をこのリストに取り込みます。
Public Sub AddSet(ByVal OtherSet As ObjectSet)
    Dim idx As Long

    For idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(idx))
    Next idx
End Sub

'* このリストの複製を生成します。
'*
'* @return 複製された新しいリスト。
'*
'* @details
'* このリストの複製を生成します。
'* リストの要素はもとのインスタンスがそのまま利用されます。(シャロー コピー)
Public Function CopyList() As ObjectList
    Dim idx As Long
    Dim new_list As ObjectList

    Set new_list = New ObjectList

    For idx = 0 To pCount - 1
        Call new_list.Add(Item(idx))
    Next idx

    Set CopyList = new_list
End Function

'* 指定されたインデックスの要素を削除します。
'*
'* @param ItemIndex 削除する要素のインデックス。
'* @return 削除された要素。
'*
'* @details
'* 指定されたインデックスの要素を削除します。
Public Function Remove(ByVal ItemIndex As Long) As Variant
    Dim idx As Long
    Dim old_obj As Variant

    Call pCheckIndex(ItemIndex)

    If pIsObject Then
        Set old_obj = Item(ItemIndex)
    Else
        old_obj = Item(ItemIndex)
    End If

    For idx = ItemIndex To pCount - 2
        Call pUpdateCore(idx, Item(idx + 1))
    Next idx

    Call pUpdateCore(pCount - 1, Empty)

    pCount = pCount - 1

    If pIsObject Then
        Set Remove = old_obj
    Else
        Remove = old_obj
    End If
End Function

'* このリストの内容をすべて削除します。
'*
'* @details
'* このリストの内容をすべて削除します。
Public Sub RemoveAll()
    Call pInitializeCore
End Sub

'* このリストに登録されている、引数と同じ要素を削除します。
'*
'* @param ItemObject 削除対象の要素。
'* @return 削除が成功したかどうかを示すブール値。
'*
'* @details
'* このリストに登録されている、引数と同じ要素を削除します。
'* - このリストの要素が IDuplicateCheckable であれば、IsDuplicateOf の結果に基づいて処理します。
'* - このリストの要素が IEquatable であれば、Equals の結果に基づいて処理します。
'* - IEquotable かつ IDuplicateCheckable の場合は、IDuplicateCheckable としての扱いが優先されます。
Public Function RemoveItem(ByVal ItemObject As Variant) As Boolean
    Dim idx As Long
    Dim is_removed As Boolean

    is_removed = False
    Do While idx < pCount
        If pCheckRemoveItem(Item(idx), ItemObject) Then
            Call Remove(idx)
            is_removed = True
        Else
            idx = idx + 1
        End If
    Loop

    RemoveItem = is_removed
End Function

Private Function pCheckRemoveItem(ByVal ItemObject1 As Variant, ByVal ItemObject2 As Variant) As Boolean
    pCheckRemoveItem = pItemsEqual(ItemObject1, ItemObject2)
End Function

Private Function pItemsEqual(ByVal ItemObject1 As Variant, ByVal ItemObject2 As Variant) As Boolean
    If pIsObject Then
        If Not IsObject(ItemObject1) Or Not IsObject(ItemObject2) Then Exit Function

        If ItemObject1 Is Nothing Then
            pItemsEqual = (ItemObject2 Is Nothing)
            Exit Function
        End If

        If ItemObject2 Is Nothing Then Exit Function

        If ItemObject1 Is ItemObject2 Then
            pItemsEqual = True
            Exit Function
        End If

        If pIsDuplicateCheckable Then
            Dim duplicate_item As IDuplicateCheckable
            Set duplicate_item = ItemObject1
            pItemsEqual = duplicate_item.IsDuplicateOf(ItemObject2)
        ElseIf pIsEquatable Then
            Dim equatable_item As IEquatable
            Set equatable_item = ItemObject1
            pItemsEqual = equatable_item.Equals(ItemObject2)
        End If
    Else
        If IsObject(ItemObject1) Or IsObject(ItemObject2) Then Exit Function
        pItemsEqual = (ItemObject1 = ItemObject2)
    End If
End Function

'* このリストに重複して登録されている要素を削除します。
'*
'* @return 重複削除が実行されたかどうかを示すブール値。
'*
'* @details
'* このリストに重複して登録されている要素を削除します。
Public Function RemoveDuplicate() As Boolean
    Dim idx As Long
    Dim search_idx As Long
    Dim is_removed As Boolean

    is_removed = False
    Do While idx < pCount
        Dim search_item As Variant
        If pIsObject Then
            Set search_item = Item(idx)
        Else
            search_item = Item(idx)
        End If

        search_idx = idx + 1
        Do While search_idx < pCount
            If pItemsEqual(search_item, Item(search_idx)) Then
                Call Remove(search_idx)
                is_removed = True
            Else
                search_idx = search_idx + 1
            End If
        Loop
        idx = idx + 1
    Loop

    RemoveDuplicate = is_removed
End Function

'* このリストの末尾に要素を追加します。
'*
'* @param ItemObject 追加する要素。
'*
'* @details
'* このリストの末尾に要素を追加します。Add(ItemObject) と同じです。
Public Sub PushItem(ByVal ItemObject As Variant)
    Call Me.Add(ItemObject)
End Sub

'* このリストの末尾から要素を取り出します。
'*
'* @return 取り出された要素。
'*
'* @details
'* このリストの末尾から要素を取り出します。Remove(Me.Count - 1) と同じです。
Public Function PopItem() As Variant
    If pIsObject Then
        Set PopItem = Me.Remove(Me.Count - 1)
    Else
        PopItem = Me.Remove(Me.Count - 1)
    End If
End Function

'* 等しいものを探しインデックスを返します。
'*
'* @param ItemObject 探す要素。
'* @param Optional StartIndex 探索を開始するインデックス。
'* @param Optional Reverse 逆順に探索するかどうか。
'* @return 見つかった要素のインデックス。見つからない場合は -1。
'*
'* @details
'* 等しいものを探しインデックスを返します。
Public Function GetIndexByItem(ByVal ItemObject As Variant, Optional ByVal StartIndex As Long = -1, Optional ByVal Reverse As Variant = False) As Long
    If Count <= 0 _
            Or pCount <= StartIndex And Not Reverse Then
        GetIndexByItem = -1
        Exit Function
    End If

    If StartIndex < 0 Then
        If Reverse Then
            StartIndex = pCount - 1
        Else
            StartIndex = 0
        End If
    End If

    If pCount <= StartIndex And Reverse Then
        StartIndex = pCount - 1
    End If

    Dim idx_step As Long
    Dim idx_limit As Long

    If Reverse Then
        idx_step = -1
        idx_limit = 0
    Else
        idx_step = 1
        idx_limit = pCount - 1
    End If

    GetIndexByItem = pGetIndexByItemCore(ItemObject, StartIndex, idx_limit, idx_step)
End Function

Private Function pGetIndexByItemCore(ByVal ItemObject As Variant, ByVal StartIndex As Long, ByVal IndexLimit As Long, ByVal IndexStep As Long) As Long
    Dim idx As Long
    For idx = StartIndex To IndexLimit Step IndexStep
        If pItemsEqual(Me.Item(idx), ItemObject) Then
            pGetIndexByItemCore = idx
            Exit Function
        End If
    Next idx

    pGetIndexByItemCore = -1
End Function

'* 等しいものが含まれるかを確認します。
'*
'* @param ItemObject 探す要素。
'* @return 含まれているかどうかを示すブール値。
'*
'* @details
'* 等しいものが含まれるかを確認します。
Public Function Exists(ByVal ItemObject As Variant) As Boolean
    Dim idx As Long
    idx = pGetIndexByItemCore(ItemObject, 0, pCount - 1, 1)
    If 0 <= idx Then
        Exists = True
    Else
        Exists = False
    End If
End Function

'* このリストの中身をソートします。
'*
'* @param Optional Descending 降順にソートするかどうか。
'*
'* @details
'* このリストの中身をソートします。
'* リスト内に IComparable でもなく、不等号演算で比較もできないものが含まれている場合はエラーで失敗します。
'* 失敗するまでに行われた処理はリストに適用されます。
'* IStringable は文字列として処理されます。
'* IDuplicateCheckable は、キーに基づき処理されます。
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    Call pSortCore(0, pCount - 1, Descending)
End Sub
Private Sub pSortCore(ByVal MinIndex As Long, ByVal MaxIndex As Long, ByVal Descending As Boolean)
    If MinIndex >= MaxIndex Then Exit Sub ' MinIndex が MaxIndex 以降になる場合は関数を抜ける
    Dim idx As Long
    Dim pos_lng As Long: pos_lng = MinIndex

    Call Swap(MinIndex, Int((MinIndex + MaxIndex) \ 2)) '配列中央の値を基準値にして先頭と入れ替え
    For idx = MinIndex + 1 To MaxIndex
        If pIsLessThan(idx, MinIndex, Descending) Then ' MinIndex の要素を基準値として数値を比較していく
            pos_lng = pos_lng + 1 '基準値より小さい数があるごとにpを1つ進める
            Call Swap(idx, pos_lng) 'Swap関数で入れ替えを行う。
        End If
    Next
    Call Swap(MinIndex, pos_lng) 'Swap関数で入れ替え

    Call pSortCore(MinIndex, pos_lng - 1, Descending) '基準値より小さいグループを再帰処理
    Call pSortCore(pos_lng + 1, MaxIndex, Descending) '基準値より大きいグループを再帰処理
End Sub
Private Function pIsLessThan(ByVal ItemIndex1 As Long, ByRef ItemIndex2 As Long, ByVal Descending As Boolean) As Variant
    Dim result As Boolean
    If pIsComparable Then
        Dim comparable_obj As IComparable
        Set comparable_obj = Me.Item(ItemIndex1)
        result = comparable_obj.IsLessThan(Me.Item(ItemIndex2))
    ElseIf pIsDuplicateCheckable Then
        Dim objsetitem_obj1 As IDuplicateCheckable
        Dim objsetitem_obj2 As IDuplicateCheckable
        Set objsetitem_obj1 = Me.Item(ItemIndex1)
        Set objsetitem_obj2 = Me.Item(ItemIndex2)
        result = objsetitem_obj1.GetKey() < objsetitem_obj2.GetKey()
    ElseIf pIsStringable Then
        Dim stringable_obj1 As IStringable
        Dim stringable_obj2 As IStringable
        Set stringable_obj1 = Me.Item(ItemIndex1)
        Set stringable_obj2 = Me.Item(ItemIndex2)
        result = stringable_obj1.ToString() < stringable_obj2.ToString()
    Else
        result = Me.Item(ItemIndex1) < Me.Item(ItemIndex2)
    End If

    If Descending Then
        pIsLessThan = Not result
    Else
        pIsLessThan = result
    End If
End Function

'* 指定したインデックスの要素を入れ替えます。
'*
'* @param ItemIndex1 入れ替える要素の一つ目のインデックス。
'* @param ItemIndex2 入れ替える要素の二つ目のインデックス。
'*
'* @details
'* 指定したインデックスの要素を入れ替えます。
Public Sub Swap(ByVal ItemIndex1 As Long, ByRef ItemIndex2 As Long)
    Dim tmp_item As Variant
    If pIsObject Then
        Set tmp_item = Me.Item(ItemIndex1)
    Else
        tmp_item = Me.Item(ItemIndex1)
    End If

    Call Update(ItemIndex1, Me.Item(ItemIndex2))
    Call Update(ItemIndex2, tmp_item)
End Sub

'* このリストの列挙子を得ます。
'*
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 列挙子オブジェクト。
'*
'* @details
'* このリストの列挙子を得ます。
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result As Enumerator
    Set result = New Enumerator
    Call result.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result
End Function

Private Sub pCheckIndex(ByVal ItemIndex As Long)
    If pCount = 0 Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ObjectList", Description:="インデックスが範囲外です。(index: " & ItemIndex & ", count: " & pCount & ")"
    End If

    If ItemIndex < 0 Or pCount <= ItemIndex Then
        Err.Raise Number:=vbObjectError + 1, Source:="Class ObjectList", Description:="インデックスが範囲外です。(index: " & ItemIndex & ", count: " & pCount & ")"
    End If
End Sub

Private Sub pGetIndexArray(ByRef IndexArray() As Integer, ByRef RemainIndex As Long, ByVal ItemIndex As Long)
    Dim result() As Integer
    ReDim result(0 To pDepth)

    Dim remain_index As Long
    remain_index = ItemIndex

    Dim idx As Integer
    For idx = pDepth To 0 Step -1
        result(idx) = remain_index Mod C_CHUNK_SIZE
        remain_index = remain_index \ C_CHUNK_SIZE
    Next idx

    IndexArray = result
    RemainIndex = remain_index
End Sub

Private Function pGetChildArray(ByVal ParentArray As ArrayObject, ByVal ChildIndex As Long) As ArrayObject
    If IsObject(ParentArray.Item(ChildIndex)) Then
        Set pGetChildArray = ParentArray.Item(ChildIndex)
    Else
        Dim new_root As ArrayObject
        Set new_root = New ArrayObject
        Call new_root.ReDimArray(0, C_CHUNK_SIZE - 1)
        Call ParentArray.Update(ChildIndex, new_root)
        Set pGetChildArray = new_root
    End If
End Function

Private Sub pExpandIfNeed(ByVal RemainIndex As Long)
    If RemainIndex <= 0 Then Exit Sub

    Dim new_root As ArrayObject
    Set new_root = New ArrayObject
    Call new_root.ReDimArray(0, C_CHUNK_SIZE - 1)
    Call new_root.Update(0, pRootArray)
    Set pRootArray = new_root
    pDepth = pDepth + 1
End Sub

Private Sub pInitializeItemType(ByVal ItemObject As Variant)
    Dim item_is_obj As Boolean
    item_is_obj = IsObject(ItemObject)

    If item_is_obj Then
        pIsObject = True
        If Not ItemObject Is Nothing Then
            pTypeName = TypeName(ItemObject)
            If TypeOf ItemObject Is IComparable Then
                pIsComparable = True
            End If
            If TypeOf ItemObject Is IEquatable Then
                pIsEquatable = True
            End If
            If TypeOf ItemObject Is IDuplicateCheckable Then
                pIsDuplicateCheckable = True
            End If
            If TypeOf ItemObject Is IStringable Then
                pIsStringable = True
            End If
        End If
    Else
        If pIsObject Then
            Err.Raise vbObjectError + 1, "Class ObjectList", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        Else
            ' プリミティブ型
            pTypeName = TypeName(ItemObject)
        End If
    End If
End Sub

Private Sub pCheckItemType(ByVal ItemObject As Variant)
    If pIsObject Then
        If Not IsObject(ItemObject) Then
            Err.Raise vbObjectError + 1, "Class ObjectList", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        End If
        If ItemObject Is Nothing Then Exit Sub
    End If

    If TypeName(ItemObject) <> pTypeName Then
        Err.Raise vbObjectError + 1, "Class ObjectList", "最初に設定された要素と型が異なります。(list: " & pTypeName & ", item: " & TypeName(ItemObject) & ")"
    End If
End Sub
