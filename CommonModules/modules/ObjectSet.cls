VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! 任意のオブジェクトまたは値を格納し、効率的にアクセスできるようにするための集合を提供します。
'! 各要素はユニークなキーによって管理されます。
'!
' #############################################################################

Private pDict As Dictionary
Private pTypeName As String
Private pIsObject As Boolean
Private pIsComparable As Boolean
Private pIsStringable As Boolean
Private pIsEquatable As Boolean
Private pIsDuplicateCheckable As Boolean

'* 集合内の要素の総数。
Public Property Get Count() As Long
    Count = pDict.Count
End Property

'* 集合内の要素の型。
Public Property Get ItemTypeName() As String
    ItemTypeName = pTypeName
End Property

'* 指定されたインデックスの要素を返します。
'*
'* @param ItemIndex 要素のインデックス。
'* @return 指定されたインデックスの要素。
'*
'* @details
'* 指定されたインデックスにある要素を返します。要素は内部的にキーに基づいてアクセスされます。
Public Function Item(ByVal ItemIndex As Long) As Variant
    If pIsObject Then
        Set Item = pDict(pDict.Keys(ItemIndex))
    Else
        Item = pDict(pDict.Keys(ItemIndex))
    End If
End Function

Private Sub Class_Initialize()
    Set pDict = New Dictionary
End Sub

'* 指定されたインデックスの要素を新しいオブジェクトで更新します。
'*
'* @param ItemIndex 更新する要素のインデックス。
'* @param ItemObject 新しい要素。
'*
'* @details
'* 指定されたインデックスの要素を新しいオブジェクトで更新します。キーが変更される場合はエラーが発生します。
Public Sub Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant)
    Call pCheckItemType(ItemObject)
    
    Dim old_key As Variant
    If IsObject(pDict.Keys(ItemIndex)) Then
        Set old_key = pDict.Keys(ItemIndex)
    Else
        old_key = pDict.Keys(ItemIndex)
    End If
    
    Dim new_key As Variant
    Call pGetKey(new_key, ItemObject)
    
    If Not pKeyEquals(old_key, new_key) Then
        Err.Raise vbObjectError + 1, "Class ObjectSet", "元のオブジェクトと、変更後のオブジェクトのキーが違います。変更が必要な場合は、元の要素を Remove 後に新しい要素を Add してください。(old_key: " & old_key & " / new_key: " & new_key & ")"
        Exit Sub
    End If
    
    If pIsObject Then
        Set pDict(old_key) = ItemObject
    Else
        pDict(old_key) = ItemObject
    End If
End Sub

'* 集合を配列として返します。
'*
'* @return 配列形式の要素。
'*
'* @details
'* 集合のすべての要素を配列形式で返します。
Public Function ConvertToArray() As Variant()
    If pDict.Count < 1 Then
        Exit Function
    End If
    
    Dim result_value() As Variant
    Dim item_idx As Long
    
    ReDim result_value(0 To pDict.Count - 1)
    
    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set result_value(item_idx) = pDict(dict_key)
        Else
            result_value(item_idx) = pDict(dict_key)
        End If
        
        item_idx = item_idx + 1
    Next dict_key
    
    ConvertToArray = result_value
End Function

'* 集合を文字列の配列として返します。
'*
'* @return 文字列形式の配列。
'*
'* @details
'* 集合のすべての要素を文字列形式で配列にして返します。
Public Function ConvertToStringArray() As String()
    If pDict.Count < 1 Then
        Exit Function
    End If
    
    Dim result_value() As String
    Dim item_idx As Long
    Dim stringable_obj As IStringable
    
    ReDim result_value(0 To pDict.Count - 1)
    
    item_idx = LBound(result_value)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            If pDict(dict_key) Is Nothing Then
                result_value(item_idx) = "Nothing"
            Else
                If pIsStringable Then
                    Set stringable_obj = pDict(dict_key)
                    result_value(item_idx) = stringable_obj.ToString()
                Else
                    result_value(item_idx) = GetTypeString(pDict(dict_key))
                End If
            End If
        Else
            If IsError(pDict(dict_key)) Then
                result_value(item_idx) = pConvertErrorToString(pDict(dict_key))
            ElseIf IsArray(pDict(dict_key)) Then
                result_value(item_idx) = GetTypedValueKey(pDict(dict_key))
            Else
                result_value(item_idx) = pDict(dict_key)
            End If
        End If
        item_idx = item_idx + 1
    Next dict_key
    
    ConvertToStringArray = result_value
End Function

'* 新しい要素を集合に追加します。
'*
'* @param ItemObject 追加する要素。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'* @return 要素が追加されたかどうかを示すブール値。
'*
'* @details
'* 新しい要素を集合に追加します。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Function Add(ByVal ItemObject As Variant, Optional ByVal ErrorIfExists As Boolean = True) As Boolean
    If pTypeName = "" Then
        Call pInitializeItemType(ItemObject)
    Else
        Call pCheckItemType(ItemObject)
    End If
    
    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    
    If Not ErrorIfExists Then
        If pDict.Exists(key_var) Then
            Add = False
            Exit Function
        End If
    End If
    
    Call pDict.Add(key_var, ItemObject)
    Add = True
End Function

'* 配列の各要素を集合に追加します。
'*
'* @param ItemArray 追加する要素の配列。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 配列の各要素を集合に追加します。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddArray(ByVal ItemArray As Variant, Optional ByVal ErrorIfExists As Boolean = True)
    If Not IsArray(ItemArray) Then Err.Raise vbObjectError + 1, "Class ObjectSet", "引数が配列ではありません"
    If IsEmptyArray(ItemArray) Then Exit Sub
    
    Dim key_var As Variant
    Dim set_item_obj As IDuplicateCheckable
    Dim item_obj As Variant

    For Each item_obj In ItemArray
        Call Add(item_obj, ErrorIfExists)
    Next item_obj
End Sub

'* 他の集合の要素をこの集合に取り込みます。
'*
'* @param OtherSet 取り込む要素を持つ他の集合。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他の集合の要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddOther(ByVal OtherSet As ObjectSet, Optional ByVal ErrorIfExists As Boolean = True)
    Dim item_idx As Long
    
    For item_idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(item_idx), ErrorIfExists:=ErrorIfExists)
    Next item_idx
End Sub

'* 他のリストの要素をこの集合に取り込みます。
'*
'* @param OtherList 取り込む要素を持つ他のリスト。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他のリストの要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddList(ByVal OtherList As ObjectList, Optional ByVal ErrorIfExists As Boolean = True)
    Dim item_idx As Long
    
    For item_idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(item_idx), ErrorIfExists)
    Next item_idx
End Sub

'* この集合の複製を作成します。
'*
'* @return 複製された新しい集合。
'*
'* @details
'* この集合の複製を作成します。
Public Function CopySet() As ObjectSet
    Dim result_value As ObjectSet
    Set result_value = New ObjectSet
    
    Dim key_var As Variant
    For Each key_var In pDict.Keys
        Call result_value.Add(pDict(key_var))
    Next key_var
    
    Set CopySet = result_value
End Function

'* 指定されたインデックスの要素を削除します。
'*
'* @param Index 削除する要素のインデックス。
'* @return 削除された要素。
'*
'* @details
'* 指定されたインデックスの要素を削除します。
Public Function Remove(ByVal Index As Long) As Variant
    Dim old_key As Variant
    Dim old_obj As Variant
    
    If pIsObject Then
        Set old_obj = pDict.Item(pDict.Keys(Index))
    Else
        old_obj = pDict.Item(pDict.Keys(Index))
    End If
    
    If pIsObject Then
        Set Remove = RemoveItem(old_obj)
    Else
        Remove = RemoveItem(old_obj)
    End If
End Function

'* 指定された要素を削除します。
'*
'* @param ItemObject 削除する要素。
'* @param Optional IgnoreNotExists 存在しない場合に無視するかどうか。
'* @return 削除された要素。
'*
'* @details
'* 指定された要素と重複と判断される要素を削除します。オプションで、要素が存在しない場合の挙動を指定できます。
Public Function RemoveItem(ByVal ItemObject As Variant, Optional ByVal IgnoreNotExists As Boolean = True) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet", "存在しない要素を削除しようとしました。"
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet", "存在しない要素を削除しようとしました。"
        End If
        If pIsObject Or IsObject(ItemObject) Then
            Set RemoveItem = Nothing
            Exit Function
        Else
            RemoveItem = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Call pDict.Remove(key_var)
        Set RemoveItem = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)
    Call pDict.Remove(key_var)

    RemoveItem = result_value
End Function

'* 集合のすべての要素を削除します。
'*
'* @details
'* 集合のすべての要素を削除します。
Public Sub RemoveAll()
    Call pDict.RemoveAll
    pTypeName = ""
    pIsObject = False
    pIsComparable = False
    pIsEquatable = False
    pIsDuplicateCheckable = False
    pIsStringable = False
End Sub

'* 引数の要素がすでに含まれているかを確認します。
'*
'* @param ItemObject 確認する要素。
'* @return 要素が含まれているかどうかを示すブール値。
'*
'* @details
'* 指定されたオブジェクトと重複している要素が含まれるかを確認します。
Public Function Exists(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    If Not pCanSearchItem(ItemObject) Then
        Exists = result_value
        Exit Function
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    result_value = pDict.Exists(key_var)

    Exists = result_value
End Function

'* 引数と重複するとみなされる要素を返します。
'*
'* @param ItemObject 確認する要素。
'* @return 重複するとみなされる要素。
'*
'* @details
'* 引数と重複するとみなされる要素を返します。要素が存在しない場合は Nothing を返します。
Public Function GetContains(ByVal ItemObject As Variant) As Variant
    If Not pCanSearchItem(ItemObject) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set GetContains = Nothing
            Exit Function
        Else
            GetContains = Empty
            Exit Function
        End If
    End If

    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)

    If Not pDict.Exists(key_var) Then
        If pIsObject Or IsObject(ItemObject) Then
            Set GetContains = Nothing
            Exit Function
        Else
            GetContains = Empty
            Exit Function
        End If
    End If

    If pIsObject Then
        Dim result_obj As Object
        Set result_obj = pDict(key_var)
        Set GetContains = result_obj
        Exit Function
    End If

    Dim result_value As Variant
    result_value = pDict(key_var)

    GetContains = result_value
End Function

'* この集合の中身をソートします。
'*
'* @param Optional Descending 降順にソートするかどうか。
'*
'* @details
'* この集合の中身をソートします。
'* 内部的に ObjectList.Sort() を利用するため、挙動の詳細はそちらを参照してください。
Public Sub Sort(Optional ByVal Descending As Boolean = False)
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.AddSet(Me)
    Call obj_list.Sort(Descending:=Descending)
    Call Me.RemoveAll
    Call Me.AddList(obj_list)
End Sub

'* この集合の列挙子を得ます。
'*
'* @param Descending [省略可] 降順で列挙するか否か。
'* @param IsReadOnly [省略可] 読み取り専用にするか否か。
'* @return 列挙子オブジェクト。
'*
'* @details
'* この集合の列挙子を得ます。
Public Function GetEnumerator(Optional ByVal Descending As Boolean = False, Optional ByVal IsReadOnly As Boolean = False) As IEnumerator
    Dim result_value As Enumerator
    Set result_value = New Enumerator
    Call result_value.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result_value
End Function

Private Sub pInitializeItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    Dim item_is_obj As Boolean
    item_is_obj = IsObject(ItemObject)
    
    If item_is_obj Then
        pIsObject = True
        If Not ItemObject Is Nothing Then
            pTypeName = GetTypeString(ItemObject)
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
            Err.Raise vbObjectError + 1, "Class ObjectSet", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        Else
            ' プリミティブ型
            pTypeName = GetTypeString(ItemObject)
        End If
    End If
End Sub

Private Function pCanSearchItem(ByVal ItemObject As Variant) As Boolean
    Dim result_value As Boolean
    result_value = False

    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pTypeName <> "" Then
        Call pCheckItemType(ItemObject)
        result_value = True
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not pIsObject Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If Not IsObject(ItemObject) Then
        pCanSearchItem = result_value
        Exit Function
    End If

    If ItemObject Is Nothing Then
        result_value = True
    Else
        result_value = False
    End If

    pCanSearchItem = result_value
End Function

Private Sub pCheckItemType(ByVal ItemObject As Variant)
    Call pCheckUnsupportedSpecialValue(ItemObject)

    If pIsObject Then
        If Not IsObject(ItemObject) Then
            Err.Raise vbObjectError + 1, "Class ObjectSet", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        End If
        If ItemObject Is Nothing Then Exit Sub
    End If
    
    If GetTypeString(ItemObject) <> pTypeName Then
        Err.Raise vbObjectError + 1, "Class ObjectSet", "最初に設定された要素と型が異なります。(list: " & pTypeName & ", item: " & GetTypeString(ItemObject) & ")"
    End If
End Sub

Private Sub pGetKey(ByRef ObjectKey As Variant, ByVal ObjectItem As Variant)
    Call pCheckUnsupportedSpecialValue(ObjectItem)

    Dim object_mode As Long
    If pIsDuplicateCheckable Then
        object_mode = G_TYPED_VALUE_STRING_OBJECT_DUPLICATE_CHECKABLE
    ElseIf pIsEquatable Then
        object_mode = G_TYPED_VALUE_STRING_OBJECT_I_EQUATABLE
    Else
        object_mode = G_TYPED_VALUE_STRING_OBJECT_REFERENCE
    End If

    ObjectKey = GetTypedValueKey(ObjectItem, object_mode)
End Sub

Private Sub pCheckUnsupportedSpecialValue(ByVal ItemObject As Variant)
    If IsObject(ItemObject) Then Exit Sub
    If IsArray(ItemObject) Then Exit Sub
    If IsEmpty(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet", "Empty は要素として使用できません。"
    If IsNull(ItemObject) Then Err.Raise vbObjectError + 1, "Class ObjectSet", "Null は要素として使用できません。"
End Sub

Private Function pConvertErrorToString(ByVal ErrorValue As Variant) As String
    Dim error_texts As Dictionary
    Set error_texts = New Dictionary
    Call error_texts.Add(CVErr(xlErrNull), "#NULL!")
    Call error_texts.Add(CVErr(xlErrDiv0), "#DIV/0!")
    Call error_texts.Add(CVErr(xlErrValue), "#VALUE!")
    Call error_texts.Add(CVErr(xlErrRef), "#REF!")
    Call error_texts.Add(CVErr(xlErrName), "#NAME?")
    Call error_texts.Add(CVErr(xlErrNum), "#NUM!")
    Call error_texts.Add(CVErr(xlErrNA), "#N/A")

    If error_texts.Exists(ErrorValue) Then
        pConvertErrorToString = error_texts(ErrorValue)
    Else
        pConvertErrorToString = "Unknown error"
    End If
End Function

Private Function pErrorValuesEqual(ByVal ErrorValue1 As Variant, ByVal ErrorValue2 As Variant) As Boolean
    Dim error_values As Dictionary
    Set error_values = New Dictionary
    Call error_values.Add(ErrorValue1, True)
    pErrorValuesEqual = error_values.Exists(ErrorValue2)
End Function

Private Function pKeyEquals(ByVal ItemKey1 As Variant, ByVal ItemKey2 As Variant) As Boolean
    Dim result_value As Boolean
    result_value = (CStr(ItemKey1) = CStr(ItemKey2))

    pKeyEquals = result_value
End Function
