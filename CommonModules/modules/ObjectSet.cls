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
    
    Dim result() As Variant
    Dim idx As Long
    
    ReDim result(0 To pDict.Count - 1)
    
    idx = LBound(result)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            Set result(idx) = pDict(dict_key)
        Else
            result(idx) = pDict(dict_key)
        End If
        
        idx = idx + 1
    Next dict_key
    
    ConvertToArray = result
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
    
    Dim result() As String
    Dim idx As Long
    Dim stringable_obj As IStringable
    
    ReDim result(0 To pDict.Count - 1)
    
    idx = LBound(result)
    Dim dict_key As Variant
    For Each dict_key In pDict
        If pIsObject Then
            If pDict(dict_key) Is Nothing Then
                result(idx) = "Nothing"
            Else
                If pIsStringable Then
                    Set stringable_obj = pDict(dict_key)
                    result(idx) = stringable_obj.ToString()
                Else
                    result(idx) = TypeName(pDict(dict_key))
                End If
            End If
        Else
            result(idx) = pDict(dict_key)
        End If
        idx = idx + 1
    Next dict_key
    
    ConvertToStringArray = result
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
    Dim ItemObject As Variant
    
    For Each ItemObject In ItemArray
        Call Add(ItemObject, ErrorIfExists)
    Next ItemObject
End Sub

'* 他の集合の要素をこの集合に取り込みます。
'*
'* @param OtherSet 取り込む要素を持つ他の集合。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他の集合の要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddOther(ByVal OtherSet As ObjectSet, Optional ByVal ErrorIfExists As Boolean = True)
    Dim idx As Long
    
    For idx = 0 To OtherSet.Count - 1
        Call Me.Add(OtherSet.Item(idx), ErrorIfExists:=ErrorIfExists)
    Next idx
End Sub

'* 他のリストの要素をこの集合に取り込みます。
'*
'* @param OtherList 取り込む要素を持つ他のリスト。
'* @param Optional ErrorIfExists 既に存在する場合にエラーを発生させるかどうか。
'*
'* @details
'* 他のリストの要素をこの集合に取り込みます。オプションで、要素が既に存在する場合の挙動を指定できます。
Public Sub AddList(ByVal OtherList As ObjectList, Optional ByVal ErrorIfExists As Boolean = True)
    Dim idx As Long
    
    For idx = 0 To OtherList.Count - 1
        Call Me.Add(OtherList.Item(idx), ErrorIfExists)
    Next idx
End Sub

'* この集合の複製を作成します。
'*
'* @return 複製された新しい集合。
'*
'* @details
'* この集合の複製を作成します。
Public Function CopySet() As ObjectSet
    Dim result As ObjectSet
    Set result = New ObjectSet
    
    Dim key_var As Variant
    For Each key_var In pDict.Keys
        Call result.Add(pDict(key_var))
    Next key_var
    
    Set CopySet = result
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
    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    
    If Not pDict.Exists(key_var) Then
        If Not IgnoreNotExists Then
            Err.Raise vbObjectError + 1, "Class ObjectSet", "存在しない要素を削除しようとしました。"
        End If
        If pIsObject Then Set RemoveItem = Nothing
        Exit Function
    End If
    
    If pIsObject Then
        Set RemoveItem = pDict(key_var)
    Else
        RemoveItem = pDict(key_var)
    End If
    
    Call pDict.Remove(key_var)
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
    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    Exists = pDict.Exists(key_var)
End Function

'* 引数と重複するとみなされる要素を返します。
'*
'* @param ItemObject 確認する要素。
'* @return 重複するとみなされる要素。
'*
'* @details
'* 引数と重複するとみなされる要素を返します。要素が存在しない場合は Nothing を返します。
Public Function GetContains(ByVal ItemObject As Variant) As Variant
    Dim key_var As Variant
    Call pGetKey(key_var, ItemObject)
    
    If Not pDict.Exists(key_var) Then
        If pIsObject Then
            Set GetContains = Nothing
        End If
        Exit Function
    End If
    
    If pIsObject Then
        Set GetContains = pDict(key_var)
    Else
        GetContains = pDict(key_var)
    End If
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
    Dim result As Enumerator
    Set result = New Enumerator
    Call result.Initialize(Me, Descending:=Descending, IsReadOnly:=IsReadOnly)
    Set GetEnumerator = result
End Function

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
            Err.Raise vbObjectError + 1, "Class ObjectSet", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        Else
            ' プリミティブ型
            pTypeName = TypeName(ItemObject)
        End If
    End If
End Sub

Private Sub pCheckItemType(ByVal ItemObject As Variant)
    If pIsObject Then
        If Not IsObject(ItemObject) Then
            Err.Raise vbObjectError + 1, "Class ObjectSet", "最初の要素はオブジェクト型でしたが、新しい要素はオブジェクト型ではありません。(" & pTypeName & ")"
        End If
        If ItemObject Is Nothing Then Exit Sub
    End If
    
    If TypeName(ItemObject) <> pTypeName Then
        Err.Raise vbObjectError + 1, "Class ObjectSet", "最初に設定された要素と型が異なります。(list: " & pTypeName & ", item: " & TypeName(ItemObject) & ")"
    End If
End Sub

Private Sub pGetKey(ByRef ObjectKey As Variant, ByVal ObjectItem As Variant)
    If IsObject(ObjectItem) Then
        If ObjectItem Is Nothing Then
            Set ObjectKey = ObjectItem
            Exit Sub
        End If
    End If
    
    If pIsDuplicateCheckable Then
        Dim objset_item As IDuplicateCheckable
        Set objset_item = ObjectItem
        ObjectKey = objset_item.GetKey()
    ElseIf pIsEquatable Then
        Dim eq_item As IEquatable
        Set eq_item = ObjectItem
        ObjectKey = eq_item.GetIdentityString
    ElseIf IsObject(ObjectItem) Then
        Set ObjectKey = ObjectItem
    Else
        ObjectKey = ObjectItem
    End If
End Sub

Private Function pKeyEquals(ByVal ItemKey1 As Variant, ByVal ItemKey2 As Variant) As Boolean
    If IsObject(ItemKey1) Then
        If IsObject(ItemKey2) Then
            If ItemKey1 Is ItemKey2 Then
                pKeyEquals = True
                Exit Function
            End If
        End If
    Else
        If Not IsObject(ItemKey2) Then
            If ItemKey1 = ItemKey2 Then
                pKeyEquals = True
                Exit Function
            End If
        End If
    End If
    
    pKeyEquals = False
End Function
