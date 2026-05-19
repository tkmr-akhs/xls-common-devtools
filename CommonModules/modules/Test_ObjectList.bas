Attribute VB_Name = "Test_ObjectList"
Option Explicit

' #############################################################################
'!
'! @brief
'! ObjectList クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

' =============================================================================
'  テンプレート
' =============================================================================
' Sub Test_試験対象_引数などの条件_戻り値などの予測結果(ByVal Assert As UnitTestAssert)
'     On Error Resume Next
'
'     '--- Arrange ---
'     Dim local_xxxx As Variant
'     Call ...
'
'     '--- Act ---
'     local_xxxx = ...
'
'     '--- Assert ---
'     If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
'
'     Assert.Equals ...
'
'     Assert.ErrorNotRaised 0, Err.Number, err.source, Err.Description
' End Sub
' =============================================================================

Private TUtl As New UnitTestUtils

' ----------------------------------------------------------------------------
' Add/Item
' ----------------------------------------------------------------------------

Public Sub Test_Count_AfterAddAndRemove_ReturnsItemCount(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Call obj_list.Add(10)
    Call obj_list.Add(20)
    Call obj_list.Remove(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_list.Count
End Sub

Public Sub Test_Add_AddLong_ReturnsLongItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim idx As Long
    
    ' Act/Assert
    For idx = 0 To 19
        Call obj_list.Add(idx)
    Next idx
    
    For idx = 0 To 19
        Assert.Equals idx, obj_list.Item(idx)
    Next idx
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Add_AddObject_ReturnsObjectItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    
    Dim expect_arr(0 To 19) As Object
    
    Dim idx As Long
    For idx = 0 To 19
        Set expect_arr(idx) = New ObjectList
    Next idx
    
    ' Act/Assert
    For idx = 0 To 19
        Call obj_list.Add(expect_arr(idx))
    Next idx
    
    For idx = 0 To 19
        Assert.Equals expect_arr(idx), obj_list.Item(idx)
    Next idx
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' Update
' ----------------------------------------------------------------------------

Public Sub Test_Update_UpdateLong_UpdateEntry(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    
    Dim idx As Long
    For idx = 0 To 19
        Call obj_list.Add(idx)
    Next idx
    
    ' Act
    Call obj_list.Update(16, 99&)
    
    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.Equals 99&, obj_list.Item(16)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Update_UpdateObject_UpdateEntry(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    
    Dim idx As Long
    For idx = 0 To 19
        Call obj_list.Add(New ObjectList)
    Next idx
    
    Dim expect_obj As Object
    Set expect_obj = New ObjectList
    
    ' Act
    
    Call obj_list.Update(16, expect_obj)
    
    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.Equals expect_obj, obj_list.Item(16)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub


' ----------------------------------------------------------------------------
' Remove
' ----------------------------------------------------------------------------

Public Sub Test_Remove_RemoveLong_ReturnsRemovedValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim idx As Long
    For idx = 0 To 9
        Call obj_list.Add(idx)
    Next idx
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.Remove(5)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 9, obj_list.Count
    If obj_list.Count <> 9 Then Exit Sub
    
    Assert.Equals 5&, actual_val
    Assert.EqualsNumeric 6, obj_list.Item(5)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Remove_RemoveObject_ReturnsRemovedObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim expect_obj_arr(0 To 4) As Object
    Dim idx As Long
    For idx = 0 To 4
        Set expect_obj_arr(idx) = New ObjectList
        Call obj_list.Add(expect_obj_arr(idx))
    Next idx
    
    '--- Act ---
    Dim actual_obj As Object
    Set actual_obj = obj_list.Remove(2)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 4, obj_list.Count
    If obj_list.Count <> 4 Then Exit Sub
    
    Assert.Equals expect_obj_arr(2), actual_obj
    Assert.Equals expect_obj_arr(3), obj_list.Item(2)
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' RemoveAll
' ----------------------------------------------------------------------------

Public Sub Test_RemoveAll_WhenCalled_ClearsList(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(1)
    Call obj_list.Add(2)
    Call obj_list.Add(3)
    
    '--- Act ---
    Call obj_list.RemoveAll
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 0, obj_list.Count
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' AddArray
' ----------------------------------------------------------------------------

Public Sub Test_AddArray_PassedArray_AddMultipleElements(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim expect_arr(0 To 2) As Variant
    expect_arr(0) = 10
    expect_arr(1) = 20
    expect_arr(2) = 30
    
    '--- Act ---
    Call obj_list.AddArray(expect_arr)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Dim idx As Long
    For idx = LBound(expect_arr) To UBound(expect_arr)
        Assert.Equals expect_arr(idx), obj_list.Item(idx)
    Next idx
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' ConvertToArray / ConvertToStringArray
' ----------------------------------------------------------------------------

Public Sub Test_ConvertToArray_WithItems_ReturnsArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("gamma")

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_list.ConvertToArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 2, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
    Assert.Equals "gamma", actual_arr(2)
End Sub

Public Sub Test_ConvertToStringArray_WithNumericItems_ReturnsStringArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(10&)
    Call obj_list.Add(20&)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_list.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "10", actual_arr(0)
    Assert.Equals "20", actual_arr(1)
End Sub

' ----------------------------------------------------------------------------
' AddOther / AddSet
' ----------------------------------------------------------------------------

Public Sub Test_AddOther_ObjectList_AppendsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    Dim other_list As ObjectList
    Set other_list = New ObjectList
    Call other_list.Add("beta")
    Call other_list.Add("gamma")

    ' Act
    Call obj_list.AddOther(other_list)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
    Assert.Equals "beta", obj_list.Item(1)
    Assert.Equals "gamma", obj_list.Item(2)
End Sub

Public Sub Test_AddSet_ObjectSet_AppendsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")

    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("beta")
    Call obj_set.Add("gamma")

    ' Act
    Call obj_list.AddSet(obj_set)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_list.Count
    Assert.Equals "alpha", obj_list.Item(0)
    Assert.Equals "beta", obj_list.Item(1)
    Assert.Equals "gamma", obj_list.Item(2)
End Sub

' ----------------------------------------------------------------------------
' CopyList
' ----------------------------------------------------------------------------

Public Sub Test_CopyList_WhenCalled_ReturnsSameElements(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(100)
    Call obj_list.Add(200)
    Call obj_list.Add(300)
    
    '--- Act ---
    Dim actual_obj As ObjectList
    Set actual_obj = obj_list.CopyList
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals obj_list.Count, actual_obj.Count
    Dim idx As Long
    For idx = 0 To obj_list.Count - 1
        Assert.Equals obj_list.Item(idx), actual_obj.Item(idx)
    Next idx
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' RemoveItem
' ----------------------------------------------------------------------------

Public Sub Test_RemoveItem_SpecifiedElement_Removes(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    ' リストに重複を含む数値を追加: 5, 10, 5, 15
    Call obj_list.Add(5)
    Call obj_list.Add(10)
    Call obj_list.Add(5)
    Call obj_list.Add(15)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.RemoveItem(5)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 2, obj_list.Count
    If obj_list.Count <> 2 Then Exit Sub
    
    Assert.Equals True, actual_val
    Dim idx As Long
    For idx = 0 To obj_list.Count - 1
        Assert.NotEquals 5, obj_list.Item(idx)
    Next idx
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' RemoveDuplicate
' ----------------------------------------------------------------------------

Public Sub Test_RemoveDuplicate_WhenCalled_RemovesDuplicates(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    ' 重複を含む数値を追加: 7, 8, 7, 9, 8, 10
    Call obj_list.Add(7)
    Call obj_list.Add(8)
    Call obj_list.Add(7)
    Call obj_list.Add(9)
    Call obj_list.Add(8)
    Call obj_list.Add(10)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.RemoveDuplicate
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 4, obj_list.Count
    If obj_list.Count <> 4 Then Exit Sub
    
    Assert.Equals True, actual_val
    Assert.Equals 7, obj_list.Item(0)
    Assert.Equals 8, obj_list.Item(1)
    Assert.Equals 9, obj_list.Item(2)
    Assert.Equals 10, obj_list.Item(3)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' PushItem / PopItem
' ----------------------------------------------------------------------------

Public Sub Test_PushItem_PushItem_AddsItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(1)
    Call obj_list.Add(2)
    
    '--- Act ---
    Call obj_list.PushItem(3)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 3, obj_list.Count
    If obj_list.Count <> 3 Then Exit Sub
    
    Assert.Equals 3, obj_list.Item(2)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_PopItem_PopItem_PopItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(1)
    Call obj_list.Add(2)
    Call obj_list.PushItem(3)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.PopItem()
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 2, obj_list.Count
    If obj_list.Count <> 2 Then Exit Sub
    
    Assert.Equals 3, actual_val
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' GetIndexByItem
' ----------------------------------------------------------------------------

Public Sub Test_GetIndexByItem_WhenExists_ReturnsCorrectIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(100)
    Call obj_list.Add(200)
    Call obj_list.Add(300)
    Call obj_list.Add(400)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.GetIndexByItem(300)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.Equals 2&, actual_val
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_GetIndexByItem_WhenNotExists_ReturnsCorrectIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(100)
    Call obj_list.Add(200)
    Call obj_list.Add(300)
    Call obj_list.Add(400)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.GetIndexByItem(500)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.Equals -1&, actual_val
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_GetIndexByItem_WithStartIndexAndReverse_ReturnsExpectedIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("alpha")
    Call obj_list.Add("beta")
    Call obj_list.Add("alpha")
    Call obj_list.Add("gamma")

    ' Act
    Dim actual_forward As Long
    actual_forward = obj_list.GetIndexByItem("alpha", StartIndex:=1)

    Dim actual_reverse As Long
    actual_reverse = obj_list.GetIndexByItem("alpha", Reverse:=True)

    Dim actual_reverse_from_middle As Long
    actual_reverse_from_middle = obj_list.GetIndexByItem("alpha", StartIndex:=1, Reverse:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_forward
    Assert.EqualsNumeric 2, actual_reverse
    Assert.EqualsNumeric 0, actual_reverse_from_middle
End Sub

' ----------------------------------------------------------------------------
' Exists
' ----------------------------------------------------------------------------

Public Sub Test_Exists_ReturnsTrueForExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(10)
    Call obj_list.Add(20)
    Call obj_list.Add(30)
    
    '--- Act ---
    Dim actual_val As Variant
    actual_val = obj_list.Exists(20)
    Assert.Equals True, actual_val
    actual_val = obj_list.Exists(40)
    Assert.Equals False, actual_val
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_GetIndexByItem_EmptyList_DoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem("1")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric -1, actual_idx

    Call obj_list.Add(CLng(1))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_ExistsAndGetIndexByItem_EmptyListNothing_DoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Set nothing_item = Nothing

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(nothing_item)

    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_exists
    Assert.EqualsNumeric -1, actual_idx

    Call obj_list.Add(CLng(1))
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_ExistsGetIndexAndRemoveItem_NothingOnlyNothing_FindsAndRemovesNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Set nothing_item = Nothing

    Call obj_list.Add(nothing_item)

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(nothing_item)

    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(nothing_item)

    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(nothing_item)

    Dim actual_exists_after_remove As Boolean
    actual_exists_after_remove = obj_list.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
    Assert.EqualsNumeric 0, actual_idx
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 0, obj_list.Count
    Assert.IsFalse actual_exists_after_remove
End Sub

Public Sub Test_GetIndexByItem_NothingOnlyNonNothing_ReturnsMissingAndDoesNotInitializeType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Call obj_list.Add(nothing_item)

    Dim search_item As Test_ObjectSetEquatableDouble
    Set search_item = New Test_ObjectSetEquatableDouble
    search_item.IdentityKey = "search-id"

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric -1, actual_idx

    Call obj_list.Add(search_item)
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals search_item, obj_list.Item(1)
End Sub

Public Sub Test_Exists_PrimitiveTypeMismatch_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CLng(1))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(CStr(1))

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
    Err.Clear
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_RemoveItem_PrimitiveTypeMismatch_RaisesTypeErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CLng(1))

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(CStr(1))

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
    Err.Clear
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_Exists_PrimitiveSetThenNothing_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CLng(1))

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Set nothing_item = Nothing

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
    Err.Clear
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_RemoveItem_PrimitiveSetThenNothing_RaisesTypeErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CLng(1))

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Set nothing_item = Nothing

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(nothing_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
    Err.Clear
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.EqualsNumeric 1, obj_list.Item(0)
End Sub

Public Sub Test_GetIndexByItem_ObjectTypeMismatch_RaisesTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Collection
    Set first_item = New Collection
    Call obj_list.Add(first_item)

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(CLng(1))

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
End Sub
' ----------------------------------------------------------------------------
' Comparison contract
' ----------------------------------------------------------------------------

Public Sub Test_GetIndexByItem_ObjectReference_ReturnsSameReferenceIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Collection
    Set first_item = New Collection

    Dim second_item As Collection
    Set second_item = New Collection

    Dim same_ref_item As Collection
    Set same_ref_item = second_item

    Call obj_list.Add(first_item)
    Call obj_list.Add(second_item)

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(same_ref_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_idx
End Sub

Public Sub Test_GetIndexByItem_DuplicateCheckable_ReturnsDuplicateKeyIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetDupCheckDouble
    Set first_item = New Test_ObjectSetDupCheckDouble
    first_item.DuplicateKey = "first-key"

    Dim second_item As Test_ObjectSetDupCheckDouble
    Set second_item = New Test_ObjectSetDupCheckDouble
    second_item.DuplicateKey = "target-key"

    Dim search_item As Test_ObjectSetDupCheckDouble
    Set search_item = New Test_ObjectSetDupCheckDouble
    search_item.DuplicateKey = "target-key"

    Call obj_list.Add(first_item)
    Call obj_list.Add(second_item)

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_idx
End Sub

Public Sub Test_GetIndexByItem_Equatable_ReturnsIdentityIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim second_item As Test_ObjectSetEquatableDouble
    Set second_item = New Test_ObjectSetEquatableDouble
    second_item.IdentityKey = "target-id"

    Dim search_item As Test_ObjectSetEquatableDouble
    Set search_item = New Test_ObjectSetEquatableDouble
    search_item.IdentityKey = "target-id"

    Call obj_list.Add(first_item)
    Call obj_list.Add(second_item)

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_idx
End Sub

Public Sub Test_GetIndexByItem_Nothing_ReturnsNothingIndex(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Call obj_list.Add(first_item)
    Call obj_list.Add(nothing_item)

    ' Act
    Dim actual_idx As Long
    actual_idx = obj_list.GetIndexByItem(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, actual_idx
End Sub

Public Sub Test_Exists_Nothing_ReturnsTrueForStoredNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetEquatableDouble
    Set nothing_item = Nothing

    Call obj_list.Add(nothing_item)

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_Exists_Equatable_ReturnsTrueForIdentityMatch(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim search_item As Test_ObjectSetEquatableDouble
    Set search_item = New Test_ObjectSetEquatableDouble
    search_item.IdentityKey = "first-id"

    Call obj_list.Add(first_item)

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_RemoveItem_EquatableWithNothing_RemovesMatchingItemsAndKeepsNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "target-id"

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Dim duplicate_item As Test_ObjectSetEquatableDouble
    Set duplicate_item = New Test_ObjectSetEquatableDouble
    duplicate_item.IdentityKey = "target-id"

    Dim search_item As Test_ObjectSetEquatableDouble
    Set search_item = New Test_ObjectSetEquatableDouble
    search_item.IdentityKey = "target-id"

    Call obj_list.Add(first_item)
    Call obj_list.Add(nothing_item)
    Call obj_list.Add(duplicate_item)

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(search_item)

    Dim actual_item As Test_ObjectSetEquatableDouble
    If Err.Number = 0 Then
        Set actual_item = obj_list.Item(0)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.IsNothing actual_item
End Sub

Public Sub Test_RemoveItem_Nothing_RemovesNothingAndKeepsOtherItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_list.Add(first_item)
    Call obj_list.Add(nothing_item)

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.Equals first_item, obj_list.Item(0)
End Sub

Public Sub Test_RemoveDuplicate_Equatable_RemovesDuplicateIdentityItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "same-id"

    Dim duplicate_item As Test_ObjectSetEquatableDouble
    Set duplicate_item = New Test_ObjectSetEquatableDouble
    duplicate_item.IdentityKey = "same-id"

    Dim other_item As Test_ObjectSetEquatableDouble
    Set other_item = New Test_ObjectSetEquatableDouble
    other_item.IdentityKey = "other-id"

    Call obj_list.Add(first_item)
    Call obj_list.Add(duplicate_item)
    Call obj_list.Add(other_item)

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveDuplicate()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals first_item, obj_list.Item(0)
    Assert.Equals other_item, obj_list.Item(1)
End Sub

Public Sub Test_RemoveDuplicate_DuplicateCheckableWithNothing_RemovesDuplicatesAndKeepsOneNothing(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim nothing_item As Test_ObjectSetDupCheckDouble

    Dim first_item As Test_ObjectSetDupCheckDouble
    Set first_item = New Test_ObjectSetDupCheckDouble
    first_item.DuplicateKey = "same-key"

    Dim duplicate_item As Test_ObjectSetDupCheckDouble
    Set duplicate_item = New Test_ObjectSetDupCheckDouble
    duplicate_item.DuplicateKey = "same-key"

    Call obj_list.Add(nothing_item)
    Call obj_list.Add(first_item)
    Call obj_list.Add(duplicate_item)
    Call obj_list.Add(nothing_item)

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveDuplicate()

    Dim actual_item As Test_ObjectSetDupCheckDouble
    If Err.Number = 0 Then
        Set actual_item = obj_list.Item(0)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.IsNothing actual_item
    Assert.Equals first_item, obj_list.Item(1)
End Sub

Public Sub Test_Add_ObjectListThenPrimitive_RaisesObjectListTypeError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_item As Collection
    Set first_item = New Collection

    Call obj_list.Add(first_item)

    ' Act
    Call obj_list.Add(1&)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class ObjectList", Err.Source
End Sub
' ----------------------------------------------------------------------------
' Sort
' ----------------------------------------------------------------------------

Public Sub Test_Sort_WhenCalled_SortsAscending(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim unsorted_vals As Variant
    unsorted_vals = Array(5, 3, 8, 1, 9)
    Dim idx As Long
    For idx = LBound(unsorted_vals) To UBound(unsorted_vals)
        Call obj_list.Add(unsorted_vals(idx))
    Next idx
    
    '--- Act ---
    Call obj_list.Sort(False)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Dim expect_sorted As Variant
    expect_sorted = Array(1, 3, 5, 8, 9)
    For idx = LBound(expect_sorted) To UBound(expect_sorted)
        Assert.Equals expect_sorted(idx), obj_list.Item(idx)
    Next idx
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_Sort_WhenCalledAsDescending_SortsDescending(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Dim unsorted_vals As Variant
    unsorted_vals = Array(5, 3, 8, 1, 9)
    Dim idx As Long
    For idx = LBound(unsorted_vals) To UBound(unsorted_vals)
        Call obj_list.Add(unsorted_vals(idx))
    Next idx
    
    '--- Act ---
    Call obj_list.Sort(True)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Dim expect_sorted As Variant
    expect_sorted = Array(9, 8, 5, 3, 1)
    For idx = LBound(expect_sorted) To UBound(expect_sorted)
        Assert.Equals expect_sorted(idx), obj_list.Item(idx)
    Next idx
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' Swap
' ----------------------------------------------------------------------------

Public Sub Test_Swap_WhenPassedIndexes_SwapsElements(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(111)
    Call obj_list.Add(222)
    
    '--- Act ---
    Call obj_list.Swap(0, 1)
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals 222, obj_list.Item(0)
    Assert.Equals 111, obj_list.Item(1)
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' GetEnumerator
' ----------------------------------------------------------------------------

Public Sub Test_GetEnumerator_WhenCalled_IteratesOverElements(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    '--- Arrange ---
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(10)
    Call obj_list.Add(20)
    Call obj_list.Add(30)
    
    '--- Act ---
    Dim Count As Long
    Count = 0
    Dim Enumerator As IEnumerator
    Set Enumerator = obj_list.GetEnumerator(False)
    Do While Enumerator.MoveNext
        Count = Count + 1
    Loop
    
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals 3&, Count
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

' ----------------------------------------------------------------------------
' Special Variant values
' ----------------------------------------------------------------------------

Public Sub Test_Add_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Call obj_list.Add(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_Add_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Call obj_list.Add(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_Exists_CVErr_ReturnsTrueForSameError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CVErr(xlErrNA))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_Exists_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Dim actual_exists As Boolean
    actual_exists = False
    actual_exists = obj_list.Exists(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_exists
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_Exists_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Dim actual_exists As Boolean
    actual_exists = False
    actual_exists = obj_list.Exists(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_exists
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_RemoveItem_CVErr_RemovesSameError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CVErr(xlErrNA))
    Call obj_list.Add(CVErr(xlErrValue))

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.Equals CVErr(xlErrValue), obj_list.Item(0)
End Sub

Public Sub Test_RemoveItem_Empty_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Dim actual_removed As Boolean
    actual_removed = False
    actual_removed = obj_list.RemoveItem(Empty)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_removed
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_RemoveItem_Null_RaisesError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Dim actual_removed As Boolean
    actual_removed = False
    actual_removed = obj_list.RemoveItem(Null)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse actual_removed
    Assert.EqualsNumeric 0, obj_list.Count
End Sub

Public Sub Test_RemoveDuplicate_CVErr_RemovesDuplicateSameError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CVErr(xlErrNA))
    Call obj_list.Add(CVErr(xlErrNA))
    Call obj_list.Add(CVErr(xlErrValue))

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveDuplicate()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals CVErr(xlErrNA), obj_list.Item(0)
    Assert.Equals CVErr(xlErrValue), obj_list.Item(1)
End Sub

Public Sub Test_ConvertToStringArray_CVErr_ReturnsErrorString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CVErr(xlErrNA))

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_list.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#N/A", actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_ModernCVErr_ReturnsErrorString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add(CVErr(xlErrSpill))

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_list.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "#SPILL!", actual_arr(0)
End Sub

Public Sub Test_ConvertToStringArray_ArrayItem_ReturnsTypedValueKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim item_arr As Variant
    item_arr = Array("alpha")
    Call obj_list.Add(item_arr)

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_list.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Variant[0:0](String(alpha))", actual_arr(0)
End Sub

Public Sub Test_Exists_ArrayItem_ReturnsTrueForSameTypedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim item_arr As Variant
    item_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(item_arr)

    Dim search_arr As Variant
    search_arr = Array(CLng(1), CStr("alpha"))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(search_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_exists
End Sub

Public Sub Test_Exists_ArrayItemWithDifferentElementType_ReturnsFalse(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim item_arr As Variant
    item_arr = Array(CLng(1))
    Call obj_list.Add(item_arr)

    Dim search_arr As Variant
    search_arr = Array(CStr(1))

    ' Act
    Dim actual_exists As Boolean
    actual_exists = obj_list.Exists(search_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_RemoveItem_ArrayItem_RemovesMatchingTypedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_arr As Variant
    first_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(first_arr)

    Dim other_arr As Variant
    other_arr = Array(CLng(2), CStr("beta"))
    Call obj_list.Add(other_arr)

    Dim search_arr As Variant
    search_arr = Array(CLng(1), CStr("alpha"))

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveItem(search_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.IsTrue IsArray(obj_list.Item(0))
    Assert.Equals "Variant[0:1](Long(2),String(beta))", GetTypedValueKey(obj_list.Item(0))
End Sub

Public Sub Test_RemoveDuplicate_ArrayItem_RemovesDuplicateTypedValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim first_arr As Variant
    first_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(first_arr)

    Dim duplicate_arr As Variant
    duplicate_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(duplicate_arr)

    Dim other_arr As Variant
    other_arr = Array(CStr(1), CStr("alpha"))
    Call obj_list.Add(other_arr)

    ' Act
    Dim actual_removed As Boolean
    actual_removed = obj_list.RemoveDuplicate()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_removed
    Assert.EqualsNumeric 2, obj_list.Count
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(obj_list.Item(0))
    Assert.Equals "Variant[0:1](String(1),String(alpha))", GetTypedValueKey(obj_list.Item(1))
End Sub

Public Sub Test_ConvertToArray_ArrayItem_ReturnsArrayItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    Dim item_arr As Variant
    item_arr = Array(CLng(1), CStr("alpha"))
    Call obj_list.Add(item_arr)

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_list.ConvertToArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 0, UBound(actual_arr)
    Assert.IsTrue IsArray(actual_arr(0))
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(actual_arr(0))
End Sub

Public Sub Test_AddSet_ArrayItems_AppendsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim item_arr As Variant
    item_arr = Array(CLng(1), CStr("alpha"))
    Call obj_set.Add(item_arr)

    Dim obj_list As ObjectList
    Set obj_list = New ObjectList

    ' Act
    Call obj_list.AddSet(obj_set)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_list.Count
    Assert.IsTrue IsArray(obj_list.Item(0))
    Assert.Equals "Variant[0:1](Long(1),String(alpha))", GetTypedValueKey(obj_list.Item(0))
End Sub
