Attribute VB_Name = "Test_ObjectSet"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! ObjectSet クラスのユニット テストです。
'!
' #############################################################################

' -----------------------------------------------------------------------------
' Add/Item/Count
' -----------------------------------------------------------------------------

Public Sub Test_Add_AddStrings_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add("alpha")
    Call obj_set.Add("beta")
    Call obj_set.Add("gamma")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_Add_DuplicateWithErrorIgnored_KeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Call obj_set.Add("alpha", ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_AddArray_StringArray_AddsEachItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_arr(0 To 2) As String
    src_arr(0) = "alpha"
    src_arr(1) = "beta"
    src_arr(2) = "gamma"

    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    ' Act
    Call obj_set.AddArray(src_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_AddOther_ObjectSet_AddsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    Dim other_set As ObjectSet
    Set other_set = New ObjectSet
    Call other_set.Add("beta")
    Call other_set.Add("gamma")

    ' Act
    Call obj_set.AddOther(other_set)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_AddList_ObjectListWithDuplicateIgnored_AddsUniqueItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    Dim obj_list As ObjectList
    Set obj_list = New ObjectList
    Call obj_list.Add("beta")
    Call obj_list.Add("alpha")
    Call obj_list.Add("gamma")

    ' Act
    Call obj_set.AddList(obj_list, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

Public Sub Test_ConvertToArray_WithItems_ReturnsArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_arr() As Variant
    actual_arr = obj_set.ConvertToArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
End Sub

Public Sub Test_ConvertToStringArray_WithItems_ReturnsStringArray(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_arr() As String
    actual_arr = obj_set.ConvertToStringArray()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 0, LBound(actual_arr)
    Assert.EqualsNumeric 1, UBound(actual_arr)
    Assert.Equals "alpha", actual_arr(0)
    Assert.Equals "beta", actual_arr(1)
End Sub

Public Sub Test_Add_DuplicateEquatableWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As WorksheetRangeBounds
    Set first_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Dim duplicate_item As WorksheetRangeBounds
    Set duplicate_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_Add_Nothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableDouble

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    Dim actual_item As Test_ObjectSetEquatableDouble
    Set actual_item = obj_set.Item(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_DuplicateNothingWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_RemoveItem_Nothing_RemovesNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_set.Add(nothing_item)

    ' Act
    Dim removed_item As Test_ObjectSetEquatableDouble
    Set removed_item = obj_set.RemoveItem(nothing_item)

    Dim actual_exists As Boolean
    actual_exists = obj_set.Exists(nothing_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 0, obj_set.Count
    Assert.IsFalse actual_exists
End Sub

Public Sub Test_Add_NothingFirstThenDuplicateCheckable_DetectsDuplicateByGetKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetDupCheckDouble

    Dim first_item As Test_ObjectSetDupCheckDouble
    Set first_item = New Test_ObjectSetDupCheckDouble
    first_item.DuplicateKey = "same-key"

    Dim duplicate_item As Test_ObjectSetDupCheckDouble
    Set duplicate_item = New Test_ObjectSetDupCheckDouble
    duplicate_item.DuplicateKey = "same-key"

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenDuplicateCheckable_RejectsOtherType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetDupCheckDouble

    Dim first_item As Test_ObjectSetDupCheckDouble
    Set first_item = New Test_ObjectSetDupCheckDouble
    first_item.DuplicateKey = "first-key"

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Add(other_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenEquatable_DetectsDuplicateByIdentityString(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "same-id"

    Dim duplicate_item As Test_ObjectSetEquatableDouble
    Set duplicate_item = New Test_ObjectSetEquatableDouble
    duplicate_item.IdentityKey = "same-id"

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenEquatable_RejectsOtherType(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Add(other_item)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_Add_NothingFirstThenObjectReference_UsesReferenceIdentity(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim nothing_item As Collection

    Dim first_item As Collection
    Set first_item = New Collection

    Dim same_ref_item As Collection
    Set same_ref_item = first_item

    Dim other_item As Collection
    Set other_item = New Collection

    Call obj_set.Add(nothing_item)
    Call obj_set.Add(first_item)

    ' Act
    Dim actual_same_ref_added As Boolean
    actual_same_ref_added = obj_set.Add(same_ref_item, ErrorIfExists:=False)

    Dim actual_other_added As Boolean
    actual_other_added = obj_set.Add(other_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_same_ref_added
    Assert.IsTrue actual_other_added
    Assert.EqualsNumeric 3, obj_set.Count
End Sub

Public Sub Test_Add_ObjectImplementsDuplicateCheckableAndEquatable_UsesDuplicateCheckableKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetKeyPriorityDouble
    Set first_item = New Test_ObjectSetKeyPriorityDouble
    first_item.DuplicateKey = "same-duplicate-key"
    first_item.IdentityKey = "first-identity"

    Dim duplicate_by_key_item As Test_ObjectSetKeyPriorityDouble
    Set duplicate_by_key_item = New Test_ObjectSetKeyPriorityDouble
    duplicate_by_key_item.DuplicateKey = "same-duplicate-key"
    duplicate_by_key_item.IdentityKey = "second-identity"

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(duplicate_by_key_item, ErrorIfExists:=False)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 1, obj_set.Count
End Sub
Public Sub Test_Add_DuplicateCheckableSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetDupCheckDouble
    Set first_item = New Test_ObjectSetDupCheckDouble
    first_item.DuplicateKey = "first-key"

    Dim nothing_item As Test_ObjectSetDupCheckDouble

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Test_ObjectSetDupCheckDouble
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_EquatableSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Test_ObjectSetEquatableDouble
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_ObjectReferenceSetThenNothing_AddsNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Collection
    Set first_item = New Collection

    Dim nothing_item As Collection

    Call obj_set.Add(first_item)

    ' Act
    Dim actual_added As Boolean
    actual_added = obj_set.Add(nothing_item)

    Dim actual_exists As Boolean
    Dim actual_item As Collection
    If Err.Number = 0 Then
        actual_exists = obj_set.Exists(nothing_item)
        Set actual_item = obj_set.Item(1)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_added
    Assert.EqualsNumeric 2, obj_set.Count
    Assert.IsTrue actual_exists
    Assert.IsNothing actual_item
End Sub

Public Sub Test_Add_EquatableSetThenDuplicateNothingWithErrorIgnored_ReturnsFalseAndKeepsExistingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_set.Add(first_item)
    Call obj_set.Add(nothing_item)

    ' Act
    Dim actual_added As Boolean
    If Err.Number = 0 Then
        actual_added = obj_set.Add(nothing_item, ErrorIfExists:=False)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsFalse actual_added
    Assert.EqualsNumeric 2, obj_set.Count
End Sub

Public Sub Test_RemoveItem_EquatableSetThenNothing_RemovesNothingItem(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "first-id"

    Dim nothing_item As Test_ObjectSetEquatableDouble

    Call obj_set.Add(first_item)
    Call obj_set.Add(nothing_item)

    ' Act
    Dim removed_item As Test_ObjectSetEquatableDouble
    Dim actual_exists As Boolean
    If Err.Number = 0 Then
        Set removed_item = obj_set.RemoveItem(nothing_item)
        If Err.Number = 0 Then
            actual_exists = obj_set.Exists(nothing_item)
        End If
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsFalse actual_exists
End Sub
' -----------------------------------------------------------------------------
' Update / GetContains
' -----------------------------------------------------------------------------

Public Sub Test_Update_EquatableSameIdentity_ReplacesStoredObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim first_item As Test_ObjectSetEquatableDouble
    Set first_item = New Test_ObjectSetEquatableDouble
    first_item.IdentityKey = "same-id"

    Dim replacement_item As Test_ObjectSetEquatableDouble
    Set replacement_item = New Test_ObjectSetEquatableDouble
    replacement_item.IdentityKey = "same-id"

    Call obj_set.Add(first_item)

    ' Act
    Call obj_set.Update(0, replacement_item)

    Dim actual_item As Test_ObjectSetEquatableDouble
    If Err.Number = 0 Then
        Set actual_item = obj_set.Item(0)
    End If

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals replacement_item, actual_item
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.IsTrue obj_set.Exists(first_item)
End Sub

Public Sub Test_GetContains_Equatable_ReturnsStoredObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim stored_item As Test_ObjectSetEquatableDouble
    Set stored_item = New Test_ObjectSetEquatableDouble
    stored_item.IdentityKey = "same-id"

    Dim search_item As Test_ObjectSetEquatableDouble
    Set search_item = New Test_ObjectSetEquatableDouble
    search_item.IdentityKey = "same-id"

    Call obj_set.Add(stored_item)

    ' Act
    Dim actual_item As Test_ObjectSetEquatableDouble
    Set actual_item = obj_set.GetContains(search_item)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals stored_item, actual_item
End Sub

' -----------------------------------------------------------------------------
' Exists/Remove
' -----------------------------------------------------------------------------

Public Sub Test_Exists_PresentAndMissing_ReturnsExpectedResult(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_present As Boolean
    Dim actual_missing As Boolean
    actual_present = obj_set.Exists("alpha")
    actual_missing = obj_set.Exists("gamma")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTrue actual_present
    Assert.IsFalse actual_missing
End Sub

Public Sub Test_Remove_RemoveByIndex_RemovesItemAndReturnsValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim removed_value As String
    removed_value = obj_set.Remove(0)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha", removed_value
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "beta", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingWithIgnoreNotExists_DoesNotRaiseAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Dim removed_value As Variant
    removed_value = obj_set.RemoveItem("missing", IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingWithErrorNotIgnored_RaisesErrorAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Dim removed_value As Variant
    removed_value = obj_set.RemoveItem("missing", IgnoreNotExists:=False)

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
End Sub

Public Sub Test_RemoveItem_MissingObjectWithIgnoreNotExists_ReturnsNothingAndKeepsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet

    Dim existing_item As WorksheetRangeBounds
    Set existing_item = New_RangeBounds(Row:=1, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Dim missing_item As WorksheetRangeBounds
    Set missing_item = New_RangeBounds(Row:=2, Column:=1, Sheet:="SheetA", Book:="BookA.xlsm")

    Call obj_set.Add(existing_item)

    ' Act
    Dim removed_item As WorksheetRangeBounds
    Set removed_item = obj_set.RemoveItem(missing_item, IgnoreNotExists:=True)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsNothing removed_item
    Assert.EqualsNumeric 1, obj_set.Count
End Sub

Public Sub Test_RemoveAll_AfterItems_ClearsItemsAndTypeState(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")

    ' Act
    Call obj_set.RemoveAll
    Call obj_set.Add(10&)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 1, obj_set.Count
    Assert.EqualsNumeric 10, obj_set.Item(0)
End Sub

' -----------------------------------------------------------------------------
' CopySet
' -----------------------------------------------------------------------------

Public Sub Test_CopySet_WithItems_ReturnsCopiedSet(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Dim actual_set As ObjectSet
    Set actual_set = obj_set.CopySet()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 2, actual_set.Count
    Assert.Equals "alpha", actual_set.Item(0)
    Assert.Equals "beta", actual_set.Item(1)
End Sub
' -----------------------------------------------------------------------------
' Sort
' -----------------------------------------------------------------------------

Public Sub Test_Sort_WhenCalled_SortsItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("gamma")
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")

    ' Act
    Call obj_set.Sort()

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.EqualsNumeric 3, obj_set.Count
    Assert.Equals "alpha", obj_set.Item(0)
    Assert.Equals "beta", obj_set.Item(1)
    Assert.Equals "gamma", obj_set.Item(2)
End Sub

' -----------------------------------------------------------------------------
' Sort
' -----------------------------------------------------------------------------



' -----------------------------------------------------------------------------
' GetEnumerator
' -----------------------------------------------------------------------------

Public Sub Test_GetEnumerator_WhenCalled_IteratesOverItems(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim obj_set As ObjectSet
    Set obj_set = New ObjectSet
    Call obj_set.Add("alpha")
    Call obj_set.Add("beta")
    Call obj_set.Add("gamma")

    ' Act
    Dim enum_obj As IEnumerator
    Set enum_obj = obj_set.GetEnumerator()

    Dim actual_value As String
    Do While enum_obj.MoveNext()
        If actual_value = "" Then
            actual_value = enum_obj.Current
        Else
            actual_value = actual_value & "," & enum_obj.Current
        End If
    Loop

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub
