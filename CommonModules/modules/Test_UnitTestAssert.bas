Attribute VB_Name = "Test_UnitTestAssert"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! UnitTestAssert クラスのユニット テストです。
'!
' #############################################################################

Public Sub Test_Equals_NullValues_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(Null, Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_EmptyValues_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_value As Variant
    Dim actual_value As Variant

    ' Act
    Call target_assert.Equals(expected_value, actual_value)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrSameCode_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrDiv0), CVErr(xlErrDiv0))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrDifferentCode_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrDiv0), CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_Equals_CVErrAndNull_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.Equals(CVErr(xlErrValue), Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_Equals_NullAndEmpty_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim empty_value As Variant

    ' Act
    Call target_assert.Equals(Null, empty_value)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_EqualsNumeric_NullValues_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsNumeric(Null, Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_EmptyVariantArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array()

    Dim actual_arr As Variant
    actual_arr = Array()

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_UninitializedVariantArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr() As Variant
    Dim actual_arr() As Variant

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_OneEmptyOneNonEmpty_FailsWithoutError(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array()

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
End Sub

Public Sub Test_EqualsArray_WithSpecialPrimitiveElements_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim empty_value As Variant

    Dim expected_arr As Variant
    expected_arr = Array(Null, CVErr(xlErrNA), empty_value)

    Dim actual_arr As Variant
    actual_arr = Array(Null, CVErr(xlErrNA), empty_value)

    ' Act
    Call target_assert.EqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
End Sub
