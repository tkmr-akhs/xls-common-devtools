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

Public Sub Test_Equals_ArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("value")

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.Equals(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array arguments are not supported") > 0
End Sub

Public Sub Test_NotEquals_ArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("expected")

    Dim actual_arr As Variant
    actual_arr = Array("actual")

    ' Act
    Call target_assert.NotEquals(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array arguments are not supported") > 0
End Sub

Public Sub Test_EqualsNumeric_NonNumericString_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.EqualsNumeric("abc", 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Numeric assertions require") > 0
End Sub

Public Sub Test_NotEqualsNumeric_NonNumericString_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.NotEqualsNumeric("abc", 1)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Numeric assertions require") > 0
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

Public Sub Test_EqualsArray_NonArrayArgument_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.EqualsArray("value", actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_DifferentArrays_DoesNotFail(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("expected")

    Dim actual_arr As Variant
    actual_arr = Array("actual")

    ' Act
    Call target_assert.NotEqualsArray(expected_arr, actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsFalse target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_NotEqualsArray_NonArrayExpected_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim actual_arr As Variant
    actual_arr = Array("value")

    ' Act
    Call target_assert.NotEqualsArray("value", actual_arr)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_NonArrayActual_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    Dim expected_arr As Variant
    expected_arr = Array("value")

    ' Act
    Call target_assert.NotEqualsArray(expected_arr, "value")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
End Sub

Public Sub Test_NotEqualsArray_BothNonArray_FailsWithExplicitMessage(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Call target_assert.NotEqualsArray("expected", "actual")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
    Assert.IsTrue InStr(target_assert.ResultMessage, "Array assertions require") > 0
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

Public Sub Test_AssertionCount_NewAssert_ReturnsZero(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    Dim actual_count As Long
    actual_count = target_assert.AssertionCount

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, actual_count
End Sub

Public Sub Test_AssertionCount_AfterSuccessfulAssertion_ReturnsOne(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    target_assert.Equals "same", "same"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub

Public Sub Test_AssertionCount_AfterFailedAssertion_ReturnsOne(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim target_assert As UnitTestAssert
    Set target_assert = New UnitTestAssert

    ' Act
    target_assert.Equals "expected", "actual"

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue target_assert.IsFailed
    Assert.EqualsNumeric 1, target_assert.AssertionCount
End Sub
