Attribute VB_Name = "Test_UnitTestUtils"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! UnitTestUtils クラスのユニット テストです。
'!
' #############################################################################

Public Sub Test_GetValue_StoredString_ReturnsValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary
    Call test_util.SetValue(stored_values, "stored", "arg1", 10&)

    ' Act
    Dim actual_value As String
    actual_value = test_util.GetValue(stored_values, "arg1", 10&)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "stored", actual_value
    Assert.EqualsNumeric 1, stored_values.Count
End Sub

Public Sub Test_GetValue_StoredObject_ReturnsSameObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary

    Dim expected_value As WorksheetRangeBounds
    Set expected_value = New_RangeBounds(Row:=1, Column:=2, Sheet:="SheetA", Book:="BookA.xlsm")
    Call test_util.SetValue(stored_values, expected_value, "range")

    ' Act
    Dim actual_value As WorksheetRangeBounds
    Set actual_value = test_util.GetValue(stored_values, "range")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_value, actual_value
    Assert.EqualsNumeric 1, stored_values.Count
End Sub

Public Sub Test_GetValue_WithIEquatableArgumentWithoutPublicIdentityMember_ReturnsValue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary

    Dim stored_arg As Test_ObjectSetEquatableDouble
    Set stored_arg = New Test_ObjectSetEquatableDouble
    stored_arg.IdentityKey = "same-key"
    Call test_util.SetValue(stored_values, "stored", stored_arg)

    Dim lookup_arg As Test_ObjectSetEquatableDouble
    Set lookup_arg = New Test_ObjectSetEquatableDouble
    lookup_arg.IdentityKey = "same-key"

    ' Act
    Dim actual_value As String
    actual_value = test_util.GetValue(stored_values, lookup_arg)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "stored", actual_value
    Assert.EqualsNumeric 1, stored_values.Count
End Sub

Public Sub Test_GetValue_MissingKey_RaisesErrorAndDoesNotAddKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary

    ' Act
    Dim actual_value As Variant
    actual_value = test_util.GetValue(stored_values, "missing")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, stored_values.Count
End Sub

Public Sub Test_GetValue_WithSpecialPrimitiveArguments_ReturnsDistinctValues(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary

    Dim empty_arg As Variant
    Call test_util.SetValue(stored_values, "empty-value", empty_arg)
    Call test_util.SetValue(stored_values, "null-value", Null)
    Call test_util.SetValue(stored_values, "div0-value", CVErr(xlErrDiv0))
    Call test_util.SetValue(stored_values, "na-value", CVErr(xlErrNA))

    ' Act
    Dim actual_empty As String
    actual_empty = test_util.GetValue(stored_values, empty_arg)

    Dim actual_null As String
    actual_null = test_util.GetValue(stored_values, Null)

    Dim actual_div0 As String
    actual_div0 = test_util.GetValue(stored_values, CVErr(xlErrDiv0))

    Dim actual_na As String
    actual_na = test_util.GetValue(stored_values, CVErr(xlErrNA))

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.Equals "empty-value", actual_empty
    Assert.Equals "null-value", actual_null
    Assert.Equals "div0-value", actual_div0
    Assert.Equals "na-value", actual_na
    Assert.EqualsNumeric 4, stored_values.Count
End Sub

Public Sub Test_HasValue_WithNullArgument_ReturnsTrue(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim stored_values As Dictionary
    Set stored_values = New Dictionary
    Call test_util.SetValue(stored_values, "stored", Null)

    ' Act
    Dim actual_value As Boolean
    actual_value = test_util.HasValue(stored_values, Null)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.IsTrue actual_value
End Sub
