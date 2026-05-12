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

    Dim values As Dictionary
    Set values = New Dictionary
    Call test_util.SetValue(values, "stored", "arg1", 10&)

    ' Act
    Dim actual_value As String
    actual_value = test_util.GetValue(values, "arg1", 10&)

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "stored", actual_value
    Assert.EqualsNumeric 1, values.Count
End Sub

Public Sub Test_GetValue_StoredObject_ReturnsSameObject(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim values As Dictionary
    Set values = New Dictionary

    Dim expected_value As WorksheetRangeBounds
    Set expected_value = New_RangeBounds(Row:=1, Column:=2, Sheet:="SheetA", Book:="BookA.xlsm")
    Call test_util.SetValue(values, expected_value, "range")

    ' Act
    Dim actual_value As WorksheetRangeBounds
    Set actual_value = test_util.GetValue(values, "range")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_value, actual_value
    Assert.EqualsNumeric 1, values.Count
End Sub

Public Sub Test_GetValue_MissingKey_RaisesErrorAndDoesNotAddKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim test_util As UnitTestUtils
    Set test_util = New UnitTestUtils

    Dim values As Dictionary
    Set values = New Dictionary

    ' Act
    Dim actual_value As Variant
    actual_value = test_util.GetValue(values, "missing")

    ' Assert
    If Not Assert.ErrorRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Err.Clear
    Assert.EqualsNumeric 0, values.Count
End Sub
