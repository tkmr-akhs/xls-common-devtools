Attribute VB_Name = "Test_DiffStringArray"
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
'Sub Test_試験対象_引数などの条件_戻り値などの予測結果(ByVal Assert As UnitTestAssert)
'    On Error Resume Next
'
'    '--- Arrange ---
'    Dim local_xxxx As XXXX
'    Set local_xxxx = New XXXX
'
'    '--- Act ---
'    Dim actual_val as Variant
'    actual_val = ...
'
'    '--- Assert ---
'    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
'
'    Assert.Equals 99, actual_val
'
'    Assert.ErrorNotRaised 0, Err.Number, err.source, Err.Description
'End Sub
' =============================================================================

Private TUtl As New UnitTestUtils

Public Sub Test_DiffStringArray_置換なし_正しく差分の結果が返る(ByVal Assert As UnitTestAssert)
    'On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String
    
    ReDim old_arr(0 To 5)
    old_arr(0) = "aaa"
    old_arr(1) = "bbb"
    old_arr(2) = "ccc"
    old_arr(3) = "aaa"
    old_arr(4) = "bbb"
    old_arr(5) = "ccc"
    
    ReDim new_arr(0 To 5)
    new_arr(0) = "bbb"
    new_arr(1) = "ccc"
    new_arr(2) = "ddd"
    new_arr(3) = "bbb"
    new_arr(4) = "ccc"
    new_arr(5) = "ddd"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr, EnableReplaceType:=False)
     
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 7, UBound(old_arr)
    If UBound(old_arr) <> 7 Then Exit Sub
    
    Assert.EqualsNumeric 7, UBound(new_arr)
    If UBound(new_arr) <> 7 Then Exit Sub
    
    Assert.EqualsNumeric 7, UBound(type_arr)
    If UBound(type_arr) <> 7 Then Exit Sub
    
    Assert.Equals "aaa", old_arr(0)
    Assert.Equals "", new_arr(0)
    Assert.Equals "DEL", type_arr(0)
    
    Assert.Equals "bbb", old_arr(1)
    Assert.Equals "bbb", new_arr(1)
    Assert.Equals "", type_arr(1)
    
    Assert.Equals "ccc", old_arr(2)
    Assert.Equals "ccc", new_arr(2)
    Assert.Equals "", type_arr(2)
    
    Assert.Equals "aaa", old_arr(3)
    Assert.Equals "", new_arr(3)
    Assert.Equals "DEL", type_arr(3)
    
    Assert.Equals "", old_arr(4)
    Assert.Equals "ddd", new_arr(4)
    Assert.Equals "ADD", type_arr(4)
    
    Assert.Equals "bbb", old_arr(5)
    Assert.Equals "bbb", new_arr(5)
    Assert.Equals "", type_arr(5)
    
    Assert.Equals "ccc", old_arr(6)
    Assert.Equals "ccc", new_arr(6)
    Assert.Equals "", type_arr(6)
    
    Assert.Equals "", old_arr(7)
    Assert.Equals "ddd", new_arr(7)
    Assert.Equals "ADD", type_arr(7)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub

Public Sub Test_DiffStringArray_置換あり_正しく差分の結果が返る(ByVal Assert As UnitTestAssert)
    'On Error Resume Next

    '--- Arrange ---
    Dim old_arr() As String, new_arr() As String, type_arr() As String
    
    ReDim old_arr(0 To 5)
    old_arr(0) = "aaa"
    old_arr(1) = "bbb"
    old_arr(2) = "ccc"
    old_arr(3) = "aaa"
    old_arr(4) = "bbb"
    old_arr(5) = "ccc"
    
    ReDim new_arr(0 To 5)
    new_arr(0) = "bbb"
    new_arr(1) = "ccc"
    new_arr(2) = "ddd"
    new_arr(3) = "bbb"
    new_arr(4) = "ccc"
    new_arr(5) = "ddd"

    '--- Act ---
    Call DiffStringArray(old_arr, new_arr, type_arr)
     
    '--- Assert ---
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    
    Assert.EqualsNumeric 6, UBound(old_arr)
    If UBound(old_arr) <> 6 Then Exit Sub
    
    Assert.EqualsNumeric 6, UBound(new_arr)
    If UBound(new_arr) <> 6 Then Exit Sub
    
    Assert.EqualsNumeric 6, UBound(type_arr)
    If UBound(type_arr) <> 6 Then Exit Sub
    
    Assert.Equals "aaa", old_arr(0)
    Assert.Equals "", new_arr(0)
    Assert.Equals "DEL", type_arr(0)
    
    Assert.Equals "bbb", old_arr(1)
    Assert.Equals "bbb", new_arr(1)
    Assert.Equals "", type_arr(1)
    
    Assert.Equals "ccc", old_arr(2)
    Assert.Equals "ccc", new_arr(2)
    Assert.Equals "", type_arr(2)
    
    Assert.Equals "aaa", old_arr(3)
    Assert.Equals "ddd", new_arr(3)
    Assert.Equals "MOD", type_arr(3)
    
    Assert.Equals "bbb", old_arr(4)
    Assert.Equals "bbb", new_arr(4)
    Assert.Equals "", type_arr(4)
    
    Assert.Equals "ccc", old_arr(5)
    Assert.Equals "ccc", new_arr(5)
    Assert.Equals "", type_arr(5)
    
    Assert.Equals "", old_arr(6)
    Assert.Equals "ddd", new_arr(6)
    Assert.Equals "ADD", type_arr(6)
    
    Assert.ErrorNotRaised 0, Err.Number, Err.Source, Err.Description
End Sub
