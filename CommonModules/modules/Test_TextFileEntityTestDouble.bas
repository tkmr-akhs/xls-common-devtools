Attribute VB_Name = "Test_TextFileEntityTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! TextFileEntityTestDouble クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private TUtl As New UnitTestUtils

' ----------------------------------------------------------------------------
' Initialize
' ----------------------------------------------------------------------------

Public Sub Test_Initialize_WithFilePath_StoresResult(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    
    Dim test_path As String
    test_path = "C:\Temp\test.txt"
    
    ' Act
    Call file_stub.Initialize(test_path)
    
    ' Assert
    ' Initialize_Results に "test_path" をキーにして True が入っているか
    Assert.Equals True, TUtl.GetValue(file_stub.Initialize_Results, test_path)
End Sub

' ----------------------------------------------------------------------------
' OpenFile
' ----------------------------------------------------------------------------

Public Sub Test_OpenFile_DefaultParams_RecordsCall(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Assert.Equals False, file_stub.IsOpen
    
    ' Act
    file_stub.OpenFile
    
    ' Assert
    Assert.Equals True, TUtl.GetValue(file_stub.OpenFile_Results, False, False, False, False, False)
    Assert.Equals True, file_stub.AsRead
    Assert.Equals False, file_stub.AsWrite
    Assert.Equals False, file_stub.AsAppend
    Assert.Equals False, file_stub.GetReadLock
    Assert.Equals False, file_stub.GetWriteLock
    Assert.Equals True, file_stub.IsOpen
End Sub

Public Sub Test_OpenFile_AsWriteTrue_RecordsCall(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Assert.Equals False, file_stub.IsOpen
    
    ' Act
    Call file_stub.OpenFile(AsWrite:=True, AsAppend:=False, GetReadLock:=False, GetWriteLock:=True, Force:=True)
    
    ' Assert
    Assert.Equals True, TUtl.GetValue(file_stub.OpenFile_Results, True, False, False, True, True)
    Assert.Equals False, file_stub.AsRead
    Assert.Equals True, file_stub.AsWrite
    Assert.Equals False, file_stub.AsAppend
    Assert.Equals False, file_stub.GetReadLock
    Assert.Equals True, file_stub.GetWriteLock
    Assert.Equals True, file_stub.IsOpen
End Sub

Public Sub Test_OpenFile_AsAppendTrue_RecordsCall(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Assert.Equals False, file_stub.IsOpen
    
    ' Act
    Call file_stub.OpenFile(AsWrite:=True, AsAppend:=True, GetReadLock:=True, GetWriteLock:=False)
    
    ' Assert
    Assert.Equals True, TUtl.GetValue(file_stub.OpenFile_Results, True, True, True, False, False)
    Assert.Equals False, file_stub.AsRead
    Assert.Equals True, file_stub.AsWrite
    Assert.Equals True, file_stub.AsAppend
    Assert.Equals True, file_stub.GetReadLock
    Assert.Equals False, file_stub.GetWriteLock
    Assert.Equals True, file_stub.IsOpen
End Sub

' ----------------------------------------------------------------------------
' ReadLine/IsEndOfFile
' ----------------------------------------------------------------------------

Public Sub Test_ReadLine_WithValuesFromDictionary_ReturnsCorrectStrings(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    
    Call file_stub.ReadLine_Values.Add(0, "First line")
    Call file_stub.ReadLine_Values.Add(1, "Second line")
    Call file_stub.OpenFile
    
    ' Act/Assert
    Dim actual_str As String
    actual_str = file_stub.ReadLine
    
    Assert.Equals "First line", actual_str
    Assert.EqualsNumeric 1, file_stub.ReadCount
    Assert.Equals False, file_stub.IsEndOfFile
    
    ' もう一回読む
    actual_str = file_stub.ReadLine
    
    Assert.Equals "Second line", actual_str
    Assert.EqualsNumeric 2, file_stub.ReadCount
    Assert.Equals True, file_stub.IsEndOfFile
End Sub

Public Sub Test_ReadLine_MissingReadCount_RaisesWithoutCreatingKey(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.OpenFile

    ' Act
    Dim actual_str As String
    actual_str = file_stub.ReadLine

    ' Assert
    If Not Assert.ErrorRaised(vbObjectError + 1, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "Class TextFileEntityTestDouble", Err.Source
    Assert.EqualsNumeric 0, file_stub.ReadCount
    Assert.IsFalse file_stub.ReadLine_Values.Exists(0)
End Sub
' ----------------------------------------------------------------------------
' WriteLine
' ----------------------------------------------------------------------------

Public Sub Test_WriteLine_CalledTwice_IncrementsCount(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    
    ' Act
    Call file_stub.WriteLine("First line")
    Call file_stub.WriteLine("Second line")
    
    ' Assert
    Assert.EqualsNumeric 2, file_stub.WriteCount
    Assert.Equals "First line", file_stub.WriteLine_Results(0)
    Assert.Equals "Second line", file_stub.WriteLine_Results(1)
End Sub

Public Sub Test_WriteLine_NoCalls_WriteCountIsZero(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    
    ' Act
    
    ' Assert
    Assert.EqualsNumeric 0, file_stub.WriteCount
End Sub

' ----------------------------------------------------------------------------
' CloseFile
' ----------------------------------------------------------------------------

Public Sub Test_CloseFile_AfterSomeReadWrite_ResetsReadCounter(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    Call file_stub.ReadLine_Values.Add(0, "test")
    
    Call file_stub.OpenFile
    Dim dummy As String
    dummy = file_stub.ReadLine
    
    Assert.EqualsNumeric 1, file_stub.ReadCount
    Assert.Equals True, file_stub.IsOpen
    
    ' Act
    file_stub.CloseFile
    
    ' Assert
    ' ReadCount が0 にリセットされるか
    Assert.EqualsNumeric 0, file_stub.ReadCount
    Assert.Equals False, file_stub.IsOpen
End Sub

Public Sub Test_CloseFile_AfterSomeReadWrite_ResetsWriteCounter(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim file_stub As TextFileEntityTestDouble
    Set file_stub = New TextFileEntityTestDouble
    
    Call file_stub.OpenFile(AsWrite:=True)
    Call file_stub.WriteLine("ABC")
    
    Assert.EqualsNumeric 1, file_stub.WriteCount
    Assert.EqualsNumeric 1, file_stub.WriteLine_Results.Count
    Assert.Equals True, file_stub.IsOpen
    
    ' Act
    file_stub.CloseFile
    
    ' Assert
    ' WriteCount がリセットされないか
    Assert.EqualsNumeric 1, file_stub.WriteCount
    Assert.EqualsNumeric 1, file_stub.WriteLine_Results.Count
    Assert.Equals False, file_stub.IsOpen
End Sub


