Attribute VB_Name = "Test_Lib_FileSystem"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Lib_FileSystem モジュールのユニット テストです。
'!
' #############################################################################

Public Sub Test_InitializeFileSystemService_NotInitialized_InitializesFileSystemService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set FsSrv = Nothing

    ' Act
    Call InitializeFileSystemService

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "FileSystemService", FsSrv
End Sub
