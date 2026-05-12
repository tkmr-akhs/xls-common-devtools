Attribute VB_Name = "Test_Lib_TextFile"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Lib_TextFile モジュールのユニット テストです。
'!
' #############################################################################

Public Sub Test_InitializeTextFileService_NotInitialized_InitializesTextFileService(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set TfSrv = Nothing

    ' Act
    Call InitializeTextFileService

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "TextFileService", TfSrv
End Sub
