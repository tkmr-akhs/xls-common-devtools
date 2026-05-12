Attribute VB_Name = "Test_ProgressStatus"
Option Explicit

' #############################################################################
'!
'! @brief
'! ProgressStatus クラスのユニット テストです。
'!
' #############################################################################

Public Sub Test_ClassTerminate_ObjectDisposed_StatusBarReset(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Application.StatusBar = "test progress"

    Dim prog_stat As ProgressStatus
    Set prog_stat = New ProgressStatus
    prog_stat.TaskName = "test"
    Call prog_stat.Start

    ' Act
    Set prog_stat = Nothing

    ' Assert
    Dim actual_value As Boolean
    actual_value = CBool(Application.StatusBar)

    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then
        Application.StatusBar = False
        Exit Sub
    End If
    Assert.Equals False, actual_value

    Application.StatusBar = False
End Sub
