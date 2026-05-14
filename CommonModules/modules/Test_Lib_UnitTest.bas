Attribute VB_Name = "Test_Lib_UnitTest"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Lib_UnitTest モジュールのユニット テストです。
'!
' #############################################################################

Rem Public Sub Test_RunAllTest_RemCommentDeclaration_IsIgnored(ByVal Assert As UnitTestAssert)

Public Sub Test_RunAllTest_RequiredNonAssertArg_IsIgnored( _
        ByVal Assert As UnitTestAssert, _
        ByVal RequiredValue As Long)

    Err.Raise vbObjectError + 1, "Test_Lib_UnitTest", "This fixture must not be discovered."
End Sub

Public Sub Test_RunAllTest_OptionalNonAssertArg_IsDiscovered( _
        ByVal Assert As UnitTestAssert, _
        Optional ByVal OptionalValue As Long = 0)

    Assert.Equals CLng(0), OptionalValue
End Sub

Public Sub Test_RunAllTest_MultilineSubDeclaration_IsDiscovered( _
        ByVal Assert As UnitTestAssert)

    Assert.IsTrue True
End Sub
