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

Public Sub Test_RunAllTest_MultilineSubDeclaration_IsDiscovered( _
        ByVal Assert As UnitTestAssert)

    Assert.IsTrue True
End Sub
