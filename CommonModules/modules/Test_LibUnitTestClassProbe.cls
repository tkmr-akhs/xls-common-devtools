VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_LibUnitTestClassProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Test_RunAllTest_ClassModuleDeclaration_IsIgnored(ByVal Assert As UnitTestAssert)
    Err.Raise vbObjectError + 1, "Class Test_LibUnitTestClassProbe", "Class module test candidate should be ignored."
End Sub
