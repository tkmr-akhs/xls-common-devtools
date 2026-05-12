VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFileServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! TextFileEntity を生成するサービスのテストダブルです。
'!
' #############################################################################

Implements ITextFileService

Private TUtl As UnitTestUtils

'* GetTextFileEntity の読み取り数→戻り値辞書。引数である FilePath をキーとして、返してほしい TextFileEntityTestDouble を戻り値として格納します。
Public GetTextFileEntity_Values As Dictionary

Private Sub Class_Initialize()
    Set TUtl = New UnitTestUtils

    Set GetTextFileEntity_Values = New Dictionary
End Sub

'* GetTextFileEntity のスタブ
'*
'* @details
'* Set GetTextFileEntity = TUtl.GetValue(GetTextFileEntity_Values, FilePath)
Public Function GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Set GetTextFileEntity = TUtl.GetValue(GetTextFileEntity_Values, FilePath)
End Function
Private Function ITextFileService_GetTextFileEntity(ByVal FilePath As String) As ITextFileEntity
    Set ITextFileService_GetTextFileEntity = GetTextFileEntity(FilePath)
End Function

