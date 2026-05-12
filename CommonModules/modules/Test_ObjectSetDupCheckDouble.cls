VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ObjectSetDupCheckDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ObjectSet の IDuplicateCheckable 判定を確認するためのテスト ダブルです。
'!
' #############################################################################

Implements IDuplicateCheckable

Private pDuplicateKey As String

'* 重複判定キー。
Public Property Get DuplicateKey() As String
    DuplicateKey = pDuplicateKey
End Property

'* 重複判定キー。
Public Property Let DuplicateKey(ByVal NewValue As String)
    pDuplicateKey = NewValue
End Property

Private Function IDuplicateCheckable_GetKey() As String
    IDuplicateCheckable_GetKey = pDuplicateKey
End Function

Private Function IDuplicateCheckable_IsDuplicateOf(ByVal OtherObject As Object) As Boolean
    If TypeName(OtherObject) <> TypeName(Me) Then Exit Function

    Dim other_item As Test_ObjectSetDupCheckDouble
    Set other_item = OtherObject

    IDuplicateCheckable_IsDuplicateOf = (pDuplicateKey = other_item.DuplicateKey)
End Function
