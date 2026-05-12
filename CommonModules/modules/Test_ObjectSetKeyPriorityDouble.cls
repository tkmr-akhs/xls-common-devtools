VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Test_ObjectSetKeyPriorityDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ObjectSet のキー生成優先順位を確認するためのテスト ダブルです。
'!
' #############################################################################

Implements IDuplicateCheckable
Implements IEquatable

Private pDuplicateKey As String
Private pIdentityKey As String

'* 重複判定キー。
Public Property Get DuplicateKey() As String
    DuplicateKey = pDuplicateKey
End Property

'* 重複判定キー。
Public Property Let DuplicateKey(ByVal NewValue As String)
    pDuplicateKey = NewValue
End Property

'* 同一性判定キー。
Public Property Get IdentityKey() As String
    IdentityKey = pIdentityKey
End Property

'* 同一性判定キー。
Public Property Let IdentityKey(ByVal NewValue As String)
    pIdentityKey = NewValue
End Property

Private Function IDuplicateCheckable_GetKey() As String
    IDuplicateCheckable_GetKey = pDuplicateKey
End Function

Private Function IDuplicateCheckable_IsDuplicateOf(ByVal OtherObject As Object) As Boolean
    If TypeName(OtherObject) <> TypeName(Me) Then Exit Function

    Dim other_item As Test_ObjectSetKeyPriorityDouble
    Set other_item = OtherObject

    IDuplicateCheckable_IsDuplicateOf = (pDuplicateKey = other_item.DuplicateKey)
End Function

Private Function IEquatable_Equals(ByVal OtherObject As Object) As Boolean
    If TypeName(OtherObject) <> TypeName(Me) Then Exit Function

    Dim other_item As Test_ObjectSetKeyPriorityDouble
    Set other_item = OtherObject

    IEquatable_Equals = (pIdentityKey = other_item.IdentityKey)
End Function

Private Function IEquatable_GetIdentityString() As String
    IEquatable_GetIdentityString = pIdentityKey
End Function
