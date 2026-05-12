VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pArray() As Variant

Public Property Get LowerBound() As Long
    LowerBound = LBound(pArray)
End Property

Public Property Get UpperBound() As Long
    UpperBound = UBound(pArray)
End Property

Public Function Item(ByVal ItemIndex As Long) As Variant
    If IsObject(pArray(ItemIndex)) Then
        Set Item = pArray(ItemIndex)
    Else
        Item = pArray(ItemIndex)
    End If
End Function

'Private Sub Class_Terminate()
'    On Error Resume Next
'    Debug.Print "Terminate: ArrayObject(0x" & Hex(ObjPtr(Me)) & ")"
'
'    Dim idx As Long
'    For idx = LBound(pArray) To UBound(pArray)
'        pArray(idx) = Empty
'    Next idx
'End Sub

Public Sub ReDimArray(ByVal LowerBound As Long, ByVal UpperBonnd As Long, Optional ByVal PreserveItems As Boolean = False)
    If PreserveItems Then
        ReDim Preserve pArray(LowerBound To UpperBonnd)
    Else
        ReDim pArray(LowerBound To UpperBonnd)
    End If
End Sub

Public Function Update(ByVal ItemIndex As Long, ByVal ItemObject As Variant) As Variant
    If IsObject(ItemObject) Then
        Set pArray(ItemIndex) = ItemObject
    Else
        pArray(ItemIndex) = ItemObject
    End If
End Function
