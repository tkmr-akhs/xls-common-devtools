VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ユニット テストのユーティリティ クラスです。
'!
' #############################################################################

'* 指定された引数群から対応する値を取得します。
'*
'* @param Values 引数をキーにして値を格納している辞書。
'* @param Args 引数。
'*
'* @details
'* 指定された引数群から値 (スタブ目的の戻り値、スパイ目的の結果値) を取得します。
Public Function GetValue(ByVal Values As Dictionary, ParamArray Args() As Variant) As Variant
    Dim key_str As String
    key_str = pGetKey(Args)
    
    If Not Values.Exists(key_str) Then
        Err.Raise vbObjectError + 1, "Class UnitTestUtils", "指定された引数に対応する値が存在しません。(key: " & key_str & ")"
        Exit Function
    End If
    
    If IsObject(Values(key_str)) Then
        Set GetValue = Values(key_str)
    Else
        GetValue = Values(key_str)
    End If
End Function

'* 指定された引数群に対応する値が存在するかを判断します。
'*
'* @param Values 引数をキーにして値を格納している辞書。
'* @param Args 引数。
'* @return 対応する値が存在するか。
'*
'* @details
'* 指定された引数群に対応する値が存在するかを判断します。
Public Function HasValue(ByVal Values As Dictionary, ParamArray Args() As Variant) As Boolean
    If Values.Exists(pGetKey(Args)) Then
        HasValue = True
    Else
        HasValue = False
    End If
End Function

'* 指定された引数群に対する値を保存します。
'*
'* @param Values 引数をキーにして値を格納している辞書。
'* @param Value 保存する値。
'* @param Args 引数。
'*
'* @details
'* 指定された引数群に対応する値 (スタブ目的の戻り値、スパイ目的の結果値) を保存します。
Public Sub SetValue(ByVal Values As Dictionary, ByRef Value As Variant, ParamArray Args() As Variant)
    If IsObject(Value) Then
        Set Values.Item(pGetKey(Args)) = Value
    Else
        Values.Item(pGetKey(Args)) = Value
    End If
End Sub

Private Function pGetKey(ByVal Args As Variant) As String
    If IsEmpty(Args) Then Exit Function
    If UBound(Args) < LBound(Args) Then Exit Function
    
    Dim result_value As String
    result_value = pGetKeyCore(Args(LBound(Args)))
    
    Dim item_idx As Long
    For item_idx = LBound(Args) + 1 To UBound(Args)
        result_value = result_value & "|" & pGetKeyCore(Args(item_idx))
    Next item_idx
    
    pGetKey = result_value
End Function

Private Function pGetKeyCore(ByVal ArgItem As Variant) As String
    If IsObject(ArgItem) Then
        If ArgItem Is Nothing Then
            pGetKeyCore = "<Nothing>"
        ElseIf TypeOf ArgItem Is IEquatable Then
            Dim eq_item As IEquatable
            Set eq_item = ArgItem
            pGetKeyCore = eq_item.GetIdentityString() & "<" & TypeName(ArgItem) & ">"
        Else
            pGetKeyCore = Hex(ObjPtr(ArgItem)) & "<" & TypeName(ArgItem) & ">"
        End If
    Else
        pGetKeyCore = pGetPrimitiveKey(ArgItem)
    End If
End Function

Private Function pGetPrimitiveKey(ByVal ArgItem As Variant) As String
    If IsError(ArgItem) Then
        pGetPrimitiveKey = "CVErr(" & CStr(CLng(ArgItem)) & ")<" & TypeName(ArgItem) & ">"
    ElseIf IsNull(ArgItem) Then
        pGetPrimitiveKey = "Null<" & TypeName(ArgItem) & ">"
    ElseIf IsEmpty(ArgItem) Then
        pGetPrimitiveKey = "<Empty>"
    Else
        pGetPrimitiveKey = pEscapeSpecial(CStr(ArgItem)) & "<" & TypeName(ArgItem) & ">"
    End If
End Function

Private Function pEscapeSpecial(ByVal Expression As String) As String
    pEscapeSpecial = Replace(Replace(Replace(Replace(Expression, "\", "\\"), "|", "\|"), "<", "\<"), ">", "\>")
End Function
