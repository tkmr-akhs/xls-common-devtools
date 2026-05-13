VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnitTestAssert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ユニット テストの assert です。
'!
' #############################################################################

Private pCallCount As Long
Private pIsFailed As Boolean
Private pResultMessage As String

'* assert と異なっていたか
Public Property Get IsFailed() As Boolean
    IsFailed = pIsFailed
End Property

'* assert と異なっていた場合の詳細
Public Property Get ResultMessage() As String
    ResultMessage = pResultMessage
End Property

Private Sub Class_Initialize()
    pCallCount = 0
    pIsFailed = False
    pResultMessage = ""
End Sub

'* 予測された値と実際の値が等しい。
'*
'* @param ExpectedValue 予測された値。
'* @param ActualValue 実際の値。
'*
'* @details
'* 予測された値と実際の値が等しいことを示します。
'* 値の比較だけでなく、型の比較も行います。
Public Sub Equals(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean
    
    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, False)
    
    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

'* 予測された値と実際の値が等くない。
'*
'* @param ExpectedValue 予測された値。
'* @param ActualValue 実際の値。
'*
'* @details
'* 予測された値と実際の値が異なることを示します。
'* 値の比較だけでなく、型の比較も行います。
Public Sub NotEquals(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean
    
    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, False)
    
    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

'* 予測された数値と実際の数値が等しい。
'*
'* @param ExpectedValue 予測された値。
'* @param ActualValue 実際の値。
'*
'* @details
'* 予測された数値と実際の数値が等しいことを示します。
'* 数値であれば、型はチェックしません。
'* 注意: Currency と Decimal は現時点では型無視の数値比較対象に含めません。
Public Sub EqualsNumeric(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean
    
    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, True)
    
    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

'* 予測された数値と実際の数値が等くない。
'*
'* @param ExpectedValue 予測された値。
'* @param ActualValue 実際の値。
'*
'* @details
'* 予測された数値と実際の数値が等しくないことを示します。
'* 数値であれば、型はチェックしません。
'* 注意: Currency と Decimal は現時点では型無視の数値比較対象に含めません。
Public Sub NotEqualsNumeric(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_equal As Boolean
    
    Call pEqualsCore(is_equal, expected_str, actual_str, ExpectedValue, ActualValue, True)
    
    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

Private Sub pEqualsCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant, _
        ByVal IgnoreNumericType As Boolean)
    
    Dim actual_is_object As Boolean
    Dim expected_is_object As Boolean
    
    expected_is_object = IsObject(ExpectedValue)
    actual_is_object = IsObject(ActualValue)
    
    If expected_is_object And Not actual_is_object _
            Or Not expected_is_object And actual_is_object Then
        
        ExpectedValueString = "Entity of " & TypeName(ExpectedValue)
        ActualValueString = "Entity of " & TypeName(ActualValue)
        CheckResult = False
        Exit Sub
    End If
    
    ' ここまでで、片方がプリミティブ型で、片方がオブジェクト型の場合は早期リターンされている
    
    If actual_is_object Then
        ' 両方オブジェクト型
    
        Dim actual_is_nothing As Boolean
        Dim expected_is_nothing As Boolean
        
        actual_is_nothing = (ActualValue Is Nothing)
        expected_is_nothing = (ExpectedValue Is Nothing)
        
        If actual_is_nothing Then
            ActualValueString = "Nothing"
        Else
            ActualValueString = "Entity of " & TypeName(ActualValue)
        End If
        
        If expected_is_nothing Then
            ExpectedValueString = "Nothing"
        Else
            ExpectedValueString = "Entity of " & TypeName(ExpectedValue)
        End If
        
        If expected_is_nothing Or actual_is_nothing Then
            If expected_is_nothing And actual_is_nothing Then
                CheckResult = True
            Else
                CheckResult = False
            End If
            Exit Sub
        End If
        
        ' ここまでで、どちらかが、あるいはどちらも Nothing の場合は早期リターンされている
    
        If Not ExpectedValue Is ActualValue Then
            CheckResult = False
        Else
            CheckResult = True
        End If
    Else
        ' 両方プリミティブ型
        If IgnoreNumericType Then
            Call pEqualsPrimitiveNumericCore(CheckResult, ExpectedValueString, ActualValueString, ExpectedValue, ActualValue)
        Else
            Call pEqualsPrimitiveCore(CheckResult, ExpectedValueString, ActualValueString, ExpectedValue, ActualValue)
        End If
    End If
End Sub

Private Sub pEqualsPrimitiveCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant)
    
    
    Dim expect_type As String
    expect_type = TypeName(ExpectedValue)
    
    Dim actual_type As String
    actual_type = TypeName(ActualValue)
    
    ExpectedValueString = pPrimitiveValueString(ExpectedValue)
    ActualValueString = pPrimitiveValueString(ActualValue)
    
    If pIsSpecialPrimitive(ExpectedValue) Or pIsSpecialPrimitive(ActualValue) Then
        CheckResult = pSpecialPrimitivesEqual(ExpectedValue, ActualValue)
        Exit Sub
    End If
    
    If expect_type = "String" And Not actual_type = "String" _
            Or Not expect_type = "String" And actual_type = "String" Then
        CheckResult = False
        Exit Sub
    End If
    
    If ExpectedValue = ActualValue Then
        If TypeName(ExpectedValue) <> TypeName(ActualValue) Then
            CheckResult = False
        Else
            CheckResult = True
        End If
    Else
        CheckResult = False
    End If
End Sub

Private Sub pEqualsPrimitiveNumericCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedValue As Variant, _
        ByVal ActualValue As Variant)
    
    ExpectedValueString = pPrimitiveValueString(ExpectedValue)
    ActualValueString = pPrimitiveValueString(ActualValue)
    
    If pIsSpecialPrimitive(ExpectedValue) Or pIsSpecialPrimitive(ActualValue) Then
        CheckResult = False
        Exit Sub
    End If
    
    If ExpectedValue = ActualValue Then
        If pIsNumericType(TypeName(ExpectedValue)) And pIsNumericType(TypeName(ActualValue)) Then
            CheckResult = True
        Else
            CheckResult = False
        End If
    Else
        CheckResult = False
    End If
End Sub

Private Function pPrimitiveValueString(ByVal TargetValue As Variant) As String
    If VBA.IsError(TargetValue) Then
        pPrimitiveValueString = "CVErr(" & CStr(CLng(TargetValue)) & ")<" & TypeName(TargetValue) & ">"
    ElseIf VBA.IsNull(TargetValue) Then
        pPrimitiveValueString = "Null<" & TypeName(TargetValue) & ">"
    ElseIf VBA.IsEmpty(TargetValue) Then
        pPrimitiveValueString = "<Empty>"
    Else
        pPrimitiveValueString = CStr(TargetValue) & "<" & TypeName(TargetValue) & ">"
    End If
End Function

Private Function pIsSpecialPrimitive(ByVal TargetValue As Variant) As Boolean
    pIsSpecialPrimitive = VBA.IsError(TargetValue) Or VBA.IsNull(TargetValue) Or VBA.IsEmpty(TargetValue)
End Function

Private Function pSpecialPrimitivesEqual(ByVal ExpectedValue As Variant, ByVal ActualValue As Variant) As Boolean
    If VBA.IsError(ExpectedValue) Or VBA.IsError(ActualValue) Then
        If VBA.IsError(ExpectedValue) And VBA.IsError(ActualValue) Then
            pSpecialPrimitivesEqual = CLng(ExpectedValue) = CLng(ActualValue)
        Else
            pSpecialPrimitivesEqual = False
        End If
    ElseIf VBA.IsNull(ExpectedValue) Or VBA.IsNull(ActualValue) Then
        pSpecialPrimitivesEqual = VBA.IsNull(ExpectedValue) And VBA.IsNull(ActualValue)
    ElseIf VBA.IsEmpty(ExpectedValue) Or VBA.IsEmpty(ActualValue) Then
        pSpecialPrimitivesEqual = VBA.IsEmpty(ExpectedValue) And VBA.IsEmpty(ActualValue)
    Else
        pSpecialPrimitivesEqual = False
    End If
End Function

Private Function pIsNumericType(ByVal TypeNameString As String) As Boolean
    Select Case TypeNameString
     Case "Byte"
        pIsNumericType = True
     Case "Short"
        pIsNumericType = True
     Case "Integer"
        pIsNumericType = True
     Case "Long"
        pIsNumericType = True
     Case "LongLong"
        pIsNumericType = True
     Case "Single"
        pIsNumericType = True
     Case "Double"
        pIsNumericType = True
     Case Else
        pIsNumericType = False
    End Select
End Function

'* True が予測される。
'*
'* @param ActualValue 実際の値。
'*
'* @details
'* 結果が真であることを示します。
Public Sub IsTrue(ByVal ActualValue As Variant)
    Call Equals(True, ActualValue)
End Sub

'* False が予測される。
'*
'* @param ActualValue 実際の値。
'*
'* @details
'* 結果が偽であることを示します。
Public Sub IsFalse(ByVal ActualValue As Variant)
    Call Equals(False, ActualValue)
End Sub

'* 予測された型と実際の型が等しい。
'*
'* @param ExpectedTypeName 予測された型名の文字列。
'* @param ActualValue 実際の値。
'*
'* @details
'* 予測された型と実際の型が等しいことを示します。
'* 型が一致するかをチェックするのであり、キャスト可能であるかをチェックするのではありません。
Public Sub IsTypeOf(ByVal ExpectedTypeName As String, ByVal ActualValue As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    If TypeName(ActualValue) <> ExpectedTypeName Then
        Call pSetResultMessage("<" & ExpectedTypeName & ">", "<" & TypeName(ActualValue) & ">")
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

'* Nothing が予測される。
'*
'* @param ActualValue 実際の値。
'*
'* @details
'* Nothing が予測される。
'* IsTypeOf("Nothing", ActualValue) と同じです。
Public Sub IsNothing(ByVal ActualValue As Variant)
    Call IsTypeOf("Nothing", ActualValue)
End Sub

'* Empty が予測される。
'*
'* @param ActualValue 実際の値。
'*
'* @details
'* Empty が予測される。
'* IsTypeOf("Empty", ActualValue) と同じです。
Public Sub IsEmpty(ByVal ActualValue As Variant)
    Call IsTypeOf("Empty", ActualValue)
End Sub

'* 指定されたエラーが発生することが予測される。
'*
'* @param ExpectedCode 予測されるエラーコード。0 を渡した場合、すべてのエラーが対象となります
'* @param ActualCodeCode 実際のエラーコード。Err.Number を渡してください。
'* @param ErrorSource 実際のエラー発生元。Err.Source を渡してください。
'* @param ErrorDescription 実際のエラー詳細。Err.Description を渡してください。
'* @return エラーが発生したかどうか。
'*
'* @details
'* 指定されたエラーが発生することが予測されることを示します。
'* ExpectedCode として 0 が指定された場合、なんらかのエラーが発生することを示します。
'* 0 以外が指定された場合、そのエラーが発生することを示します (他のエラーが発生しても無視します)。
Public Function ErrorRaised(ByVal ExpectedCode As Long, ByVal ActualCode As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String) As Boolean
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Function
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_error As Boolean
    
    Call pErrorRaisedCore(is_error, expected_str, actual_str, ExpectedCode, ActualCode, ErrorSource, ErrorDescription, False)
    
    If is_error Then
        ErrorRaised = True
    Else
        Call pSetResultMessage(expected_str, actual_str)
        ErrorRaised = False
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Function

'* 指定されたエラーが発生しないことが予測される。
'*
'* @param ExpectedCode 予測されるエラーコード。0 を渡した場合、すべてのエラーが対象となります
'* @param ActualCodeCode 実際のエラーコード。Err.Number を渡してください。
'* @param ErrorSource 実際のエラー発生元。Err.Source を渡してください。
'* @param ErrorDescription 実際のエラー詳細。Err.Description を渡してください。
'* @return エラーが発生しなかったかどうか。
'*
'* @details
'* 指定されたエラーが発生しないことが予測されることを示します。
'* ExpectedCode として 0 が指定された場合、どのようなエラーも発生しないことを示します。
'* 0 以外が指定された場合、そのエラーが発生しない (他のエラーは発生するかもしれない) ことを示します。
Public Function ErrorNotRaised(ByVal ExpectedCode As Long, ByVal ActualCode As Long, ByVal ErrorSource As String, ByVal ErrorDescription As String) As Boolean
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Function
    
    Dim expected_str As String
    Dim actual_str As String
    Dim is_error As Boolean
    
    Call pErrorRaisedCore(is_error, expected_str, actual_str, ExpectedCode, ActualCode, ErrorSource, ErrorDescription, True)
    
    If is_error Then
        Call pSetResultMessage(expected_str, actual_str)
        ErrorNotRaised = False
    Else
        ErrorNotRaised = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Function

Private Sub pErrorRaisedCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByVal ExpectedCode As Variant, _
        ByVal ActualCode As Variant, _
        ByVal ErrorSource As String, _
        ByVal ErrorDescription As String, _
        ByVal AsNotRaised As Boolean)
    
    If ActualCode = 0 Then
        CheckResult = False
        If ExpectedCode = 0 Then
            ExpectedValueString = "Some error is raised."
            ActualValueString = "Any error is not raised."
            CheckResult = False
        Else
            ExpectedValueString = "The error [0x" & Hex(ExpectedCode) & "] is raised."
            ActualValueString = "Any error is not raised."
            CheckResult = False
        End If
    Else
        If ExpectedCode = 0 Then
            ' 予測コード 0 の時は何らかのエラーが挙がったら OK。
            If AsNotRaised Then
                ExpectedValueString = "Any error is not raised."
                ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
                CheckResult = True
            Else
                ExpectedValueString = "Some error is raised."
                ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
                CheckResult = True
            End If
        Else
            ExpectedValueString = "The error [0x" & Hex(ExpectedCode) & "] is raised."
            ActualValueString = "The error [0x" & Hex(ActualCode) & "]@<" & ErrorSource & "> """ & ErrorDescription & """ is raised."
            If ActualCode = ExpectedCode Then
                CheckResult = True
            Else
                CheckResult = False
            End If
        End If
    End If
End Sub

'* 予測された配列と実際の配列が等しい。
'*
'* @param ExpectedValue 予測された配列。
'* @param ActualValue 実際の配列。
'*
'* @details
'* 予測された配列と実際の配列が等しいことを示します。
'* 値の比較だけでなく、型の比較も行います。
Public Sub EqualsArray(ByVal ExpectedArray As Variant, ByVal ActualArray As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim additional_str As String
    Dim is_equal As Boolean
    
    Call pEqualsArrayCore(is_equal, expected_str, actual_str, additional_str, ExpectedArray, ActualArray, True)
    
    If Not is_equal Then
        Call pSetResultMessage(expected_str, actual_str, AdditionalInformation:=additional_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

'* 予測された配列と実際の配列が等しくない。
'*
'* @param ExpectedValue 予測された配列。
'* @param ActualValue 実際の配列。
'*
'* @details
'* 予測された配列と実際の配列が異なることを示します。
'* 値の比較だけでなく、型の比較も行います。
Public Sub NotEqualsArray(ByVal ExpectedArray As Variant, ByVal ActualArray As Variant)
    ' すでに失敗済みの時は処理をしない
    If pIsFailed Then Exit Sub
    
    Dim expected_str As String
    Dim actual_str As String
    Dim additional_str As String
    Dim is_equal As Boolean
    
    Call pEqualsArrayCore(is_equal, expected_str, actual_str, additional_str, ExpectedArray, ActualArray, True)
    
    If is_equal Then
        Call pSetResultMessage(expected_str, actual_str, AdditionalInformation:=additional_str)
        pIsFailed = True
    End If
    
    ' 常に行う処理
    pCallCount = pCallCount + 1
End Sub

Private Sub pEqualsArrayCore( _
        ByRef CheckResult As Boolean, _
        ByRef ExpectedValueString As String, _
        ByRef ActualValueString As String, _
        ByRef AdditionalInformation As String, _
        ByVal ExpectedArray As Variant, _
        ByVal ActualArray As Variant, _
        ByVal IgnoreNumericType As Boolean)
    
    ExpectedValueString = TypeName(ExpectedArray)
    ActualValueString = TypeName(ActualArray)
    
    ' 型が違う場合は失敗
    If ExpectedValueString <> ActualValueString Then
        CheckResult = False
        Exit Sub
    End If
    
    ' どちらかが配列でなければ失敗
    If Not IsArray(ExpectedArray) Or Not IsArray(ActualArray) Then
        CheckResult = False
        Exit Sub
    End If
    
    Dim expect_is_empty As Boolean
    expect_is_empty = IsEmptyArray(ExpectedArray)
    
    Dim actual_is_empty As Boolean
    actual_is_empty = IsEmptyArray(ActualArray)
    
    If expect_is_empty Or actual_is_empty Then
        CheckResult = expect_is_empty And actual_is_empty
        If Not CheckResult Then
            If expect_is_empty Then
                ExpectedValueString = "empty array"
            End If
            If actual_is_empty Then
                ActualValueString = "empty array"
            End If
        End If
        Exit Sub
    End If
    
    ' 配列の LBound および UBound を得る
    Dim expect_lb_arr() As Long
    Dim expect_ub_arr() As Long
    Call GetArrayBounds(expect_lb_arr, expect_ub_arr, ExpectedArray)
    
    Dim actual_lb_arr() As Long
    Dim actual_ub_arr() As Long
    Call GetArrayBounds(actual_lb_arr, actual_ub_arr, ActualArray)
    
    ' 配列の次元数を算出
    Dim expect_dimension As Long
    expect_dimension = UBound(expect_lb_arr) - LBound(expect_lb_arr) + 1
    
    Dim actual_dimension As Long
    actual_dimension = UBound(actual_lb_arr) - LBound(actual_lb_arr) + 1
    
    ' 配列の次元数を比較
    If expect_dimension <> actual_dimension Then
        ExpectedValueString = "dimension is " & expect_dimension
        ActualValueString = "dimension is " & actual_dimension
        CheckResult = False
        Exit Sub
    End If
    
    ' 次元ごとの LBound および UBound を比較
    Dim item_idx As Long
    For item_idx = LBound(expect_lb_arr) To UBound(expect_lb_arr)
        If expect_lb_arr(item_idx) <> actual_lb_arr(item_idx) Then
            ExpectedValueString = "lbound of dimension " & (item_idx + 1) & " is " & expect_lb_arr(item_idx)
            ActualValueString = "lbound of dimension " & (item_idx + 1) & " is " & actual_lb_arr(item_idx)
            CheckResult = False
            Exit Sub
        End If
        
        If expect_ub_arr(item_idx) <> actual_ub_arr(item_idx) Then
            ExpectedValueString = "ubound of dimension " & (item_idx + 1) & " is " & expect_ub_arr(item_idx)
            ActualValueString = "ubound of dimension " & (item_idx + 1) & " is " & actual_ub_arr(item_idx)
            CheckResult = False
            Exit Sub
        End If
    Next item_idx
    
    ' expect を 1 次元配列にする
    Dim expect_item As Variant
    Dim expect_arr() As Variant
    item_idx = 0
    For Each expect_item In ExpectedArray
        ReDim Preserve expect_arr(0 To item_idx)
        If IsObject(expect_item) Then
            Set expect_arr(item_idx) = expect_item
        Else
            expect_arr(item_idx) = expect_item
        End If
        item_idx = item_idx + 1
    Next expect_item
    
    ' すべての要素を Equals で比較
    ' 注意: 配列比較は型も含めて厳密一致とするため、IgnoreNumericType にかかわらず要素比較では数値型差を無視しない。
    Dim actual_item As Variant
    item_idx = 0
    For Each actual_item In ActualArray
        Dim check_result As Boolean
        Dim expect_str As String
        Dim actual_str As String
        Call pEqualsCore(check_result, expect_str, actual_str, expect_arr(item_idx), actual_item, False)
        If Not check_result Then
            ExpectedValueString = expect_str
            ActualValueString = actual_str
            AdditionalInformation = "@" & item_idx
            Exit Sub
        End If
        
        item_idx = item_idx + 1
    Next actual_item
    
    CheckResult = True
End Sub

'Private Function pEqualsMultidimensionalArrayCore( _
'        ByRef TargetArray() As Variant, _
'        ByRef CurrentDimension As Integer, _
'        ByRef MaxDimension As Integer, _
'        ByVal ExpectedArray As Variant, _
'        ByVal ActualArray As Variant, _
'        ByVal IgnoreNumericType As Boolean)
'
'    Dim expected_value_string As String
'    Dim actual_value_string As String
'    expected_value_string = TypeName(ExpectedArray)
'    actual_value_string = TypeName(ActualArray)
'
'    If expected_value_string <> actual_value_string Then
'        CheckResult = False
'        Exit Function
'    End If
'
'    If Not IsArray(ExpectedArray) Or Not IsArray(ActualArray) Then
'        CheckResult = False
'        Exit Function
'    End If
'
'    Dim expect_dimension As Long
'    expect_dimension = pGetArrayDimension(ExpectedArray)
'
'    Dim actual_dimension As Long
'    actual_dimension = pGetArrayDimension(ActualArray)
'
'    If expect_dimension <> actual_dimension Then
'        expected_value_string = "dimension is " & expect_dimension
'        actual_value_string = "dimension is " & actual_dimension
'        CheckResult = False
'        Exit Function
'    End If
'
'    'CheckResult = True
'End Function

Private Sub pSetResultMessage(ByVal ExpectedValueString As String, ByVal ActualValueString As String, Optional ByVal AdditionalInformation As String = "")
    If AdditionalInformation = "" Then
        pResultMessage = "[" & pCallCount & "] expected: " & ExpectedValueString & " | actual: " & ActualValueString
    Else
        pResultMessage = "[" & pCallCount & "] expected: " & ExpectedValueString & " | actual: " & ActualValueString & " | " & AdditionalInformation
    End If
    pIsFailed = True
End Sub
