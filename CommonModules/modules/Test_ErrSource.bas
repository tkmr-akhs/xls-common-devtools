Attribute VB_Name = "Test_ErrSource"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Err.Raise の Source に記載されたクラス名を確認するユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Public Sub Test_ErrRaiseSource_ClassNameMatchesModuleName(ByVal Assert As UnitTestAssert)
    Dim mismatch_list As ObjectList
    Set mismatch_list = New ObjectList

    Dim vb_proj As Variant
    Set vb_proj = Application.VBE.ActiveVBProject

    Dim vb_comp As Variant
    For Each vb_comp In vb_proj.VBComponents
        If vb_comp.Type = 2 Then
            Call pCollectMismatch(mismatch_list, vb_comp)
        End If
    Next vb_comp

    Assert.Equals "", JoinStringList(mismatch_list)
End Sub

Private Sub pCollectMismatch(ByVal MismatchList As ObjectList, ByVal VbComponent As Variant)
    Dim module_name As String
    module_name = VbComponent.Name

    Dim code_module As Variant
    Set code_module = VbComponent.CodeModule

    Dim line_idx As Long
    For line_idx = 1 To code_module.CountOfLines
        Dim code_line As String
        code_line = code_module.Lines(line_idx, 1)

        Dim source_class_name As String
        source_class_name = pExtractClassSourceName(code_line)

        If source_class_name <> "" And source_class_name <> module_name Then
            Call MismatchList.Add(module_name & ":" & line_idx & " Source=Class " & source_class_name)
        End If
    Next line_idx
End Sub

Private Function pExtractClassSourceName(ByVal CodeLine As String) As String
    If InStr(1, CodeLine, "Err.Raise", vbBinaryCompare) = 0 Then Exit Function

    Dim marker_text As String
    marker_text = """Class "

    Dim start_pos As Long
    start_pos = InStr(1, CodeLine, marker_text, vbBinaryCompare)
    If start_pos = 0 Then Exit Function

    start_pos = start_pos + Len(marker_text)

    Dim end_pos As Long
    end_pos = InStr(start_pos, CodeLine, """", vbBinaryCompare)
    If end_pos = 0 Then Exit Function

    pExtractClassSourceName = Mid(CodeLine, start_pos, end_pos - start_pos)
End Function
