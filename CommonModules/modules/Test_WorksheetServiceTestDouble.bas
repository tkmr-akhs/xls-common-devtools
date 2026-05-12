Attribute VB_Name = "Test_WorksheetServiceTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! WorksheetServiceTestDouble クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private TUtl As New UnitTestUtils

Public Sub Test_WriteCell_WithSimpleValues_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 2, Sheet:="TestSheet") ' Row=1,Col=2
    
    ' Act
    sheet_srv.WriteCell range_bounds, "Hello", "m/d", True, True, False, True
    
    ' Assert
    ' 1) コールが1件記録されているか
    Assert.EqualsNumeric 1, test_dbl.WriteCell_Results.Count
    
    ' 2) 格納されているキー(内部的にはpGetKey)は1件だけ
    Dim dict_key As Variant
    dict_key = test_dbl.WriteCell_Results.Keys()(0)
    
    Dim stored_value As Variant
    stored_value = test_dbl.WriteCell_Results.Items()(0)
    ' stored_value は Array(Expression, NumberFormat, AsFormula, ClearWhenEmpty, IgnoreEmpty, TypeConvert) の順
    
    Assert.Equals "Hello", stored_value(0)
    Assert.Equals "m/d", stored_value(1)
    Assert.IsTrue stored_value(2)          ' AsFormula = True
    Assert.IsTrue stored_value(3)          ' ClearWhenEmpty = True
    Assert.IsFalse stored_value(4)         ' IgnoreEmpty = False
    Assert.IsTrue stored_value(5)          ' TypeConvert = True
End Sub

Public Sub Test_WriteRange_With2DArray_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(2, 3, Sheet:="TestSheet")
    
    Dim arr_values(1 To 2, 1 To 2) As Variant
    arr_values(1, 1) = "A1"
    arr_values(1, 2) = "B1"
    arr_values(2, 1) = "A2"
    arr_values(2, 2) = "B2"
    
    ' Act
    sheet_srv.WriteRange range_bounds, arr_values
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.WriteRange_Results.Count
    
    Dim stored_value As Variant
    stored_value = test_dbl.WriteRange_Results.Items()(0) ' items(0) = Array(arr_values)
    Dim copied_arr As Variant
    copied_arr = stored_value(0)
    
    Assert.Equals "A1", copied_arr(1, 1)
    Assert.Equals "B2", copied_arr(2, 2)
End Sub

Public Sub Test_WriteArrayFormula_WithSimpleFormula_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 3, Sheet:="TestSheet")
    
    ' Act
    sheet_srv.WriteArrayFormula range_bounds, "=TRANSPOSE({1,2,3})"
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.WriteArrayFormula_Results.Count
    
    Dim stored_value As Variant
    stored_value = test_dbl.WriteArrayFormula_Results.Items()(0) ' array(FormulaArray)
    Assert.Equals "=TRANSPOSE({1,2,3})", stored_value(0)
End Sub

Public Sub Test_ReadCell_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(4, 1, Sheet:="TestSheet")
    
    Call TUtl.SetValue(test_dbl.ReadCell_Values, Array("StubValue", "m/d"), range_bounds, True)
    
    ' Act
    Dim actual_expr As String
    Dim actual_fmt As String
    sheet_srv.ReadCell range_bounds, actual_expr, actual_fmt, True

    ' Assert: 読み取った値がStubValueになるか
    Assert.Equals "StubValue", actual_expr
    Assert.Equals "m/d", actual_fmt
End Sub

Public Sub Test_CopyCell_WithAsValueTrue_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(1, 1, Sheet:="SrcSheet")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(2, 2, Sheet:="DstSheet")
    
    ' Act
    sheet_srv.CopyCell src_bounds, dst_bounds, True, True
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.CopyCell_Results.Count
    
    Dim stored_value As Variant
    stored_value = test_dbl.CopyCell_Results.Items()(0)
    ' stored_value = Array(SourceRangeBounds, AsValue, CopyNumberFormat)
    
    Dim as_value As Boolean: as_value = stored_value(1)
    Dim copy_num_fmt As Boolean: copy_num_fmt = stored_value(2)
    
    Assert.IsTrue as_value
    Assert.IsTrue copy_num_fmt
End Sub

Public Sub Test_CopyRange_WithDefaultOptions_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim src_bounds As WorksheetRangeBounds
    Set src_bounds = New_RangeBounds(1, 2, Sheet:="SrcSheet")
    
    Dim dst_bounds As WorksheetRangeBounds
    Set dst_bounds = New_RangeBounds(5, 5, Sheet:="DstSheet")
    
    ' Act
    sheet_srv.CopyRange src_bounds, dst_bounds
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.CopyRange_Results.Count
    
    Dim stored_value As Variant
    stored_value = test_dbl.CopyRange_Results.Items()(0)
    ' Array(SourceRangeBounds, AsValue, CopyNumberFormat)
    Assert.IsFalse stored_value(1) ' AsValue= False(既定)
    Assert.IsFalse stored_value(2) ' CopyNumberFormat= False(既定)
End Sub

Public Sub Test_IsEmptyCell_WithIgnoreEmptyFalse_ReturnsStub(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(10, 2, Sheet:="TestSheet")
    
    Call TUtl.SetValue(test_dbl.IsEmptyCell_Values, True, range_bounds, False)
    
    ' Act
    Dim actual_result As Boolean
    actual_result = test_dbl.IsEmptyCell(range_bounds, False)
    
    ' Assert
    Assert.IsTrue actual_result
End Sub

Public Sub Test_HasFormula_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim TUtl As New UnitTestUtils
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(Row:=2, Column:=3, Sheet:="Sheet1")
    Call TUtl.SetValue(test_dbl.HasFormula_Values, True, range_bounds)

    ' Act
    Dim actual_value As Boolean
    actual_value = sheet_srv.HasFormula(range_bounds)

    ' Assert
    Assert.IsTrue actual_value
End Sub
Public Sub Test_ClearRange_WithDefaults_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 3, Sheet:="TestSheet")
    
    ' Act
    sheet_srv.ClearRange range_bounds
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.ClearRange_Results.Count
    
    Dim stored_value As Variant
    stored_value = test_dbl.ClearRange_Results.Items()(0)
    ' Array(ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks)
    Assert.IsFalse stored_value(0) ' ClearAll=False
    Assert.IsTrue stored_value(1)  ' ClearContents=True
    Assert.IsFalse stored_value(2) ' ClearNumberFormats=False
End Sub

Public Sub Test_RemoveDuplicates_WithColumns_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(3, 1, 10, 4, Sheet:="TestSheet")
    Dim duplicate_columns As Variant
    duplicate_columns = Array(1, 3)
    
    ' Act
    Call sheet_srv.RemoveDuplicates(range_bounds, duplicate_columns)
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.RemoveDuplicates_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.RemoveDuplicates_Results.Items()(0)
    Assert.EqualsNumeric 1, stored_value(0)
    Assert.EqualsNumeric 3, stored_value(1)
End Sub

Public Sub Test_InsertRows_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 1, 8, 1, Sheet:="TestSheet")
    ' Act
    Call sheet_srv.InsertRows(range_bounds)
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.InsertRows_Results.Count
    Assert.IsTrue test_dbl.InsertRows_Results.Items()(0)
End Sub

Public Sub Test_DeleteRows_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 1, 8, 1, Sheet:="TestSheet")
    
    ' Act
    Call sheet_srv.DeleteRows(range_bounds)
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.DeleteRows_Results.Count
    Assert.IsTrue test_dbl.DeleteRows_Results.Items()(0)
End Sub

Public Sub Test_SetAllDataVisible_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")
    
    ' Act
    sheet_srv.SetAllDataVisible range_bounds
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetAllDataVisible_Results.Count
    Dim stored_val As Variant
    stored_val = test_dbl.SetAllDataVisible_Results.Items()(0)
    Assert.IsTrue stored_val ' True が格納される
End Sub

Public Sub Test_SetSheetOutlineLevel_WithRowAndColumn_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="TestSheet")
    
    ' Act
    sheet_srv.SetSheetOutlineLevel range_bounds, 2, 3
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetSheetOutlineLevel_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetSheetOutlineLevel_Results.Items()(0)
    ' Array(RowLevels, ColumnLevels)
    Assert.EqualsNumeric 2, stored_value(0)
    Assert.EqualsNumeric 3, stored_value(1)
End Sub

Public Sub Test_SetSheetTabColor_WithColorIndex_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="Sheet1")
    
    ' Act
    sheet_srv.SetSheetTabColor range_bounds, 5, -1
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetSheetTabColor_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetSheetTabColor_Results.Items()(0)
    ' Array(TabColorIndex, TabColor)
    Assert.EqualsNumeric 5, stored_value(0)
    Assert.EqualsNumeric -1, stored_value(1)
End Sub

Public Sub Test_SetRangeColor_WithColorIndex_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(2, 2, Sheet:="Sheet2")
    
    ' Act
    sheet_srv.SetRangeColor range_bounds, 3, -1, 6, -1
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetRangeColor_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetRangeColor_Results.Items()(0)
    ' Array(FontColorIndex, FontColor, InteriorColorIndex, InteriorColor)
    Assert.EqualsNumeric 3, stored_value(0)
    Assert.EqualsNumeric -1, stored_value(1)
    Assert.EqualsNumeric 6, stored_value(2)
End Sub

Public Sub Test_SetWrapText_True_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 10, Sheet:="WrapSheet")
    
    ' Act
    sheet_srv.SetWrapText range_bounds, True
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetWrapText_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetWrapText_Results.Items()(0)
    ' Array(WrapText)
    Assert.IsTrue stored_value(0)
End Sub

Public Sub Test_SetShrinkToFit_True_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(9, 9, Sheet:="ShrinkSheet")
    
    ' Act
    sheet_srv.SetShrinkToFit range_bounds, True
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetShrinkToFit_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetShrinkToFit_Results.Items()(0)
    ' Array(ShrinkToFit)
    Assert.IsTrue stored_value(0)
End Sub

Public Sub Test_SetAlignment_CustomOrientation_StoresParamsInDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(5, 5, Sheet:="AlignSheet")
    
    ' Act
    sheet_srv.SetAlignment range_bounds, xlCenter, xlTop, 45, 1
    
    ' Assert
    Assert.EqualsNumeric 1, test_dbl.SetAlignment_Results.Count
    Dim stored_value As Variant
    stored_value = test_dbl.SetAlignment_Results.Items()(0)
    ' Array(HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel)
    Assert.EqualsNumeric xlCenter, stored_value(0)
    Assert.EqualsNumeric xlTop, stored_value(1)
    Assert.EqualsNumeric 45, stored_value(2)
    Assert.EqualsNumeric 1, stored_value(3)
End Sub

Public Sub Test_EvaluateFormula_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    Dim range_bounds As WorksheetRangeBounds
    Set range_bounds = New_RangeBounds(1, 1, Sheet:="EvalSheet")
    
    Dim expected_value As Variant
    expected_value = 42
    Call TUtl.SetValue(test_dbl.EvaluateFormula_Values, expected_value, range_bounds, "=SUM(40,2)")
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.EvaluateFormula(range_bounds, "=SUM(40,2)")
    
    ' Assert
    Assert.EqualsNumeric 42, actual_value
End Sub

Public Sub Test_XLookup_WithStubValue_ReturnsExpectedValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim test_dbl As WorksheetServiceTestDouble: Set test_dbl = New WorksheetServiceTestDouble
    Dim sheet_srv As IWorksheetService: Set sheet_srv = test_dbl
    
    Dim lookup_bounds As WorksheetRangeBounds
    Set lookup_bounds = New_RangeBounds(2, 1, FinishRow:=5, FinishColumn:=1, Sheet:="LookupSheet")
    
    Dim return_bounds As WorksheetRangeBounds
    Set return_bounds = New_RangeBounds(2, 2, FinishRow:=5, FinishColumn:=2, Sheet:="LookupSheet")
    
    Dim expected_value As Variant
    expected_value = "Beta"
    Call TUtl.SetValue(test_dbl.XLookup_Values, expected_value, "B", lookup_bounds, return_bounds, "missing", 0&, 1&)
    
    ' Act
    Dim actual_value As Variant
    actual_value = sheet_srv.XLookup("B", lookup_bounds, return_bounds, "missing")
    
    ' Assert
    Assert.Equals "Beta", CStr(actual_value)
End Sub
