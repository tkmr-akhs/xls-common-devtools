VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWorksheetService

' #############################################################################
'!
'! @brief
'! ワークシートへのアクセスを管理するサービスのテストダブルです。
'!
' #############################################################################

Private TUtl As UnitTestUtils

'* Find の引数→戻り値辞書。
Public Find_Values As Dictionary

'* Sort の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public Sort_Results As Dictionary

'* ActivateRange の引数→受け取り値辞書。RangeBounds がキーとなり、値が True となります。
Public ActivateRange_Results As Dictionary

'* WriteCell の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public WriteCell_Results As Dictionary

'* WriteRange の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public WriteRange_Results As Dictionary

'* WriteArrayFormula の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public WriteArrayFormula_Results As Dictionary

'* ReadCell の引数→戻り値辞書。戻り値は [ Expression, NumberFormat ] の 1 次元配列です。
Public ReadCell_Values As Dictionary

'* ReadRange の引数→戻り値辞書。
Public ReadRange_Values As Dictionary

'* EvaluateFormula の引数→戻り値辞書。ContextRangeBounds と Formula がキーとなります。
Public EvaluateFormula_Values As Dictionary

'* XLookup の引数→戻り値辞書。
Public XLookup_Values As Dictionary

'* CopyCell の引数→受け取り値辞書。DestinationRangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public CopyCell_Results As Dictionary

'* CopyRange の引数→受け取り値辞書。DestinationRangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public CopyRange_Results As Dictionary

'* IsEmptyCell の引数→戻り値辞書。
Public IsEmptyCell_Values As Dictionary

'* HasFormula の引数→戻り値辞書。
Public HasFormula_Values As Dictionary

'* GetUsedRangeBounds の引数→戻り値辞書。
Public GetUsedRangeBounds_Values As Dictionary

'* ClearRange の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public ClearRange_Results As Dictionary

'* RemoveDuplicates の引数→受け取り値辞書。RangeBounds がキーとなり、値には重複判定列配列が入ります。
Public RemoveDuplicates_Results As Dictionary

'* InsertRows の引数→受け取り値辞書。RangeBounds がキーとなり、値には True が入ります。
Public InsertRows_Results As Dictionary

'* DeleteRows の引数→受け取り値辞書。RangeBounds がキーとなり、値には True が入ります。
Public DeleteRows_Results As Dictionary

'* SetAllDataVisible の引数→受け取り値辞書。RangeBounds がキーとなり、値には True が入ります。
Public SetAllDataVisible_Results As Dictionary

'* SetSheetOutlineLevel の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetSheetOutlineLevel_Results As Dictionary

'* SetSheetTabColor の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetSheetTabColor_Results As Dictionary

'* SetRangeColor の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetRangeColor_Results As Dictionary

'* SetWrapText の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetWrapText_Results As Dictionary

'* SetShrinkToFit の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetShrinkToFit_Results As Dictionary

'* SetAlignment の引数→受け取り値辞書。RangeBounds がキーとなり、それ以外の 1 次元配列が値となります。
Public SetAlignment_Results As Dictionary

Private Sub Class_Initialize()
    Set TUtl = New UnitTestUtils
    
    Set Find_Values = New Dictionary
    Set Sort_Results = New Dictionary
    Set ActivateRange_Results = New Dictionary
    Set WriteCell_Results = New Dictionary
    Set WriteRange_Results = New Dictionary
    Set WriteArrayFormula_Results = New Dictionary
    Set ReadCell_Values = New Dictionary
    Set ReadRange_Values = New Dictionary
    Set EvaluateFormula_Values = New Dictionary
    Set XLookup_Values = New Dictionary
    Set CopyCell_Results = New Dictionary
    Set CopyRange_Results = New Dictionary
    Set IsEmptyCell_Values = New Dictionary
    Set HasFormula_Values = New Dictionary
    Set GetUsedRangeBounds_Values = New Dictionary
    Set ClearRange_Results = New Dictionary
    Set RemoveDuplicates_Results = New Dictionary
    Set InsertRows_Results = New Dictionary
    Set DeleteRows_Results = New Dictionary
    Set SetAllDataVisible_Results = New Dictionary
    Set SetSheetOutlineLevel_Results = New Dictionary
    Set SetSheetTabColor_Results = New Dictionary
    Set SetRangeColor_Results = New Dictionary
    Set SetWrapText_Results = New Dictionary
    Set SetShrinkToFit_Results = New Dictionary
    Set SetAlignment_Results = New Dictionary
End Sub

'* Find のスタブです。
'*
'* @details
'* Find = TUtl.GetValue(Find_Values, What, RangeBounds, LookInValues, LookAtWhole, MatchCase, MatchByte)
Public Function Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    Find = TUtl.GetValue(Find_Values, What, RangeBounds, LookInValues, LookAtWhole, MatchCase, MatchByte)
End Function
Private Function IWorksheetService_Find( _
        ByVal What As String, _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    IWorksheetService_Find = Find( _
            What, RangeBounds, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

'* Sort のスパイです。
'*
'* @details
'*
Public Sub Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub
Private Sub IWorksheetService_Sort(ByVal RangeBounds As WorksheetRangeBounds, ParamArray SortKeyAndOrder() As Variant)
    Call pSortCore(RangeBounds, SortKeyAndOrder)
End Sub

Private Sub pSortCore(ByVal RangeBounds As WorksheetRangeBounds, ByVal SortKeyAndOrder As Variant)
    Call TUtl.SetValue(Sort_Results, SortKeyAndOrder, RangeBounds)
End Sub

'* ActivateRange のスパイです。
'*
'* @details
'* Call TUtl.SetValue(WriteCell_Results, Array(Expression, NumberFormat, AsFormula, ClearWhenEmpty, IgnoreEmpty, TypeConvert), RangeBounds)
Public Sub ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call TUtl.SetValue(ActivateRange_Results, True, RangeBounds)
End Sub
Private Sub IWorksheetService_ActivateRange(ByVal RangeBounds As WorksheetRangeBounds)
    Call ActivateRange(RangeBounds)
End Sub

'* WriteCell のスパイです。
'*
'* @details
'* Call TUtl.SetValue(WriteCell_Results, Array(Expression, NumberFormat, AsFormula, ClearWhenEmpty, IgnoreEmpty, TypeConvert), RangeBounds)
Public Sub WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
    
    Call TUtl.SetValue(WriteCell_Results, Array(Expression, NumberFormat, AsFormula, ClearWhenEmpty, IgnoreEmpty, TypeConvert), RangeBounds)
End Sub
Private Sub IWorksheetService_WriteCell( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        ByVal Expression As String, _
        Optional ByVal NumberFormat As String = "", _
        Optional ByVal AsFormula As Boolean = False, _
        Optional ByVal ClearWhenEmpty As Boolean = True, _
        Optional ByVal IgnoreEmpty As Boolean = False, _
        Optional ByVal TypeConvert As Boolean = True)
    
    Call WriteCell( _
            RangeBounds, _
            Expression, _
            NumberFormat:=NumberFormat, _
            AsFormula:=AsFormula, _
            ClearWhenEmpty:=ClearWhenEmpty, _
            IgnoreEmpty:=IgnoreEmpty, _
            TypeConvert:=TypeConvert)
End Sub

'* WriteRange のスパイです。
'*
'* @details
'* Call TUtl.SetValue(WriteRange_Results, Array(ValuesArray), RangeBounds)
Public Sub WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call TUtl.SetValue(WriteRange_Results, Array(ValuesArray), RangeBounds)
End Sub
Private Sub IWorksheetService_WriteRange(ByVal RangeBounds As WorksheetRangeBounds, ByVal ValuesArray As Variant)
    Call WriteRange(RangeBounds, ValuesArray)
End Sub

'* WriteArrayFormula のスパイです。
'*
'* @details
'* Call TUtl.SetValue(WriteArrayFormula_Results, Array(FormulaArray), RangeBounds)
Public Sub WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call TUtl.SetValue(WriteArrayFormula_Results, Array(FormulaArray), RangeBounds)
End Sub
Private Sub IWorksheetService_WriteArrayFormula(ByVal RangeBounds As WorksheetRangeBounds, ByVal FormulaArray As String)
    Call WriteArrayFormula(RangeBounds, FormulaArray)
End Sub

'* ReadCell のスタブです。
'*
'* @details
'* Dim value_arr As Variant
'* value_arr = TUtl.GetValue(ReadCell_Values, RangeBounds, GetText)
'* Expression = value_arr(LBound(value_arr))
'* NumberFormat = value_arr(LBound(value_arr) + 1)
Public Function ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    Dim value_arr As Variant
    value_arr = TUtl.GetValue(ReadCell_Values, RangeBounds, GetText)
    Expression = value_arr(LBound(value_arr))
    NumberFormat = value_arr(LBound(value_arr) + 1)
    ReadCell = Expression
End Function
Private Function IWorksheetService_ReadCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByRef Expression As String = "", Optional ByRef NumberFormat As String = "", Optional ByVal GetText As Boolean = False) As String
    IWorksheetService_ReadCell = ReadCell(RangeBounds, Expression, NumberFormat:=NumberFormat, GetText:=GetText)
End Function

'* ReadRange のスタブです。
'*
'* @details
'* ReadRange = TUtl.GetValue(ReadRange_Values, RangeBounds)
Public Function ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    ReadRange = TUtl.GetValue(ReadRange_Values, RangeBounds)
End Function
Private Function IWorksheetService_ReadRange(ByVal RangeBounds As WorksheetRangeBounds) As Variant()
    IWorksheetService_ReadRange = ReadRange(RangeBounds)
End Function

'* EvaluateFormula のスタブです。
'*
'* @details
'* EvaluateFormula = TUtl.GetValue(EvaluateFormula_Values, ContextRangeBounds, Formula)
Public Function EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    EvaluateFormula = TUtl.GetValue(EvaluateFormula_Values, ContextRangeBounds, Formula)
End Function
Private Function IWorksheetService_EvaluateFormula(ByVal ContextRangeBounds As WorksheetRangeBounds, ByVal Formula As String) As Variant
    IWorksheetService_EvaluateFormula = EvaluateFormula(ContextRangeBounds, Formula)
End Function

'* XLookup のスタブです。
'*
'* @details
'* XLookup = TUtl.GetValue(XLookup_Values, LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
Public Function XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
    
    XLookup = TUtl.GetValue(XLookup_Values, LookupValue, LookupRangeBounds, ReturnRangeBounds, IfNotFound, MatchMode, SearchMode)
End Function
Private Function IWorksheetService_XLookup( _
        ByVal LookupValue As Variant, _
        ByVal LookupRangeBounds As WorksheetRangeBounds, _
        ByVal ReturnRangeBounds As WorksheetRangeBounds, _
        Optional ByVal IfNotFound As Variant = "", _
        Optional ByVal MatchMode As Long = 0, _
        Optional ByVal SearchMode As Long = 1) As Variant
    
    IWorksheetService_XLookup = XLookup( _
            LookupValue, _
            LookupRangeBounds, _
            ReturnRangeBounds, _
            IfNotFound:=IfNotFound, _
            MatchMode:=MatchMode, _
            SearchMode:=SearchMode)
End Function

'* CopyCell のスパイです。
'*
'* @details
'* Call TUtl.SetValue(CopyCell_Results, Array(SourceRangeBounds, AsValue, CopyNumberFormat), DestinationRangeBounds)
Public Sub CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    Call TUtl.SetValue(CopyCell_Results, Array(SourceRangeBounds, AsValue, CopyNumberFormat), DestinationRangeBounds)
End Sub
Private Sub IWorksheetService_CopyCell( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    Call CopyCell( _
            SourceRangeBounds, _
            DestinationRangeBounds, _
            AsValue:=AsValue, _
            CopyNumberFormat:=CopyNumberFormat)
End Sub

'* CopyRange のスパイです。
'*
'* @details
'* Call TUtl.SetValue(CopyRange_Results, Array(SourceRangeBounds, AsValue, CopyNumberFormat), DestinationRangeBounds)
Public Sub CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
        
    Call TUtl.SetValue(CopyRange_Results, Array(SourceRangeBounds, AsValue, CopyNumberFormat), DestinationRangeBounds)
End Sub
Private Sub IWorksheetService_CopyRange( _
        ByVal SourceRangeBounds As WorksheetRangeBounds, _
        ByVal DestinationRangeBounds As WorksheetRangeBounds, _
        Optional ByVal AsValue As Boolean = False, _
        Optional ByVal CopyNumberFormat As Boolean = False)
    
    Call CopyRange( _
            SourceRangeBounds, _
            DestinationRangeBounds, _
            AsValue:=AsValue, _
            CopyNumberFormat:=CopyNumberFormat)
End Sub

'* IsEmptyCell のスタブです。
'*
'* @details
'* IsEmptyCell = TUtl.GetValue(IsEmptyCell_Values, RangeBounds, IgnoreEmptyString)
Public Function IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    IsEmptyCell = TUtl.GetValue(IsEmptyCell_Values, RangeBounds, IgnoreEmptyString)
End Function
Private Function IWorksheetService_IsEmptyCell(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal IgnoreEmptyString As Boolean = False) As Boolean
    IWorksheetService_IsEmptyCell = IsEmptyCell(RangeBounds, IgnoreEmptyString:=IgnoreEmptyString)
End Function

'* HasFormula のスタブです。
'*
'* @details
'* HasFormula = TUtl.GetValue(HasFormula_Values, RangeBounds)
Public Function HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    HasFormula = TUtl.GetValue(HasFormula_Values, RangeBounds)
End Function
Private Function IWorksheetService_HasFormula(ByVal RangeBounds As WorksheetRangeBounds) As Boolean
    IWorksheetService_HasFormula = HasFormula(RangeBounds)
End Function
'* GetUsedRangeBounds のスタブです。
'*
'* @details
'* Set GetUsedRangeBounds = TUtl.GetValue(GetUsedRangeBounds_Values, RangeBounds, IgnoreEmpty, IsUsedStartRow, IsUsedStartColumn, GetRawRange)
Public Function GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
    
    Set GetUsedRangeBounds = TUtl.GetValue(GetUsedRangeBounds_Values, RangeBounds, IgnoreEmpty, IsUsedStartRow, IsUsedStartColumn, GetRawRange)
End Function
Private Function IWorksheetService_GetUsedRangeBounds( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal IgnoreEmpty As Boolean = True, _
        Optional ByVal IsUsedStartRow As Boolean = True, _
        Optional ByVal IsUsedStartColumn As Boolean = True, _
        Optional ByVal GetRawRange As Boolean = False) As WorksheetRangeBounds
    
    Set IWorksheetService_GetUsedRangeBounds = GetUsedRangeBounds( _
            RangeBounds, _
            IgnoreEmpty:=IgnoreEmpty, _
            IsUsedStartRow:=IsUsedStartRow, _
            IsUsedStartColumn:=IsUsedStartColumn, _
            GetRawRange:=GetRawRange)
End Function

'* ClearRange のスパイです。
'*
'* @details
'* Call TUtl.SetValue(ClearRange_Results, Array(ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks), RangeBounds)
Public Sub ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
    
    Call TUtl.SetValue(ClearRange_Results, Array(ClearAll, ClearContents, ClearNumberFormats, ClearColors, ClearComments, ClearHyperlinks), RangeBounds)
End Sub
Private Sub IWorksheetService_ClearRange( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal ClearAll As Boolean = False, _
        Optional ByVal ClearContents As Boolean = True, _
        Optional ByVal ClearNumberFormats As Boolean = False, _
        Optional ByVal ClearColors As Boolean = False, _
        Optional ByVal ClearComments As Boolean = False, _
        Optional ByVal ClearHyperlinks As Boolean = False)
    
    Call ClearRange( _
            RangeBounds, _
            ClearAll:=ClearAll, _
            ClearContents:=ClearContents, _
            ClearNumberFormats:=ClearNumberFormats, _
            ClearColors:=ClearColors, _
            ClearComments:=ClearComments, _
            ClearHyperlinks:=ClearHyperlinks)
End Sub

'* RemoveDuplicates のスパイです。
'*
'* @details
'* Call TUtl.SetValue(RemoveDuplicates_Results, DuplicateColumns, RangeBounds)
Public Sub RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call TUtl.SetValue(RemoveDuplicates_Results, DuplicateColumns, RangeBounds)
End Sub
Private Sub IWorksheetService_RemoveDuplicates(ByVal RangeBounds As WorksheetRangeBounds, ByVal DuplicateColumns As Variant)
    Call RemoveDuplicates(RangeBounds, DuplicateColumns)
End Sub

'* InsertRows のスパイです。
'*
'* @details
'* Call TUtl.SetValue(InsertRows_Results, True, RangeBounds)
Public Sub InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call TUtl.SetValue(InsertRows_Results, True, RangeBounds)
End Sub
Private Sub IWorksheetService_InsertRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call InsertRows(RangeBounds)
End Sub

'* DeleteRows のスパイです。
'*
'* @details
'* Call TUtl.SetValue(DeleteRows_Results, True, RangeBounds)
Public Sub DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call TUtl.SetValue(DeleteRows_Results, True, RangeBounds)
End Sub
Private Sub IWorksheetService_DeleteRows(ByVal RangeBounds As WorksheetRangeBounds)
    Call DeleteRows(RangeBounds)
End Sub
'* SetAllDataVisible のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetAllDataVisible_Results, True, RangeBounds)
Public Sub SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call TUtl.SetValue(SetAllDataVisible_Results, True, RangeBounds)
End Sub
Private Sub IWorksheetService_SetAllDataVisible(ByVal RangeBounds As WorksheetRangeBounds)
    Call SetAllDataVisible(RangeBounds)
End Sub

'* SetSheetOutlineLevel のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetSheetOutlineLevel_Results, Array(RowLevels, ColumnLevels), RangeBounds)
Public Sub SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call TUtl.SetValue(SetSheetOutlineLevel_Results, Array(RowLevels, ColumnLevels), RangeBounds)
End Sub
Private Sub IWorksheetService_SetSheetOutlineLevel(ByVal RangeBounds As WorksheetRangeBounds, Optional ByVal RowLevels As Long = 0, Optional ByVal ColumnLevels As Long = 0)
    Call SetSheetOutlineLevel(RangeBounds, RowLevels:=RowLevels, ColumnLevels:=ColumnLevels)
End Sub

'* SetSheetTabColor のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetSheetTabColor_Results, Array(TabColorIndex, TabColor), RangeBounds)
Public Sub SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
    
    Call TUtl.SetValue(SetSheetTabColor_Results, Array(TabColorIndex, TabColor), RangeBounds)
End Sub
Private Sub IWorksheetService_SetSheetTabColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal TabColorIndex As Long = xlColorIndexNone, _
        Optional ByVal TabColor As Long = -1)
    
    Call SetSheetTabColor( _
            RangeBounds, _
            TabColorIndex:=TabColorIndex, _
            TabColor:=TabColor)
End Sub

'* SetRangeColor のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetRangeColor_Results, Array(FontColorIndex, FontColor, InteriorColorIndex, InteriorColor), RangeBounds)
Public Sub SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
    
    Call TUtl.SetValue(SetRangeColor_Results, Array(FontColorIndex, FontColor, InteriorColorIndex, InteriorColor), RangeBounds)
End Sub
Private Sub IWorksheetService_SetRangeColor( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal FontColorIndex As Long = 0, _
        Optional ByVal FontColor As Long = -1, _
        Optional ByVal InteriorColorIndex As Long = xlColorIndexNone, _
        Optional ByVal InteriorColor As Long = -1)
    
    Call SetRangeColor( _
        RangeBounds, _
        FontColorIndex:=FontColorIndex, _
        FontColor:=FontColor, _
        InteriorColorIndex:=InteriorColorIndex, _
        InteriorColor:=InteriorColor)
End Sub

'* SetWrapText のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetWrapText_Results, Array(WrapText), RangeBounds)
Public Sub SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call TUtl.SetValue(SetWrapText_Results, Array(WrapText), RangeBounds)
End Sub
Private Sub IWorksheetService_SetWrapText(ByVal RangeBounds As WorksheetRangeBounds, ByVal WrapText As Boolean)
    Call SetWrapText(RangeBounds, WrapText)
End Sub

'* SetShrinkToFit のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetShrinkToFit_Results, Array(ShrinkToFit), RangeBounds)
Public Sub SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call TUtl.SetValue(SetShrinkToFit_Results, Array(ShrinkToFit), RangeBounds)
End Sub
Private Sub IWorksheetService_SetShrinkToFit(ByVal RangeBounds As WorksheetRangeBounds, ByVal ShrinkToFit As Boolean)
    Call SetShrinkToFit(RangeBounds, ShrinkToFit)
End Sub

'* SetAlignment のスパイです。
'*
'* @details
'* Call TUtl.SetValue(SetAlignment_Results, Array(HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel), RangeBounds)
Public Sub SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = xlGeneral, _
        Optional ByVal VerticalAlignment As Long = xlBottom, _
        Optional ByVal Orientation As Long = 0, _
        Optional ByVal IndentLevel As Long = 0)
    
    Call TUtl.SetValue(SetAlignment_Results, Array(HorizontalAlignment, VerticalAlignment, Orientation, IndentLevel), RangeBounds)
End Sub
Private Sub IWorksheetService_SetAlignment( _
        ByVal RangeBounds As WorksheetRangeBounds, _
        Optional ByVal HorizontalAlignment As Long = xlGeneral, _
        Optional ByVal VerticalAlignment As Long = xlBottom, _
        Optional ByVal Orientation As Long = 0, _
        Optional ByVal IndentLevel As Long = 0)
    
    Call SetAlignment( _
            RangeBounds, _
            HorizontalAlignment:=HorizontalAlignment, _
            VerticalAlignment:=VerticalAlignment, _
            Orientation:=Orientation, _
            IndentLevel:=IndentLevel)
End Sub
