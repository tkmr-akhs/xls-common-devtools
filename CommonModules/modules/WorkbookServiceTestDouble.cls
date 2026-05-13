VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! Excel ファイルのオープン、クローズ、書き込みなどを管理するサービスのテストダブルです。
'!
' #############################################################################

Implements IWorkbookService

'* GetThisWorkbookName の戻り値。
Public GetThisWorkbookName_Value As String

'* GetThisWorkbookDirectoryPath の戻り値。
Public GetThisWorkbookDirectoryPath_Value As String

'* ExistsWorkbook の引数→戻り値辞書。引数がキーとなり、戻り値が値となります。この辞書に含まれていない場合は False となります。
Public ExistsWorkbook_Values As Dictionary

'* GetAllWorkbook の引数→受け取り値辞書。Nothing がキーとなり、戻り値が値となります。
Public GetAllWorkbook_Values As Dictionary

'* ExistsWorksheet の引数→戻り値辞書。引数がキーとなり、戻り値が値となります。この辞書に含まれていない場合は False となります。
Public ExistsWorksheet_Values As Dictionary

'* GetAllWorksheet の引数→戻り値辞書。
Public GetAllWorksheet_Values As Dictionary

'* OpenWorkbook の引数→受け取り値辞書。引数がキーとなり、ブック名が値となります。
Public OpenWorkbook_Results As Dictionary

'* SaveWorkbook の引数→受け取り値辞書。引数がキーとなり、FilePath のファイル名部分が値となります。
Public SaveWorkbook_Results As Dictionary

'* IsSaved の引数→戻り値辞書。引数がキーとなり、戻り値が値となります。
Public IsSaved_Values As Dictionary

'* SaveWorkbook の引数→受け取り値辞書。引数がキーとなり、値が True となります。
Public ColoseWorkbook_Results As Dictionary

'* Find の引数→戻り値辞書。
Public Find_Values As Dictionary

'* AddWorksheet の引数→受け取り値辞書。引数がキーとなり、値がシート名となります。
Public AddWorksheet_Results As Dictionary

'* RemoveWorksheet の引数→受け取り値辞書。引数がキーとなり、値が True となります。
Public RemoveWorksheet_Results As Dictionary


'* CopyWorksheet の引数→受け取り値辞書。引数がキーとなり、値が複製先シート名となります。
Public CopyWorksheet_Results As Dictionary

'* ActivateWorksheet の引数→受け取り値辞書。引数がキーとなり、値が True となります。
Public ActivateWorksheet_Results As Dictionary

'* RemoveVBComponents の引数→受け取り値辞書。引数がキーとなり、値が True となります。
Public RemoveVBComponents_Results As Dictionary

'* 生成されたブック数
Public BookCount As Long

'* 生成されたシート数
Public AddSheetCount As Long

'* 複製されたシート数
Public CopySheetCount As Long

Private TUtl As UnitTestUtils

Private Sub Class_Initialize()
    Set TUtl = New UnitTestUtils

    GetThisWorkbookName_Value = "ThisWorkbook.xlsm"
    GetThisWorkbookDirectoryPath_Value = ""
    Set ExistsWorkbook_Values = New Dictionary
    Set GetAllWorkbook_Values = New Dictionary
    Set ExistsWorksheet_Values = New Dictionary
    Set GetAllWorksheet_Values = New Dictionary
    Set OpenWorkbook_Results = New Dictionary
    Set SaveWorkbook_Results = New Dictionary
    Set IsSaved_Values = New Dictionary
    Set ColoseWorkbook_Results = New Dictionary
    Set Find_Values = New Dictionary
    Set AddWorksheet_Results = New Dictionary
    Set RemoveWorksheet_Results = New Dictionary
    Set CopyWorksheet_Results = New Dictionary
    Set ActivateWorksheet_Results = New Dictionary
    Set RemoveVBComponents_Results = New Dictionary
    
    BookCount = 1
    AddSheetCount = 1
    CopySheetCount = 1
End Sub

'* GetThisWorkbookName のスタブです。
'*
'* @details
'* GetThisWorkbookName_Value を返します。
Public Function GetThisWorkbookName() As String
    GetThisWorkbookName = GetThisWorkbookName_Value
End Function
Private Function IWorkbookService_GetThisWorkbookName() As String
    IWorkbookService_GetThisWorkbookName = GetThisWorkbookName()
End Function
'* GetThisWorkbookDirectoryPath のスタブです。
'*
'* @details
'* GetThisWorkbookDirectoryPath_Value を返します。
Public Function GetThisWorkbookDirectoryPath() As String
    GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath_Value
End Function
Private Function IWorkbookService_GetThisWorkbookDirectoryPath() As String
    IWorkbookService_GetThisWorkbookDirectoryPath = GetThisWorkbookDirectoryPath()
End Function

'* ExistsWorkbook のスタブです。
'*
'* @details
'* ExistsWorkbook = TUtl.GetValue(ExistsWorkbook, Book)
Public Function ExistsWorkbook(ByVal Book As String) As Boolean
    If TUtl.HasValue(ExistsWorkbook_Values, Book) Then
        ExistsWorkbook = TUtl.GetValue(ExistsWorkbook_Values, Book)
    Else
        ExistsWorkbook = False
    End If
End Function
Private Function IWorkbookService_ExistsWorkbook(ByVal Book As String) As Boolean
    IWorkbookService_ExistsWorkbook = ExistsWorkbook(Book)
End Function

'* GetAllWorkbook のスタブです。
'*
'* @details
'* GetAllWorkbook = TUtl.GetValue(GetAllWorkbook_Values, Nothing)
Public Function GetAllWorkbook() As String()
    GetAllWorkbook = TUtl.GetValue(GetAllWorkbook_Values, Nothing)
End Function
Private Function IWorkbookService_GetAllWorkbook() As String()
    IWorkbookService_GetAllWorkbook = GetAllWorkbook()
End Function

'* ExistsWorksheet のスタブです。
'*
'* @details
'*  ExistsWorksheet = TUtl.GetValue(ExistsWorksheet_Values, Sheet, Book)
Public Function ExistsWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    If TUtl.HasValue(ExistsWorksheet_Values, Sheet, Book) Then
        ExistsWorksheet = TUtl.GetValue(ExistsWorksheet_Values, Sheet, Book)
    Else
        ExistsWorksheet = False
    End If
End Function
Private Function IWorkbookService_ExistsWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "") As Boolean
    IWorkbookService_ExistsWorksheet = ExistsWorksheet(Sheet, Book:=Book)
End Function

'* GetAllWorksheet のスタブです。
'*
'* @details
'* GetAllWorksheet = TUtl.GetValue(GetAllWorksheet_Values, Book)
Public Function GetAllWorksheet(Optional ByVal Book As String = "") As String()
    GetAllWorksheet = TUtl.GetValue(GetAllWorksheet_Values, Book)
End Function
Private Function IWorkbookService_GetAllWorksheet(Optional ByVal Book As String = "") As String()
    IWorkbookService_GetAllWorksheet = GetAllWorksheet(Book:=Book)
End Function

'* OpenWorkbook のスパイ兼スタブです。戻り値は、"Book" & BookCount となります。
'*
'* @details
'* Call TUtl.SetValue(OpenWorkbook_Results, "Book" & BookCount, FilePath)
Public Function OpenWorkbook(Optional ByVal FilePath As String = "") As String
    Call TUtl.SetValue(OpenWorkbook_Results, "Book" & BookCount, FilePath)
    OpenWorkbook = "Book" & BookCount
    BookCount = BookCount + 1
End Function
Private Function IWorkbookService_OpenWorkbook(Optional ByVal FilePath As String = "") As String
    IWorkbookService_OpenWorkbook = OpenWorkbook(FilePath:=FilePath)
End Function

'* SaveWorkbook のスパイ兼スタブです。戻り値は FilePath のファイル名部分となります。
'*
'* @details
'* Call TUtl.SetValue(SaveWorkbook_Results, GetLeafFromPath(FilePath), Book, FilePath, Force)
Public Function SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    Call TUtl.SetValue(SaveWorkbook_Results, GetLeafFromPath(FilePath), Book, FilePath, Force)
    SaveWorkbook = GetLeafFromPath(FilePath)
End Function
Private Function IWorkbookService_SaveWorkbook(ByVal Book As String, ByVal FilePath As String, Optional ByVal Force As Boolean = True) As String
    IWorkbookService_SaveWorkbook = SaveWorkbook(Book, FilePath, Force:=Force)
End Function

'* IsSaved のスタブです。
'*
'* @details
'* IsSaved = TUtl.GetValue(IsSaved_Values, FilePath)
Public Function IsSaved(ByVal Book As String) As Boolean
    If TUtl.HasValue(IsSaved_Values, Book) Then
        IsSaved = TUtl.GetValue(IsSaved_Values, Book)
    Else
        IsSaved = False
    End If
End Function
Private Function IWorkbookService_IsSaved(ByVal Book As String) As Boolean
    IWorkbookService_IsSaved = IsSaved(Book)
End Function

'* CloseWorkbook のスパイです。
'*
'* @details
'*
Public Sub CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call TUtl.SetValue(ColoseWorkbook_Results, True, Book, Force)
End Sub
Private Sub IWorkbookService_CloseWorkbook(ByVal Book As String, Optional ByVal Force As Boolean = True)
    Call CloseWorkbook(Book, Force:=Force)
End Sub

'* Find のスタブです。
'*
'* @details
'* Find = TUtl.GetValue(Find_Values, What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)
Public Function Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    Find = TUtl.GetValue(Find_Values, What, Book, LookInValues, LookAtWhole, MatchCase, MatchByte)
End Function
Private Function IWorkbookService_Find( _
        ByVal What As String, _
        Optional ByVal Book As String = "", _
        Optional ByVal LookInValues As Boolean = True, _
        Optional ByVal LookAtWhole As Boolean = True, _
        Optional ByVal MatchCase As Boolean = True, _
        Optional ByVal MatchByte As Boolean = True) As WorksheetRangeBounds()
    
    IWorkbookService_Find = Find( _
            What, Book:=Book, _
            LookInValues:=LookInValues, _
            LookAtWhole:=LookAtWhole, _
            MatchCase:=MatchCase, _
            MatchByte:=MatchByte)
End Function

'* AddWorksheet のスパイ兼スタブです。戻り値は Sheet か、省略されたときは "Sheet" & AddSheetCount となります。
'*
'* @details
'* Call TUtl.SetValue(AddWorksheet_Results, result_value, Sheet, Book, SheetIndex, Before)
Public Function AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    Dim result_value As String
    If Sheet = "" Then
        result_value = "Sheet" & AddSheetCount
        AddSheetCount = AddSheetCount + 1
    Else
        result_value = Sheet
    End If
    
    Call TUtl.SetValue(AddWorksheet_Results, result_value, Sheet, Book, SheetIndex, Before)
    
    AddWorksheet = result_value
End Function
Private Function IWorkbookService_AddWorksheet( _
        Optional ByVal Sheet As String = "", _
        Optional ByVal Book As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    IWorkbookService_AddWorksheet = AddWorksheet( _
            Sheet:=Sheet, Book:=Book, _
            SheetIndex:=SheetIndex, Before:=Before)
End Function

'* RemoveWorksheet のスパイです。
'*
'* @details
'* Call TUtl.SetValue(RemoveWorksheet_Results, True, Sheet, Book)
Public Sub RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call TUtl.SetValue(RemoveWorksheet_Results, True, Sheet, Book)
End Sub
Private Sub IWorkbookService_RemoveWorksheet(Optional ByVal Sheet As String = "", Optional ByVal Book As String = "")
    Call RemoveWorksheet(Sheet:=Sheet, Book:=Book)
End Sub

'* CopyWorksheet のスパイ兼スタブです。戻り値は DestinationWorksheetName か、省略されたときは "Sheet" & CopySheetCount となります。
'*
'* @details
'* Call TUtl.SetValue(CopyWorksheet_Results, result_value, SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)
Public Function CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    Dim result_value As String
    If DestinationWorksheetName = "" Then
        result_value = "Sheet" & CopySheetCount
        CopySheetCount = CopySheetCount + 1
    Else
        result_value = DestinationWorksheetName
    End If
    
    Call TUtl.SetValue(CopyWorksheet_Results, result_value, SourceWorksheetName, SourceWorkbookName, DestinationWorksheetName, DestinationWorkbookName, SheetIndex, Before)
    
    CopyWorksheet = result_value
End Function
Private Function IWorkbookService_CopyWorksheet( _
        Optional ByVal SourceWorksheetName As String = "", _
        Optional ByVal SourceWorkbookName As String = "", _
        Optional ByVal DestinationWorksheetName As String = "", _
        Optional ByVal DestinationWorkbookName As String = "", _
        Optional ByVal SheetIndex As Long = 0, _
        Optional ByVal Before As Boolean = False) As String
    
    IWorkbookService_CopyWorksheet = CopyWorksheet( _
            SourceWorksheetName:=SourceWorksheetName, _
            SourceWorkbookName:=SourceWorkbookName, _
            DestinationWorksheetName:=DestinationWorksheetName, _
            DestinationWorkbookName:=DestinationWorkbookName, _
            SheetIndex:=SheetIndex, _
            Before:=Before)
End Function

'* RemoveWorksheet のスパイです。
'*
'* @details
'* Call TUtl.SetValue(RemoveWorksheet_Results, True, Sheet, Book)
Public Sub ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call TUtl.SetValue(ActivateWorksheet_Results, True, Sheet, Book)
End Sub
Private Sub IWorkbookService_ActivateWorksheet(ByVal Sheet As String, Optional ByVal Book As String = "")
    Call ActivateWorksheet(Sheet, Book:=Book)
End Sub

'* RemoveVBComponents のスパイです。
'*
'* @details
'* Call TUtl.SetValue(RemoveVBComponents_Results, True, names, ComponentType, Book)
Public Sub RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call TUtl.SetValue(RemoveVBComponents_Results, True, pBuildComponentNamesKey(ComponentNames), ComponentType, Book)
End Sub
Private Sub IWorkbookService_RemoveVBComponents(ByVal ComponentNames As Variant, Optional ByVal ComponentType As Long = 0, Optional ByVal Book As String = "")
    Call RemoveVBComponents(ComponentNames, ComponentType:=ComponentType, Book:=Book)
End Sub

Private Function pBuildComponentNamesKey(ByVal ComponentNames As Variant) As String
    If IsArray(ComponentNames) Then
        Dim result_value As String
        Dim name_item As Variant
        For Each name_item In ComponentNames
            If result_value <> "" Then result_value = result_value & ","
            result_value = result_value & CStr(name_item)
        Next name_item
        pBuildComponentNamesKey = result_value
    Else
        pBuildComponentNamesKey = CStr(ComponentNames)
    End If
End Function
