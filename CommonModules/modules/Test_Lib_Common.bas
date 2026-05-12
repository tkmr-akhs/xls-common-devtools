Attribute VB_Name = "Test_Lib_Common"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! Lib_Common モジュールのユニット テストです。
'!
' #############################################################################

Public Sub Test_InitializeCommonService_NotInitialized_InitializesWorkbookAndWorksheetServices(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Set WbSrv = Nothing
    Set WsSrv = Nothing

    ' Act
    Call InitializeCommonService

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.IsTypeOf "WorkbookService", WbSrv
    Assert.IsTypeOf "WorksheetService", WsSrv
End Sub
' -----------------------------------------------------------------------------
' GetExcelFileFormat
' -----------------------------------------------------------------------------

Public Sub Test_GetExcelFileFormat_UppercaseXlsm_ReturnsMacroEnabledWorkbook(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.XLSM")

    ' Assert
    Assert.EqualsNumeric xlOpenXMLWorkbookMacroEnabled, actual_value
End Sub

Public Sub Test_GetExcelFileFormat_MixedCaseCsv_ReturnsCsv(ByVal Assert As UnitTestAssert)
    ' Act
    Dim actual_value As Long
    actual_value = GetExcelFileFormat("sample.Csv")

    ' Assert
    Assert.EqualsNumeric xlCSV, actual_value
End Sub

' -----------------------------------------------------------------------------
' JoinStringList / JoinStringSet
' -----------------------------------------------------------------------------

Public Sub Test_JoinStringList_PassedStringList_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_lst As ObjectList
    Set src_lst = New ObjectList
    Call src_lst.Add("alpha")
    Call src_lst.Add("beta")
    Call src_lst.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringList(src_lst, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub

Public Sub Test_JoinStringSet_PassedStringSet_JoinsWithDelimiter(ByVal Assert As UnitTestAssert)
    On Error Resume Next

    ' Arrange
    Dim src_set As ObjectSet
    Set src_set = New ObjectSet
    Call src_set.Add("alpha")
    Call src_set.Add("beta")
    Call src_set.Add("gamma")

    ' Act
    Dim actual_value As String
    actual_value = JoinStringSet(src_set, ",")

    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals "alpha,beta,gamma", actual_value
End Sub
