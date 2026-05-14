VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystemServiceTestDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' #############################################################################
'!
'! @brief
'! ファイル システムトへのアクセスを管理するサービスのテストダブルです。
'!
' #############################################################################

Implements IFileSystemService


'* GetAbsolutePath 引数→戻り値辞書。未登録は引数をそのまま返します。
Public GetAbsolutePath_Values As Dictionary

'* GetFileList 引数→戻り値辞書。
Public GetFileList_Values As Dictionary

'* GetDirectoryList 引数→戻り値辞書。
Public GetDirectoryList_Values As Dictionary

'* PathExists 引数→戻り値辞書。未登録は False として扱われます。
Public PathExists_Values As Dictionary

'* IsFile 引数→戻り値辞書。未登録は False として扱われます。
Public IsFile_Values As Dictionary

'* IsDirectory 引数→戻り値辞書。未登録は False として扱われます。
Public IsDirectory_Values As Dictionary

'* GetLastModified 引数→戻り値辞書。
Public GetLastModified_Values As Dictionary

'* CreateDirectory 引数→受け取り値辞書。
Public CreateDirectory_Results As Dictionary

'* CreateDirectory 引数→戻り値辞書。未登録は True を返します。
Public CreateDirectory_Values As Dictionary

'* MoveDirectory 引数→受け取り値辞書。
Public MoveDirectory_Results As Dictionary

'* CopyDirectory 引数→受け取り値辞書。
Public CopyDirectory_Results As Dictionary

'* RemoveDirectory 引数→受け取り値辞書。
Public RemoveDirectory_Results As Dictionary

'* RemoveDirectory 引数→戻り値辞書。未登録は False を返します。
Public RemoveDirectory_Values As Dictionary

'* MoveFile 引数→受け取り値辞書。
Public MoveFile_Results As Dictionary

'* CopyFile 引数→受け取り値辞書。
Public CopyFile_Results As Dictionary

'* RemoveFile 引数→受け取り値辞書。
Public RemoveFile_Results As Dictionary

'* RemoveFile 引数→戻り値辞書。未登録は False を返します。
Public RemoveFile_Values As Dictionary

'* GetNewestFile 引数→戻り値辞書。
Public GetNewestFile_Values As Dictionary

'* CreateBackupFile 引数→受け取り値辞書。
Public CreateBackupFile_Results As Dictionary

'* CreateBackupFile 引数→戻り値辞書。未登録はランダム文字列を返します。
Public CreateBackupFile_Values As Dictionary

Public BackupCount As Long

Private TUtl As UnitTestUtils

Private Sub Class_Initialize()
    Set TUtl = New UnitTestUtils
    
    Set GetAbsolutePath_Values = New Dictionary
    
    Set GetFileList_Values = New Dictionary
    Set GetDirectoryList_Values = New Dictionary
    Set PathExists_Values = New Dictionary
    Set IsFile_Values = New Dictionary
    Set IsDirectory_Values = New Dictionary
    Set GetLastModified_Values = New Dictionary
    Set CreateDirectory_Results = New Dictionary
    Set CreateDirectory_Values = New Dictionary
    Set MoveDirectory_Results = New Dictionary
    Set CopyDirectory_Results = New Dictionary
    Set RemoveDirectory_Results = New Dictionary
    Set RemoveDirectory_Values = New Dictionary
    Set MoveFile_Results = New Dictionary
    Set CopyFile_Results = New Dictionary
    Set RemoveFile_Results = New Dictionary
    Set RemoveFile_Values = New Dictionary
    Set GetNewestFile_Values = New Dictionary
    Set CreateBackupFile_Results = New Dictionary
    Set CreateBackupFile_Values = New Dictionary
    
    BackupCount = 1
End Sub


'* GetAbsolutePath のスタブです。
'*
'* @details
'* GetAbsolutePath_Values に登録済みの戻り値があれば返します。未登録の場合は引数をそのまま返します。
Public Function GetAbsolutePath(ByVal PathLikeString As String) As String
    If TUtl.HasValue(GetAbsolutePath_Values, PathLikeString) Then
        GetAbsolutePath = TUtl.GetValue(GetAbsolutePath_Values, PathLikeString)
    Else
        GetAbsolutePath = PathLikeString
    End If
End Function
Private Function IFileSystemService_GetAbsolutePath(ByVal PathLikeString As String) As String
    IFileSystemService_GetAbsolutePath = GetAbsolutePath(PathLikeString)
End Function

'* GetFileList のスタブです。
'*
'* @details
'* GetFileList = TUtl.GetValue(GetFileList_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
Public Function GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String()
    GetFileList = TUtl.GetValue(GetFileList_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
End Function
Private Function IFileSystemService_GetFileList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String()
    IFileSystemService_GetFileList = GetFileList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp)
End Function

'* GetDirectoryList のスタブです。
'*
'* @details
'* GetDirectoryList = TUtl.GetValue(GetDirectoryList_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
Public Function GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String()
    GetDirectoryList = TUtl.GetValue(GetDirectoryList_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
End Function
Private Function IFileSystemService_GetDirectoryList(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String()
    IFileSystemService_GetDirectoryList = GetDirectoryList(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp)
End Function

'* PathExists のスタブです。
'*
'* @details
'* 指定されたパスに実体が存在するかを確認します。ファイルかディレクトリがある場合に True となります。
Public Function PathExists(ByVal PathLikeString As String) As Boolean
    If TUtl.HasValue(PathExists_Values, PathLikeString) Then
        PathExists = TUtl.GetValue(PathExists_Values, PathLikeString)
    Else
        PathExists = False
    End If
End Function
Private Function IFileSystemService_PathExists(ByVal PathLikeString As String) As Boolean
    IFileSystemService_PathExists = PathExists(PathLikeString)
End Function

'* IsFile のスタブです。
'*
'* @details
'* 指定されたパスにファイルが存在するかを確認します。
Public Function IsFile(ByVal FilePath As String) As Boolean
    If TUtl.HasValue(IsFile_Values, FilePath) Then
        IsFile = TUtl.GetValue(IsFile_Values, FilePath)
    Else
        IsFile = False
    End If
End Function
Private Function IFileSystemService_IsFile(ByVal FilePath As String) As Boolean
    IFileSystemService_IsFile = IsFile(FilePath)
End Function

'* IsDirectory のスタブです。
'*
'* @details
'* result = TUtl.GetValue(IsDirectory_Values, DirectoryPath)
Public Function IsDirectory(ByVal DirectoryPath As String) As Boolean
    If TUtl.HasValue(IsDirectory_Values, DirectoryPath) Then
        IsDirectory = TUtl.GetValue(IsDirectory_Values, DirectoryPath)
    Else
        IsDirectory = False
    End If
End Function
Private Function IFileSystemService_IsDirectory(ByVal DirectoryPath As String) As Boolean
    IFileSystemService_IsDirectory = IsDirectory(DirectoryPath)
End Function

'* GetLastModified のスタブです。
'*
'* @details
'* GetLastModified = TUtl.GetValue(GetLastModified_Values,PathLikeString)
Public Function GetLastModified(ByVal PathLikeString As String) As Date
    GetLastModified = TUtl.GetValue(GetLastModified_Values, PathLikeString)
End Function
Private Function IFileSystemService_GetLastModified(ByVal PathLikeString As String) As Date
    IFileSystemService_GetLastModified = GetLastModified(PathLikeString)
End Function

'* CreateDirectory のスパイ兼スタブです。
'*
'* @details
'*
Public Function CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    Call TUtl.SetValue(CreateDirectory_Results, True, DirectoryPath, Force, Recursive)
    
    If TUtl.HasValue(CreateDirectory_Values, DirectoryPath, Force, Recursive) Then
         CreateDirectory = TUtl.GetValue(CreateDirectory_Values, DirectoryPath, Force, Recursive)
    Else
        CreateDirectory = True
    End If
End Function
Private Function IFileSystemService_CreateDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False, Optional ByVal Recursive As Boolean = False) As Boolean
    IFileSystemService_CreateDirectory = CreateDirectory(DirectoryPath, Force:=Force, Recursive:=Recursive)
End Function

'* MoveDirectory のスパイです。
'*
'* @details
'* Call TUtl.SetValue(MoveDirectory_Results, True, SourcePath, DestinationPath, Force)
Public Sub MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call TUtl.SetValue(MoveDirectory_Results, True, SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_MoveDirectory(ByVal SourcePath As String, ByVal DestinationPath As String, Optional ByVal Force As Boolean = False)
    Call MoveDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

'* CopyDirectory のスパイです。
'*
'* @details
'* Call TUtl.SetValue(CopyDirectory_Results, True, SourcePath, DestinationPath, Force)
Public Sub CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call TUtl.SetValue(CopyDirectory_Results, True, SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_CopyDirectory(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call CopyDirectory(SourcePath, DestinationPath, Force:=Force)
End Sub

'* RemoveDirectory のスパイです。
'*
'* @details
'*
Public Function RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    Call TUtl.SetValue(RemoveDirectory_Results, True, DirectoryPath, Force)
    
    If TUtl.HasValue(RemoveDirectory_Values, DirectoryPath, Force) Then
        RemoveDirectory = TUtl.GetValue(RemoveDirectory_Values, DirectoryPath, Force)
    Else
        RemoveDirectory = False
    End If
End Function
Private Function IFileSystemService_RemoveDirectory(ByVal DirectoryPath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveDirectory = RemoveDirectory(DirectoryPath, Force:=Force)
End Function

'* MoveFile のスパイです。
'*
'* @details
'* Call TUtl.SetValue(MoveFile_Results, True, SourcePath, DestinationPath, Force)
Public Sub MoveFile(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call TUtl.SetValue(MoveFile_Results, True, SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_MoveFile(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call MoveFile(SourcePath, DestinationPath, Force:=Force)
End Sub

'* CopyFile のスパイです。
'*
'* @details
'* 指定されたパスにファイルを作成します。DestinationPath がパス区切り文字で終わる場合、そのディレクトリ配下のファイルとして、現在のファイル名のままコピーします。
Public Sub CopyFile(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call TUtl.SetValue(CopyFile_Results, True, SourcePath, DestinationPath, Force)
End Sub
Private Sub IFileSystemService_CopyFile(ByVal SourcePath As String, ByVal DestinationPath As Variant, Optional ByVal Force As Boolean = False)
    Call CopyFile(SourcePath, DestinationPath, Force:=Force)
End Sub

'* RemoveFile のスパイ兼スタブです。
'*
'* @details
'*
Public Function RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    Call TUtl.SetValue(RemoveFile_Results, True, FilePath, Force)
    
    If TUtl.HasValue(RemoveFile_Values, FilePath, Force) Then
        RemoveFile = TUtl.GetValue(RemoveFile_Values, FilePath, Force)
    Else
        RemoveFile = False
    End If
End Function
Private Function IFileSystemService_RemoveFile(ByVal FilePath As String, Optional ByVal Force As Boolean = False) As Boolean
    IFileSystemService_RemoveFile = RemoveFile(FilePath, Force:=Force)
End Function

'* GetNewestFile のスタブです。
'*
'* @details
'* GetNewestFile = TUtl.GetValue(GetNewestFile_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
Public Function GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String
    If TUtl.HasValue(GetNewestFile_Values, DirectoryPath, CatchRegExp, IgnoreRegExp) Then
        GetNewestFile = TUtl.GetValue(GetNewestFile_Values, DirectoryPath, CatchRegExp, IgnoreRegExp)
    Else
        GetNewestFile = ""
    End If
End Function
Private Function IFileSystemService_GetNewestFile(ByVal DirectoryPath As String, Optional ByVal CatchRegExp As String = "", Optional ByVal IgnoreRegExp As String = "") As String
    IFileSystemService_GetNewestFile = GetNewestFile(DirectoryPath, CatchRegExp:=CatchRegExp, IgnoreRegExp:=IgnoreRegExp)
End Function

'* CreateBackupFile のスパイ兼スタブです。ファイル名を省略したときは "Backup_" & BackupCount が返ります。
'*
'* @details
'*
Public Function CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String
    
    Call TUtl.SetValue(CreateBackupFile_Results, True, SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound)
    
    If TUtl.HasValue(CreateBackupFile_Values, SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound) Then
        CreateBackupFile = TUtl.GetValue(CreateBackupFile_Values, SourceFilePath, DestinationDirectoryPath, BackupSuffix, Force, Move, FailIfSourceNotFound)
    Else
        CreateBackupFile = "Backup_" & BackupCount
        BackupCount = BackupCount + 1
    End If
End Function
Private Function IFileSystemService_CreateBackupFile( _
        Optional ByVal SourceFilePath As String = "", _
        Optional ByVal DestinationDirectoryPath As String = "", _
        Optional ByVal BackupSuffix As String = "", _
        Optional ByVal Force As Boolean = False, _
        Optional ByVal Move As Boolean = False, _
        Optional ByVal FailIfSourceNotFound As Boolean = True) As String
    
    IFileSystemService_CreateBackupFile = CreateBackupFile( _
            SourceFilePath:=SourceFilePath, _
            DestinationDirectoryPath:=DestinationDirectoryPath, _
            BackupSuffix:=BackupSuffix, _
            Force:=Force, Move:=Move, FailIfSourceNotFound:=FailIfSourceNotFound)
End Function
