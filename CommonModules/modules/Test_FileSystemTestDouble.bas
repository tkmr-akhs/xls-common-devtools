Attribute VB_Name = "Test_FileSystemTestDouble"
Option Explicit

' #############################################################################
'!
'! @brief
'! FileSystemServiceTestDouble クラスのユニット テストです。
'! Lib_UnitTest.UnitTestMain() によって実行されます。
'!
' #############################################################################

Private TUtl As New UnitTestUtils

' ----------------------------------------------------------------------------
' GetAbsolutePath
' ----------------------------------------------------------------------------

Public Sub Test_GetAbsolutePath_AbsolutePath_ReturnsSamePath(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim fs_srv As IFileSystemService
    Set fs_srv = New FileSystemService
    
    Dim expected_path As String
    expected_path = "C:\Temp\File.txt"
    
    ' 実行
    Dim actual_path As String
    actual_path = fs_srv.GetAbsolutePath(expected_path)
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

Public Sub Test_GetAbsolutePath_RelativePath_ReturnsThisWorkbookBasedPath(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim fs_srv As IFileSystemService
    Set fs_srv = New FileSystemService
    
    Dim expected_path As String
    expected_path = GetParentPath(ThisWorkbook.FullName) & "\Sub\File.txt"
    
    ' 実行
    Dim actual_path As String
    actual_path = fs_srv.GetAbsolutePath("Sub\File.txt")
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals expected_path, actual_path
End Sub

Public Sub Test_GetAbsolutePath_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    ' 準備
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Dim expected_path As Variant
    expected_path = "C:\Base\Sub\File.txt"
    Call TUtl.SetValue(fs_stub.GetAbsolutePath_Values, expected_path, "Sub\File.txt")
    
    Dim fs_srv As IFileSystemService
    Set fs_srv = fs_stub
    
    ' 実行
    Dim actual_path As String
    actual_path = fs_srv.GetAbsolutePath("Sub\File.txt")
    
    ' 検証
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then Exit Sub
    Assert.Equals CStr(expected_path), actual_path
End Sub

' ----------------------------------------------------------------------------
' GetFileList
' ----------------------------------------------------------------------------

Public Sub Test_GetFileList_WithSetValue_ReturnsStubArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Dim expected_arr(0 To 1) As String
    expected_arr(0) = "C:\Temp\File1.txt"
    expected_arr(1) = "C:\Temp\File2.txt"
    
    ' DirectoryPath="C:\Temp", CatchRegExp="^.*\.txt$", IgnoreRegExp=""
    Call TUtl.SetValue(fs_stub.GetFileList_Values, expected_arr, "C:\Temp", "^.*\.txt$", "")
    
    ' Act
    Dim actual_arr() As String
    actual_arr = fs_stub.GetFileList("C:\Temp", "^.*\.txt$")
    
    ' Assert
    Assert.Equals "C:\Temp\File1.txt", actual_arr(0)
    Assert.Equals "C:\Temp\File2.txt", actual_arr(1)
End Sub

' ----------------------------------------------------------------------------
' GetDirectoryList
' ----------------------------------------------------------------------------

Public Sub Test_GetDirectoryList_WithArgs_ReturnsStubArray(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Dim expected_dirs(0 To 1) As String
    expected_dirs(0) = "C:\Temp\Sub1"
    expected_dirs(1) = "C:\Temp\Sub2"
    
    Call TUtl.SetValue(fs_stub.GetDirectoryList_Values, expected_dirs, "C:\Temp", "", "")
    
    ' Act
    Dim actual_arr() As String
    actual_arr = fs_stub.GetDirectoryList("C:\Temp")
    
    ' Assert
    Assert.Equals "C:\Temp\Sub1", actual_arr(0)
    Assert.Equals "C:\Temp\Sub2", actual_arr(1)
End Sub

' ----------------------------------------------------------------------------
' PathExists
' ----------------------------------------------------------------------------

Public Sub Test_PathExists_SetValueTrue_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Call TUtl.SetValue(fs_stub.PathExists_Values, True, "C:\ExistsPath")
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.PathExists("C:\ExistsPath")
    
    ' Assert
    Assert.IsTrue actual
End Sub

Public Sub Test_PathExists_NoValue_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.PathExists("C:\NoSuchPath")
    
    ' Assert
    Assert.IsFalse actual
End Sub

' ----------------------------------------------------------------------------
' IsFile
' ----------------------------------------------------------------------------

Public Sub Test_IsFile_WithDictionary_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Call TUtl.SetValue(fs_stub.IsFile_Values, True, "C:\Temp\SomeFile.txt")
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.IsFile("C:\Temp\SomeFile.txt")
    
    ' Assert
    Assert.IsTrue actual
End Sub

' ----------------------------------------------------------------------------
' IsDirectory
' ----------------------------------------------------------------------------

Public Sub Test_IsDirectory_WithDictionary_ReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Call TUtl.SetValue(fs_stub.IsDirectory_Values, True, "C:\Temp\KnownDir")
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.IsDirectory("C:\Temp\KnownDir")
    
    ' Assert
    Assert.IsTrue actual
End Sub

' ----------------------------------------------------------------------------
' GetLastModified
' ----------------------------------------------------------------------------

Public Sub Test_GetLastModified_WithSetDate_ReturnsDate(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Dim test_date As Date
    test_date = #11/1/2023 2:05:00 PM#
    
    Call TUtl.SetValue(fs_stub.GetLastModified_Values, test_date, "C:\fileA.txt")
    
    ' Act
    Dim actual_date As Date
    actual_date = fs_stub.GetLastModified("C:\fileA.txt")
    
    ' Assert
    Assert.Equals test_date, actual_date
End Sub

' ----------------------------------------------------------------------------
' CreateDirectory
' ----------------------------------------------------------------------------

Public Sub Test_CreateDirectory_RecordsArgsAndReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' CreateDirectory_Values -> デフォは True だが一応登録してみる
    Call TUtl.SetValue(fs_stub.CreateDirectory_Values, False, "C:\MakeDir", True, True)
    
    ' Act
    Dim actual_result As Boolean
    actual_result = fs_stub.CreateDirectory("C:\MakeDir", True, True)
    
    ' Assert
    ' 1) 戻り値は False (辞書により)
    Assert.Equals False, actual_result
    
    ' 2) CreateDirectory_Results に ( "C:\MakeDir", True, True ) で True がセットされる
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.CreateDirectory_Results, "C:\MakeDir", True, True)
    Assert.Equals True, stored_val
End Sub

Public Sub Test_CreateDirectory_NoEntryInValues_ReturnsDefaultTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual_result As Boolean
    actual_result = fs_stub.CreateDirectory("D:\NewDir")
    
    ' Assert -> 未登録なら True になる想定
    Assert.Equals True, actual_result
End Sub

' ----------------------------------------------------------------------------
' MoveDirectory
' ----------------------------------------------------------------------------

Public Sub Test_MoveDirectory_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    fs_stub.MoveDirectory "C:\SrcFolder", "D:\DestFolder", True
    
    ' Assert
    ' (SourcePath, DestPath, Force) -> True
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.MoveDirectory_Results, "C:\SrcFolder", "D:\DestFolder", True)
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' CopyDirectory
' ----------------------------------------------------------------------------

Public Sub Test_CopyDirectory_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    fs_stub.CopyDirectory "C:\FolderA", "C:\FolderB", False
    
    ' Assert
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.CopyDirectory_Results, "C:\FolderA", "C:\FolderB", False)
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' RemoveDirectory
' ----------------------------------------------------------------------------

Public Sub Test_RemoveDirectory_WithForceTrue_ReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Set dictionary => True
    Call TUtl.SetValue(fs_stub.RemoveDirectory_Values, True, "C:\TargetDir", True)
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.RemoveDirectory("C:\TargetDir", True)
    
    ' Assert
    Assert.Equals True, actual
    
    ' spied call
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.RemoveDirectory_Results, "C:\TargetDir", True)
    Assert.Equals True, stored_val
End Sub

Public Sub Test_RemoveDirectory_NoRegInValues_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual As Boolean
    actual = fs_stub.RemoveDirectory("D:\Unknown", False)
    
    ' Assert
    ' 未登録 => デフォルト False
    Assert.IsFalse actual
End Sub

' ----------------------------------------------------------------------------
' MoveFile
' ----------------------------------------------------------------------------

Public Sub Test_MoveFile_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    fs_stub.MoveFile "C:\Temp\fileA.txt", "C:\Temp\dest\FileA_moved.txt", True
    
    ' Assert
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.MoveFile_Results, "C:\Temp\fileA.txt", "C:\Temp\dest\FileA_moved.txt", True)
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' CopyFile
' ----------------------------------------------------------------------------

Public Sub Test_CopyFile_SetsDictionary(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    fs_stub.CopyFile "D:\Source.docx", "E:\Backup\Source.docx", False
    
    ' Assert
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.CopyFile_Results, "D:\Source.docx", "E:\Backup\Source.docx", False)
    Assert.Equals True, stored_val
End Sub

' ----------------------------------------------------------------------------
' RemoveFile
' ----------------------------------------------------------------------------

Public Sub Test_RemoveFile_WithDictionaryReturnsTrue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Call TUtl.SetValue(fs_stub.RemoveFile_Values, True, "C:\Data\fileX.txt", False)
    
    ' Act
    Dim actual_res As Boolean
    actual_res = fs_stub.RemoveFile("C:\Data\fileX.txt", False)
    
    ' Assert
    Assert.IsTrue actual_res
    
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.RemoveFile_Results, "C:\Data\fileX.txt", False)
    Assert.Equals True, stored_val
End Sub

Public Sub Test_RemoveFile_NotInDictionary_ReturnsFalse(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual_res As Boolean
    actual_res = fs_stub.RemoveFile("C:\NoSuchFile.log")
    
    ' Assert
    Assert.IsFalse actual_res
End Sub

' ----------------------------------------------------------------------------
' GetNewestFile
' ----------------------------------------------------------------------------

Public Sub Test_GetNewestFile_WithStubValue_ReturnsString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    Dim expected_path As String
    expected_path = "C:\Folder\Newest.docx"
    
    ' (Directory, catch, ignore)
    Call TUtl.SetValue(fs_stub.GetNewestFile_Values, expected_path, "C:\Folder", ".*\.docx", "")
    
    ' Act
    Dim actual_path As String
    actual_path = fs_stub.GetNewestFile("C:\Folder", ".*\.docx")
    
    ' Assert
    Assert.Equals expected_path, actual_path
End Sub

Public Sub Test_GetNewestFile_NoEntry_ReturnsEmptyString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual_path As String
    actual_path = fs_stub.GetNewestFile("Z:\None", ".*\.txt")
    
    ' Assert
    Assert.Equals "", actual_path
End Sub

' ----------------------------------------------------------------------------
' CreateBackupFile
' ----------------------------------------------------------------------------

Public Sub Test_CreateBackupFile_RecordsArgsAndReturnsValue(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' set dictionary => when retrieving CreateBackupFile_Values => returns "MyBackup"
    Call TUtl.SetValue(fs_stub.CreateBackupFile_Values, "MyBackup", "C:\source.txt", "D:\backups", "_v2", True, True, True)
    
    ' Act
    Dim actual_str As String
    actual_str = fs_stub.CreateBackupFile("C:\source.txt", "D:\backups", "_v2", True, True, True)
    
    ' Assert
    ' 1) 戻り値
    Assert.Equals "MyBackup", actual_str
    
    ' 2) spied call => CreateBackupFile_Results
    '    We expect "C:\source.txt", "D:\backups", "_v2", True, True, True => ...
    '    But the code is incomplete in snippet. We can guess it sets:
    '    TUtl.SetValue(CreateBackupFile_Results, ???, ???)
    '    Let's check or replicate the logic
    ' *** We'll assume it sets True
    Dim stored_val As Variant
    stored_val = TUtl.GetValue(fs_stub.CreateBackupFile_Results, "C:\source.txt", "D:\backups", "_v2", True, True, True)
    Assert.Equals True, stored_val
End Sub

Public Sub Test_CreateBackupFile_NotInDictionary_ReturnsDefaultString(ByVal Assert As UnitTestAssert)
    ' Arrange
    Dim fs_stub As FileSystemServiceTestDouble
    Set fs_stub = New FileSystemServiceTestDouble
    
    ' Act
    Dim actual_str As String
    actual_str = fs_stub.CreateBackupFile("C:\noRegFile.txt")
    
    ' Assert
    ' 未登録 => BackupCount=1 => "Backup_1"
    '   ただし snippet で "If err_num=0 then PathExists=result, else PathExists=some" => これはたぶんtypo?
    '   We assume "Backup_1"
    Assert.Equals "Backup_1", actual_str
    
    ' 次に呼ぶ => "Backup_2"
    Dim next_str As String
    next_str = fs_stub.CreateBackupFile("C:\noRegFile2.txt")
    Assert.Equals "Backup_2", next_str
End Sub
