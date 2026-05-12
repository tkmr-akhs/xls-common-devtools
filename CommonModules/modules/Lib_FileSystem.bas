Attribute VB_Name = "Lib_FileSystem"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! FileSystemService の共通インスタンスを管理するモジュールです。
'!
' #############################################################################

'* FileSystemService。ユニットテスト時にはテスト ダブルに置き換えてください。
Public FsSrv As IFileSystemService

'* FileSystemService を初期化します。
'*
'* @details
'* FsSrv が未設定の場合に FileSystemService を生成します。テストで差し替え済みの場合は維持します。
Public Sub InitializeFileSystemService()
    If FsSrv Is Nothing Then Set FsSrv = New FileSystemService
End Sub
