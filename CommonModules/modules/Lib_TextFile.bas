Attribute VB_Name = "Lib_TextFile"
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! TextFileService の共通インスタンスを管理するモジュールです。
'!
' #############################################################################

'* TextFileService。ユニットテスト時にはテスト ダブルに置き換えてください。
Public TfSrv As ITextFileService

'* TextFileService を初期化します。
'*
'* @details
'* TfSrv が未設定の場合に TextFileService を生成します。テストで差し替え済みの場合は維持します。
Public Sub InitializeTextFileService()
    If TfSrv Is Nothing Then Set TfSrv = New TextFileService
End Sub
