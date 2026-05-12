VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationScreenUpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

' #############################################################################
'!
'! @brief
'! アプリケーションの画面更新などを管理するクラスです。
'! インスタンス生成時点での Excel の設定をバックアップし、元に戻せるようにします。
'! `As New` で宣言された変数は、初めてアクセスされたときにインスタンスが生成されることになることに注意が必要です。
'! `Dim ASUM As New ApplicationScreenUpdateManager` の直後に `ASUM.DisableUpdates` を呼び出すか、
'! `Dim ASUM As ApplicationScreenUpdateManager: Set ASUM = New ApplicationScreenUpdateManager` とすることを推奨します。
'! インスタンス破棄時には自動的に Restore されるため、コードの途中でエラーが発生しても Excel の設定が元に戻るようになっています。
'! Esc キー入力による中断や `End` ステートメントによる中断など、インスタンスが破棄されない場合は、Restore メソッドを明示的に呼び出す必要があります。
'!
' #############################################################################

Private pSavedScreenUpdating As Boolean
Private pSavedEnableEvents As Boolean
Private pSavedDisplayAlerts As Boolean
Private pSavedCalculation As Integer
Private pIsBackedUp As Boolean

'* このオブジェクトに記録された ScreenUpdating の値。
Public Property Get SavedScreenUpdating() As Boolean
    Call pEnsureBackupExists
    SavedScreenUpdating = pSavedScreenUpdating
End Property

'* このオブジェクトに記録された EnableEvents の値。
Public Property Get SavedEnableEvents() As Boolean
    Call pEnsureBackupExists
    SavedEnableEvents = pSavedEnableEvents
End Property

'* このオブジェクトに記録された DisplayAlerts の値。
Public Property Get SavedDisplayAlerts() As Boolean
    Call pEnsureBackupExists
    SavedDisplayAlerts = pSavedDisplayAlerts
End Property

'* このオブジェクトに記録された Calculation の値。
Public Property Get SavedCalculation() As Integer
    Call pEnsureBackupExists
    SavedCalculation = pSavedCalculation
End Property

Private Sub Class_Initialize()
    Call pBackup
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Restore
    On Error GoTo 0
End Sub

' このオブジェクトに、各種設定をバックアップします。
'
' @details
' 現在の Excel の設定 (ScreenUpdating、EnableEvents、DisplayAlerts、Calculation) を記録し、
' Restore メソッドで元の設定に戻せるようにします。
Private Sub pBackup()
    pSavedScreenUpdating = Application.ScreenUpdating
    pSavedEnableEvents = Application.EnableEvents
    pSavedDisplayAlerts = Application.DisplayAlerts
    pSavedCalculation = Application.Calculation
    pIsBackedUp = True
End Sub

'* 画面アップデートやイベントのトリガーなどを停止します。
'*
'* @param StopScreenUpdating 画面の更新を停止するかどうか (デフォルトは True)
'* @param StopEvents イベントのトリガーを停止するかどうか (デフォルトは True)
'* @param StopDisplayAlerts 警告ダイアログの表示を停止するかどうか (デフォルトは True)
'* @param StopCalculation 計算を手動モードにするかどうか (デフォルトは True)
'*
'* @details
'* ScreenUpdating、EnableEvents、DisplayAlerts を無効化し、
'* Calculation を手動計算モードに設定します。
Public Sub DisableUpdates( _
            Optional ByVal StopScreenUpdating As Boolean = True, _
            Optional ByVal StopEvents As Boolean = True, _
            Optional ByVal StopDisplayAlerts As Boolean = True, _
            Optional ByVal StopCalculation As Boolean = True)
    Call pEnsureBackupExists
    If StopScreenUpdating Then Application.ScreenUpdating = False
    If StopEvents Then Application.EnableEvents = False
    If StopDisplayAlerts Then Application.DisplayAlerts = False
    If StopCalculation Then Application.Calculation = xlCalculationManual
End Sub

'* このオブジェクトに保存された設定を明示的にリストアします。
'*
'* @details
'* 保存された設定を Excel に復元します。
'* バックアップがまだ存在しない場合はエラーを発生させます。
Public Sub Restore()
    Call pEnsureBackupExists
    Application.ScreenUpdating = pSavedScreenUpdating
    Application.EnableEvents = pSavedEnableEvents
    Application.DisplayAlerts = pSavedDisplayAlerts
    Application.Calculation = pSavedCalculation
    pIsBackedUp = False
End Sub

Private Sub pEnsureBackupExists()
    If Not pIsBackedUp Then
        Err.Raise vbObjectError + 1, "Class ApplicationScreenUpdateManager", "バックアップされた設定がありません。"
    End If
End Sub
