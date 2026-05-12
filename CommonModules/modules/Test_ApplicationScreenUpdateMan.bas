Attribute VB_Name = "Test_ApplicationScreenUpdateMan"
Option Explicit

' #############################################################################
'!
'! @brief
'! ApplicationScreenUpdateManager クラスのユニット テストです。
'!
' #############################################################################

Public Sub Test_DisableUpdates_Default_DisablesAllUpdates(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager
    
    ' Act
    Call screen_manager.DisableUpdates
    
    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation
    
    Call screen_manager.Restore
    Set screen_manager = Nothing
    
    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals False, actual_screen_updating
    Assert.Equals False, actual_enable_events
    Assert.Equals False, actual_display_alerts
    Assert.EqualsNumeric xlCalculationManual, actual_calculation
    
CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub

Public Sub Test_DisableUpdates_PartialFlags_DisablesSpecifiedSettingsOnly(ByVal Assert As UnitTestAssert)
    On Error Resume Next
    
    Dim org_screen_updating As Boolean
    Dim org_enable_events As Boolean
    Dim org_display_alerts As Boolean
    Dim org_calculation As XlCalculation
    org_screen_updating = Application.ScreenUpdating
    org_enable_events = Application.EnableEvents
    org_display_alerts = Application.DisplayAlerts
    org_calculation = Application.Calculation
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    
    Dim screen_manager As ApplicationScreenUpdateManager
    Set screen_manager = New ApplicationScreenUpdateManager
    
    ' Act
    Call screen_manager.DisableUpdates( _
            StopScreenUpdating:=False, _
            StopEvents:=True, _
            StopDisplayAlerts:=False, _
            StopCalculation:=False)
    
    Dim actual_screen_updating As Boolean
    Dim actual_enable_events As Boolean
    Dim actual_display_alerts As Boolean
    Dim actual_calculation As XlCalculation
    actual_screen_updating = Application.ScreenUpdating
    actual_enable_events = Application.EnableEvents
    actual_display_alerts = Application.DisplayAlerts
    actual_calculation = Application.Calculation
    
    Call screen_manager.Restore
    Set screen_manager = Nothing
    
    ' Assert
    If Not Assert.ErrorNotRaised(0, Err.Number, Err.Source, Err.Description) Then GoTo CLEANUP
    Assert.Equals True, actual_screen_updating
    Assert.Equals False, actual_enable_events
    Assert.Equals True, actual_display_alerts
    Assert.EqualsNumeric xlCalculationAutomatic, actual_calculation
    
CLEANUP:
    Application.ScreenUpdating = org_screen_updating
    Application.EnableEvents = org_enable_events
    Application.DisplayAlerts = org_display_alerts
    Application.Calculation = org_calculation
End Sub
