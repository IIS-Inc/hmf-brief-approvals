'================================================================
' modRibbon.bas
' HMF Brief Approval System
' Custom ribbon callbacks — visibility, state, and actions
'================================================================

Option Explicit

' Ribbon object reference — stored on load for InvalidateControl
Private m_ribbon As Object

'================================================================
' Ribbon_OnLoad
' Fires when the ribbon loads. Stores ribbon reference.
'================================================================
Public Sub Ribbon_OnLoad(ribbon As Object)
    Set m_ribbon = ribbon
End Sub

'================================================================
' Ribbon_Refresh
' Called by the Refresh button — forces ribbon to re-evaluate
' all callbacks. Use after a status update.
'================================================================
Public Sub Ribbon_Refresh(control As Object)
    If Not m_ribbon Is Nothing Then
        m_ribbon.Invalidate
    End If
End Sub

'================================================================
' InvalidateRibbon
' Called internally after a successful status update to
' refresh ribbon state without user clicking Refresh.
'================================================================
Public Sub InvalidateRibbon()
    If Not m_ribbon Is Nothing Then
        m_ribbon.Invalidate
    End If
End Sub

'================================================================
' DIGITAL CALLBACKS
'================================================================
Public Sub Digital_GetVisible(control As Object, ByRef visible As Variant)
    visible = IsButtonVisible(ROLE_DIGITAL)
End Sub

Public Sub Digital_GetEnabled(control As Object, ByRef enabled As Variant)
    enabled = CanUserAct(ROLE_DIGITAL)
End Sub

Public Sub Digital_OnAction(control As Object)
    LaunchApprovalForm ROLE_DIGITAL
    InvalidateRibbon
End Sub

Public Sub Digital_GetStatusLabel(control As Object, ByRef label As Variant)
    Dim strStatus As String
    strStatus = GetCurrentDepartmentStatus("Digital")
    If strStatus = "" Then
        label = "Status: Not set"
    Else
        label = "Status: " & strStatus
    End If
End Sub

'================================================================
' RESEARCH CALLBACKS
'================================================================
Public Sub Research_GetVisible(control As Object, ByRef visible As Variant)
    visible = IsButtonVisible(ROLE_RESEARCH)
End Sub

Public Sub Research_GetEnabled(control As Object, ByRef enabled As Variant)
    enabled = CanUserAct(ROLE_RESEARCH)
End Sub

Public Sub Research_OnAction(control As Object)
    LaunchApprovalForm ROLE_RESEARCH
    InvalidateRibbon
End Sub

Public Sub Research_GetStatusLabel(control As Object, ByRef label As Variant)
    Dim strStatus As String
    strStatus = GetCurrentDepartmentStatus("Research")
    If strStatus = "" Then
        label = "Status: Not set"
    Else
        label = "Status: " & strStatus
    End If
End Sub

'================================================================
' EXECUTIVE CALLBACKS
'================================================================
Public Sub Executive_GetVisible(control As Object, ByRef visible As Variant)
    visible = IsButtonVisible(ROLE_EXECUTIVE)
End Sub

Public Sub Executive_GetEnabled(control As Object, ByRef enabled As Variant)
    enabled = CanUserAct(ROLE_EXECUTIVE)
End Sub

Public Sub Executive_OnAction(control As Object)
    LaunchApprovalForm ROLE_EXECUTIVE
    InvalidateRibbon
End Sub

Public Sub Executive_GetStatusLabel(control As Object, ByRef label As Variant)
    Dim strStatus As String
    strStatus = GetCurrentDepartmentStatus("Executive")
    If strStatus = "" Then
        label = "Status: Not set"
    Else
        label = "Status: " & strStatus
    End If
End Sub

'================================================================
' BRIEF INFO CALLBACKS
'================================================================
Public Sub Brief_GetStatusLabel(control As Object, ByRef label As Variant)
    Dim strStatus As String
    strStatus = GetCurrentBriefStatus()
    If strStatus = "" Then
        label = "Brief: Not initialized"
    Else
        label = "Brief: " & strStatus
    End If
End Sub

Public Sub Brief_GetRIDLabel(control As Object, ByRef label As Variant)
    Dim lngRID As Long
    lngRID = GetRID()
    If lngRID = 0 Then
        label = "HMF ID: Not set"
    Else
        label = "HMF ID: " & lngRID
    End If
End Sub