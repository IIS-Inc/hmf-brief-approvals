'================================================================
' modApprovals.bas
' HMF Brief Approval System
' Approval logic — status mapping, workflow rules, UserForm controller
'================================================================

Option Explicit

'================================================================
' GetAvailableStatuses
' Returns an array of status options appropriate for the current
' stage and department. These populate the UserForm listbox.
'
' Parameters:
'   strDept — "Digital", "Research", or "Executive"
'================================================================
Public Function GetAvailableStatuses(strDept As String) As String()
    Dim arrStatuses() As String

    Select Case strDept
        Case "Digital"
            ReDim arrStatuses(3)
            arrStatuses(0) = "Pending Digital review"
            arrStatuses(1) = "Editing"
            arrStatuses(2) = "Rejected"
            arrStatuses(3) = "Approved"

        Case "Research"
            ReDim arrStatuses(4)
            arrStatuses(0) = "Pending Research review"
            arrStatuses(1) = "Editing"
            arrStatuses(2) = "Rejected"
            arrStatuses(3) = "Needs Legal review"
            arrStatuses(4) = "Approved"

        Case "Executive"
            ReDim arrStatuses(4)
            arrStatuses(0) = "Pending Executive review"
            arrStatuses(1) = "Needs Digital edits"
            arrStatuses(2) = "Needs Research/Legal edits"
            arrStatuses(3) = "Rejected"
            arrStatuses(4) = "Approved"

        Case Else
            ReDim arrStatuses(0)
            arrStatuses(0) = ""
    End Select

    GetAvailableStatuses = arrStatuses
End Function

'================================================================
' GetBriefStatusForSelection
' Maps a department + unprefixed selection to the correct
' prefixed Brief Approval Status value (FID 8).
'
' Parameters:
'   strDept      — "Digital", "Research", or "Executive"
'   strSelection — unprefixed status (e.g. "Approved")
'
' Returns the prefixed master status (e.g. "Digital: Approved")
'================================================================
Public Function GetBriefStatusForSelection(strDept As String, _
                                            strSelection As String) As String
    Select Case strDept
        Case "Digital"
            Select Case strSelection
                Case "Pending Digital review":  GetBriefStatusForSelection = "Pending Digital review"
                Case "Editing":                 GetBriefStatusForSelection = "Digital: Editing"
                Case "Rejected":                GetBriefStatusForSelection = "Digital: Rejected"
                Case "Approved":                GetBriefStatusForSelection = "Digital: Approved"
                Case Else:                      GetBriefStatusForSelection = ""
            End Select

        Case "Research"
            Select Case strSelection
                Case "Pending Research review": GetBriefStatusForSelection = "Pending Research review"
                Case "Editing":                 GetBriefStatusForSelection = "Research: Editing"
                Case "Rejected":                GetBriefStatusForSelection = "Research: Rejected"
                Case "Needs Legal review":      GetBriefStatusForSelection = "Research: Needs legal review"
                Case "Approved":                GetBriefStatusForSelection = "Research: Approved"
                Case Else:                      GetBriefStatusForSelection = ""
            End Select

        Case "Executive"
            Select Case strSelection
                Case "Pending Executive review":        GetBriefStatusForSelection = "Pending Executive review"
                Case "Needs Digital edits":             GetBriefStatusForSelection = "Executive: Needs digital edits"
                Case "Needs Research/Legal edits":      GetBriefStatusForSelection = "Executive: Needs Research/Legal edits"
                Case "Rejected":                        GetBriefStatusForSelection = "Executive: Rejected"
                Case "Approved":                        GetBriefStatusForSelection = "Executive: Approved"
                Case Else:                              GetBriefStatusForSelection = ""
            End Select

        Case Else
            GetBriefStatusForSelection = ""
    End Select
End Function

'================================================================
' CanDeptAct
' Enforces the linear workflow gate.
' Returns True if the given department is allowed to act
' based on the current stage number.
'
' Rules:
'   Digital   — always active until Research takes over (stage < 2)
'   Research  — active only after Digital approves (stage >= 1.5)
'   Executive — active only after Research approves (stage >= 2.5)
'================================================================
Public Function CanDeptAct(strDept As String) As Boolean
    Select Case strDept
        Case "Digital":     CanDeptAct = IsDigitalActive()
        Case "Research":    CanDeptAct = IsResearchActive()
        Case "Executive":   CanDeptAct = IsExecutiveActive()
        Case Else:          CanDeptAct = False
    End Select
End Function

'================================================================
' LaunchApprovalForm
' Entry point called by the Custom Ribbon buttons.
' Validates stage gate, then opens the UserForm for the
' appropriate department.
'
' Parameters:
'   strDept — "Digital", "Research", or "Executive"
'================================================================
Public Sub LaunchApprovalForm(strDept As String)

    ' Ensure document is initialized
    If Not AllPropertiesInitialized() Then
        MsgBox "This document has not been initialized." & vbCrLf & _
               "Please close and reopen to set the Record ID.", _
               vbExclamation, "HMF Brief Approval"
        Exit Sub
    End If

    ' Enforce workflow gate
    If Not CanDeptAct(strDept) Then
        Dim strCurrent As String
        strCurrent = GetCurrentBriefStatus()
        MsgBox strDept & " approval is not available at this stage." & vbCrLf & vbCrLf & _
               "Current status: " & strCurrent & vbCrLf & vbCrLf & _
               GetGateMessage(strDept), _
               vbInformation, "HMF Brief Approval"
        Exit Sub
    End If

    ' Open the UserForm for this department
    Dim frm As frmApprovals
    Set frm = New frmApprovals
    frm.InitializeForm strDept
    frm.Show
    Set frm = Nothing
End Sub

'================================================================
' GetGateMessage
' Returns a helpful explanation of why a department is locked.
'================================================================
Private Function GetGateMessage(strDept As String) As String
    Select Case strDept
        Case "Digital"
            GetGateMessage = "Digital review is no longer active — " & _
                             "the brief has progressed past the Digital stage."
        Case "Research"
            GetGateMessage = "Research review is not yet available." & vbCrLf & _
                             "Digital must approve the brief before Research can act."
        Case "Executive"
            GetGateMessage = "Executive review is not yet available." & vbCrLf & _
                             "Research must approve the brief before Executive can act."
        Case Else
            GetGateMessage = ""
    End Select
End Function

'================================================================
' SubmitApproval
' Called by the UserForm OK/Submit button.
' Validates the selection, maps to Brief status, calls the API.
'
' Parameters:
'   strDept      — "Digital", "Research", or "Executive"
'   strSelection — the status chosen in the UserForm listbox
'
' Returns True on successful Quickbase update.
'================================================================
Public Function SubmitApproval(strDept As String, _
                                strSelection As String) As Boolean
    ' Validate selection was made
    If strSelection = "" Then
        MsgBox "Please select a status before submitting.", _
               vbExclamation, "HMF Brief Approval"
        SubmitApproval = False
        Exit Function
    End If

    ' Map to Brief status
    Dim strBriefStatus As String
    strBriefStatus = GetBriefStatusForSelection(strDept, strSelection)

    If strBriefStatus = "" Then
        MsgBox "Invalid status selection: " & strSelection, _
               vbCritical, "HMF Brief Approval"
        SubmitApproval = False
        Exit Function
    End If

    ' Confirm with user before writing to Quickbase
    Dim strConfirm As String
    strConfirm = "You are about to update the Brief Approval Status:" & vbCrLf & vbCrLf & _
                 "   Department:  " & strDept & vbCrLf & _
                 "   New Status:  " & strSelection & vbCrLf & vbCrLf & _
                 "This will update Quickbase immediately." & vbCrLf & _
                 "Do you want to proceed?"

    If MsgBox(strConfirm, vbQuestion + vbYesNo, "Confirm Approval Update") = vbNo Then
        SubmitApproval = False
        Exit Function
    End If

    ' Call the API
    Dim blnSuccess As Boolean
    blnSuccess = UpdateApprovalStatus(strDept, strSelection, strBriefStatus)

    If blnSuccess Then
        MsgBox "Quickbase updated successfully!" & vbCrLf & vbCrLf & _
               "Brief Status: " & strBriefStatus, _
               vbInformation, "HMF Brief Approval"
        SubmitApproval = True
    Else
        SubmitApproval = False
    End If
End Function

'================================================================
' GetDeptStatusSummary
' Returns a formatted summary of all three department statuses.
' Used by the UserForm to display current state.
'================================================================
Public Function GetDeptStatusSummary() As String
    Dim strSummary As String
    strSummary = "Current Approval Status" & vbCrLf & _
                 String(30, "-") & vbCrLf & _
                 "Brief:     " & GetCurrentBriefStatus() & vbCrLf & _
                 "Digital:   " & GetCurrentDepartmentStatus("Digital") & vbCrLf & _
                 "Research:  " & GetCurrentDepartmentStatus("Research") & vbCrLf & _
                 "Executive: " & GetCurrentDepartmentStatus("Executive")
    GetDeptStatusSummary = strSummary
End Function