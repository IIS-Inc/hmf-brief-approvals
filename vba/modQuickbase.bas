'================================================================
' modQuickbase.bas
' HMF Brief Approval System
' Quickbase REST API engine — record updates, response handling
' NOTE: QB_USER_TOKEN is installed separately by Admin via
' modUpdater.InstallToken after UpdateFromGitHub completes.
' Never commit a real token to the GitHub repository.
'================================================================

Option Explicit

'----------------------------------------------------------------
' Constants — Quickbase configuration
'----------------------------------------------------------------
Private Const QB_REALM          As String = "thehousemajoritypac.quickbase.com"
Private Const QB_TABLE_ID       As String = "bv2gw2ikx"
Private Const QB_USER_TOKEN     As String = "REPLACE_WITH_TOKEN"
Private Const QB_API_URL        As String = "https://api.quickbase.com/v1/records"

' Field IDs
Private Const FID_BRIEF_STATUS  As Long = 8
Private Const FID_DIGITAL       As Long = 10
Private Const FID_RESEARCH      As Long = 11
Private Const FID_EXECUTIVE     As Long = 12
Private Const FID_LOG           As Long = 13

'================================================================
' UpdateApprovalStatus
' Master public function — called by modApprovals after user
' confirms a status selection in the UserForm.
'================================================================
Public Function UpdateApprovalStatus(strDept As String, _
                                     strDeptStatus As String, _
                                     strBriefStatus As String) As Boolean
    ' Check token is installed
    If QB_USER_TOKEN = "REPLACE_WITH_TOKEN" Then
        MsgBox "The Quickbase API token has not been installed." & vbCrLf & vbCrLf & _
               "Please contact your administrator.", _
               vbCritical, "HMF Brief Approval — Token Not Installed"
        UpdateApprovalStatus = False
        Exit Function
    End If

    Dim lngRID As Long
    lngRID = GetRID()
    If lngRID = 0 Then
        MsgBox "Unable to update Quickbase — no Record ID is available." & vbCrLf & _
               "Please close and reopen the document to set the Record ID.", _
               vbCritical, "HMF Brief Approval"
        UpdateApprovalStatus = False
        Exit Function
    End If

    Dim lngDeptFID As Long
    lngDeptFID = GetDeptFID(strDept)
    If lngDeptFID = 0 Then
        MsgBox "Unknown department: " & strDept, vbCritical, "HMF Brief Approval"
        UpdateApprovalStatus = False
        Exit Function
    End If

    Dim strJSON As String
    strJSON = BuildPayload(lngRID, lngDeptFID, strDeptStatus, strBriefStatus)

    Dim strResponse As String
    Dim blnSuccess As Boolean
    blnSuccess = CallQuickbaseAPI(strJSON, strResponse)

    If blnSuccess Then
        SetCurrentBriefStatus strBriefStatus
        SetDepartmentStatus strDept, strDeptStatus
        ActiveDocument.Save
        UpdateApprovalStatus = True
    Else
        MsgBox "Quickbase update failed." & vbCrLf & vbCrLf & _
               "Response: " & strResponse, _
               vbCritical, "HMF Brief Approval"
        UpdateApprovalStatus = False
    End If
End Function

'================================================================
' GetDeptFID
' Maps department name to its Quickbase Field ID.
'================================================================
Private Function GetDeptFID(strDept As String) As Long
    Select Case strDept
        Case "Digital":     GetDeptFID = FID_DIGITAL
        Case "Research":    GetDeptFID = FID_RESEARCH
        Case "Executive":   GetDeptFID = FID_EXECUTIVE
        Case Else:          GetDeptFID = 0
    End Select
End Function

'================================================================
' BuildPayload
' Constructs the JSON body for the Quickbase POST request.
'================================================================
Private Function BuildPayload(lngRID As Long, _
                               lngDeptFID As Long, _
                               strDeptStatus As String, _
                               strBriefStatus As String) As String
    Dim strJSON As String

    strJSON = "{" & _
        """to"": """ & QB_TABLE_ID & """," & _
        """data"": [{" & _
            """3"": {""value"": " & lngRID & "}," & _
            """" & FID_BRIEF_STATUS & """: {""value"": """ & EscapeJSON(strBriefStatus) & """}," & _
            """" & lngDeptFID & """: {""value"": """ & EscapeJSON(strDeptStatus) & """}," & _
            """" & FID_LOG & """: {""value"": """ & EscapeJSON(strBriefStatus) & """}" & _
        "}]" & _
    "}"

    BuildPayload = strJSON
End Function

'================================================================
' CallQuickbaseAPI
' Makes the HTTP POST request to the Quickbase Records API.
'================================================================
Private Function CallQuickbaseAPI(strJSON As String, _
                                   ByRef strResponse As String) As Boolean
    Dim http        As Object
    Dim blnSuccess  As Boolean

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    With http
        .Open "POST", QB_API_URL, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "QB-Realm-Hostname", QB_REALM
        .setRequestHeader "Authorization", "QB-USER-TOKEN " & QB_USER_TOKEN
        .setRequestHeader "User-Agent", "HMF-BriefApproval-Word/1.0"
        .send strJSON

        strResponse = .responseText

        If .Status = 200 Then
            blnSuccess = True
        Else
            blnSuccess = False
            strResponse = "HTTP " & .Status & ": " & strResponse
        End If
    End With

    Set http = Nothing
    CallQuickbaseAPI = blnSuccess
    Exit Function

ErrorHandler:
    strResponse = "Connection error: " & Err.Description
    CallQuickbaseAPI = False
    If Not http Is Nothing Then Set http = Nothing
End Function

'================================================================
' EscapeJSON
'================================================================
Private Function EscapeJSON(strInput As String) As String
    Dim strOut As String
    strOut = strInput
    strOut = Replace(strOut, "\", "\\")
    strOut = Replace(strOut, """", "\""")
    strOut = Replace(strOut, "/", "\/")
    strOut = Replace(strOut, Chr(8), "\b")
    strOut = Replace(strOut, Chr(9), "\t")
    strOut = Replace(strOut, Chr(10), "\n")
    strOut = Replace(strOut, Chr(12), "\f")
    strOut = Replace(strOut, Chr(13), "\r")
    EscapeJSON = strOut
End Function

'================================================================
' IsTokenInstalled
' Public helper — returns True if token has been installed.
' Called by modUpdater.InstallToken to verify success.
'================================================================
Public Function IsTokenInstalled() As Boolean
    IsTokenInstalled = (QB_USER_TOKEN <> "REPLACE_WITH_TOKEN" And QB_USER_TOKEN <> "")
End Function

'================================================================
' TestConnection
' Dev utility — tests API connectivity.
' Run from Immediate Window: TestConnection
'================================================================
Public Sub TestConnection()
    If Not IsTokenInstalled() Then
        Debug.Print "TEST FAILED: Token not installed — run InstallToken first"
        Exit Sub
    End If

    Dim lngRID As Long
    lngRID = GetRID()

    If lngRID = 0 Then
        Debug.Print "TEST FAILED: No RID available"
        Exit Sub
    End If

    Dim strCurrentStatus As String
    strCurrentStatus = GetCurrentBriefStatus()

    If strCurrentStatus = "" Then
        Debug.Print "TEST FAILED: No current status set"
        Exit Sub
    End If

    Debug.Print "Testing Quickbase connection..."
    Debug.Print "  RID:    " & lngRID
    Debug.Print "  Status: " & strCurrentStatus

    Dim strJSON As String
    strJSON = BuildPayload(lngRID, FID_DIGITAL, _
                           GetCurrentDepartmentStatus("Digital"), _
                           strCurrentStatus)

    Debug.Print "  Payload: " & strJSON

    Dim strResponse As String
    Dim blnSuccess As Boolean
    blnSuccess = CallQuickbaseAPI(strJSON, strResponse)

    If blnSuccess Then
        Debug.Print "TEST PASSED: Quickbase responded 200 OK"
        Debug.Print "  Response: " & strResponse
    Else
        Debug.Print "TEST FAILED: " & strResponse
    End If
End Sub