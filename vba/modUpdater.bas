'================================================================
' modUpdater.bas
' HMF Brief Approval System
' GitHub auto-update engine — Admin only
' Pulls latest .bas modules from the public GitHub repository
' and reimports them into the VBA project.
'================================================================

Option Explicit

'----------------------------------------------------------------
' Constants — GitHub repository configuration
'----------------------------------------------------------------
Private Const GITHUB_RAW_BASE   As String = "https://raw.githubusercontent.com/IIS-Inc/hmf-brief-approvals/main/vba/"
Private Const TEMP_PATH         As String = "C:\Windows\Temp\HMF_Update\"

' Modules managed by the updater
' modUpdater is excluded — cannot update itself
Private Const MODULE_COUNT      As Integer = 5

'================================================================
' UpdateFromGitHub
' Master public function — called by Admin from Immediate Window
' or triggered by a ribbon button (Admin only).
' Downloads and reimports all managed modules from GitHub.
'================================================================
Public Sub UpdateFromGitHub()

    ' Enforce Admin role
    If GetCurrentUserRole() <> ROLE_ADMIN Then
        MsgBox "The auto-update feature is available to Admin users only.", _
               vbExclamation, "HMF Brief Approval — Access Denied"
        Exit Sub
    End If

    ' Confirm with user
    Dim strConfirm As String
    strConfirm = "This will update all VBA modules from GitHub." & vbCrLf & vbCrLf & _
                 "Source: " & GITHUB_RAW_BASE & vbCrLf & vbCrLf & _
                 "The following modules will be updated:" & vbCrLf & _
                 "  - modUtilities" & vbCrLf & _
                 "  - modQuickbase" & vbCrLf & _
                 "  - modApprovals" & vbCrLf & _
                 "  - modRoles" & vbCrLf & _
                 "  - modRibbon" & vbCrLf & vbCrLf & _
                 "NOTE: modUpdater cannot update itself." & vbCrLf & vbCrLf & _
                 "The document will be saved after a successful update." & vbCrLf & vbCrLf & _
                 "Do you want to proceed?"

    If MsgBox(strConfirm, vbQuestion + vbYesNo, "HMF Brief Approval — Update from GitHub") = vbNo Then
        Exit Sub
    End If

    ' Verify VBA project object model is accessible
    If Not IsVBAAccessible() Then
        MsgBox "The VBA project object model is not accessible." & vbCrLf & vbCrLf & _
               "To enable it:" & vbCrLf & _
               "File → Options → Trust Center → Trust Center Settings" & vbCrLf & _
               "→ Macro Settings → check 'Trust access to the VBA project object model'" & vbCrLf & vbCrLf & _
               "Then try again.", _
               vbCritical, "HMF Brief Approval — Update Failed"
        Exit Sub
    End If

    ' Create temp directory
    If Not CreateTempDirectory() Then
        MsgBox "Could not create temporary directory for update files." & vbCrLf & _
               "Path: " & TEMP_PATH, _
               vbCritical, "HMF Brief Approval — Update Failed"
        Exit Sub
    End If

    ' Define modules to update
    Dim arrModules(4) As String
    arrModules(0) = "modUtilities"
    arrModules(1) = "modQuickbase"
    arrModules(2) = "modApprovals"
    arrModules(3) = "modRoles"
    arrModules(4) = "modRibbon"

    ' Track results
    Dim arrResults(4) As String
    Dim intSuccess As Integer
    Dim intFailed As Integer
    intSuccess = 0
    intFailed = 0

    ' Process each module
    Dim i As Integer
    For i = 0 To MODULE_COUNT - 1
        Dim strModule   As String
        Dim strURL      As String
        Dim strFilePath As String
        Dim strContent  As String

        strModule   = arrModules(i)
        strURL      = GITHUB_RAW_BASE & strModule & ".bas"
        strFilePath = TEMP_PATH & strModule & ".bas"

        Debug.Print "Updating " & strModule & "..."

        ' Download the file
        strContent = DownloadFile(strURL)

        If strContent = "" Then
            arrResults(i) = strModule & " — FAILED (download error)"
            intFailed = intFailed + 1
            Debug.Print "  FAILED: Could not download " & strURL
        Else
            ' Write to temp file
            If WriteToFile(strFilePath, strContent) Then
                ' Remove existing module and reimport
                If ReimportModule(strModule, strFilePath) Then
                    arrResults(i) = strModule & " — OK"
                    intSuccess = intSuccess + 1
                    Debug.Print "  OK"
                Else
                    arrResults(i) = strModule & " — FAILED (import error)"
                    intFailed = intFailed + 1
                    Debug.Print "  FAILED: Could not reimport"
                End If
            Else
                arrResults(i) = strModule & " — FAILED (write error)"
                intFailed = intFailed + 1
                Debug.Print "  FAILED: Could not write temp file"
            End If
        End If

        ' Clean up temp file
        On Error Resume Next
        Kill strFilePath
        On Error GoTo 0
    Next i

    ' Clean up temp directory
    On Error Resume Next
    RmDir TEMP_PATH
    On Error GoTo 0

    ' Save document if any updates succeeded
    If intSuccess > 0 Then
        ActiveDocument.Save
    End If

    ' Show results summary
    Dim strSummary As String
    strSummary = "Update Complete" & vbCrLf & _
                 String(30, "-") & vbCrLf & vbCrLf

    For i = 0 To MODULE_COUNT - 1
        strSummary = strSummary & arrResults(i) & vbCrLf
    Next i

    strSummary = strSummary & vbCrLf & _
                 String(30, "-") & vbCrLf & _
                 "Succeeded: " & intSuccess & vbCrLf & _
                 "Failed:    " & intFailed

    If intFailed = 0 Then
        MsgBox strSummary, vbInformation, "HMF Brief Approval — Update Successful"
    Else
        MsgBox strSummary, vbExclamation, "HMF Brief Approval — Update Completed with Errors"
    End If

    ' Invalidate ribbon to reflect any changes
    InvalidateRibbon
End Sub

'================================================================
' DownloadFile
' Downloads the content of a URL as a string.
' Returns empty string on failure.
'================================================================
Private Function DownloadFile(strURL As String) As String
    Dim http As Object

    On Error GoTo ErrorHandler

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    With http
        .Open "GET", strURL, False
        .setRequestHeader "User-Agent", "HMF-BriefApproval-Word/1.0"
        .send

        If .Status = 200 Then
            DownloadFile = .responseText
        Else
            Debug.Print "  HTTP " & .Status & " for " & strURL
            DownloadFile = ""
        End If
    End With

    Set http = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "  Connection error: " & Err.Description
    DownloadFile = ""
    If Not http Is Nothing Then Set http = Nothing
End Function

'================================================================
' WriteToFile
' Writes a string to a local file path.
' Returns True on success.
'================================================================
Private Function WriteToFile(strPath As String, strContent As String) As Boolean
    Dim intFile As Integer

    On Error GoTo ErrorHandler

    intFile = FreeFile
    Open strPath For Output As #intFile
    Print #intFile, strContent
    Close #intFile

    WriteToFile = True
    Exit Function

ErrorHandler:
    Debug.Print "  File write error: " & Err.Description
    WriteToFile = False
    On Error GoTo 0
    If intFile > 0 Then Close #intFile
End Function

'================================================================
' ReimportModule
' Removes the existing module from the VBA project and
' imports the new version from the temp file path.
' Returns True on success.
'================================================================
Private Function ReimportModule(strModuleName As String, _
                                 strFilePath As String) As Boolean
    Dim vbProj      As Object
    Dim vbComp      As Object

    On Error GoTo ErrorHandler

    Set vbProj = ActiveDocument.VBProject

    ' Remove existing module if it exists
    On Error Resume Next
    Set vbComp = vbProj.VBComponents(strModuleName)
    On Error GoTo ErrorHandler

    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If

    ' Import the new module
    vbProj.VBComponents.Import strFilePath

    ReimportModule = True
    Exit Function

ErrorHandler:
    Debug.Print "  VBA import error: " & Err.Description
    ReimportModule = False
End Function

'================================================================
' CreateTempDirectory
' Creates the temporary directory for downloaded files.
' Returns True if directory exists or was created successfully.
'================================================================
Private Function CreateTempDirectory() As Boolean
    On Error GoTo ErrorHandler

    If Dir(TEMP_PATH, vbDirectory) = "" Then
        MkDir TEMP_PATH
    End If

    CreateTempDirectory = True
    Exit Function

ErrorHandler:
    CreateTempDirectory = False
End Function

'================================================================
' IsVBAAccessible
' Tests whether the VBA project object model is accessible.
' Returns False if Trust access is not enabled.
'================================================================
Private Function IsVBAAccessible() As Boolean
    On Error GoTo NotAccessible

    Dim vbProj As Object
    Set vbProj = ActiveDocument.VBProject

    ' If we can read the name, access is enabled
    Dim strTest As String
    strTest = vbProj.Name

    IsVBAAccessible = True
    Exit Function

NotAccessible:
    IsVBAAccessible = False
End Function

'================================================================
' CheckForUpdates
' Lightweight version — just checks if GitHub is reachable
' without downloading anything. Useful for connectivity testing.
' Run from Immediate Window: CheckForUpdates
'================================================================
Public Sub CheckForUpdates()
    Debug.Print "Checking GitHub connectivity..."

    Dim strTest As String
    strTest = DownloadFile(GITHUB_RAW_BASE & "modUtilities.bas")

    If strTest = "" Then
        Debug.Print "FAILED: Could not reach GitHub"
        Debug.Print "URL: " & GITHUB_RAW_BASE & "modUtilities.bas"
    Else
        Debug.Print "OK: GitHub is reachable"
        Debug.Print "modUtilities.bas — " & Len(strTest) & " bytes downloaded"
    End If
End Sub