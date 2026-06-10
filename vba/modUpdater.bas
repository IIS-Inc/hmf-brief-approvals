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
Private Const TOKEN_PLACEHOLDER As String = "REPLACE_WITH_TOKEN"

' Modules managed by the updater
' modUpdater is excluded — cannot update itself
Private Const MODULE_COUNT      As Integer = 5

'================================================================
' UpdateFromGitHub
' Master public function — called by Admin from Immediate Window.
' Downloads and reimports all managed modules from GitHub.
' Schedules PostUpdateSave to run after VBA settles.
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
                 "NOTE: modUpdater cannot update itself." & vbCrLf & _
                 "NOTE: Token will need to be reinstalled after update." & vbCrLf & vbCrLf & _
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

        strContent = DownloadFile(strURL)

        If strContent = "" Then
            arrResults(i) = strModule & " — FAILED (download error)"
            intFailed = intFailed + 1
            Debug.Print "  FAILED: Could not download " & strURL
        Else
            If WriteToFile(strFilePath, strContent) Then
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

        On Error Resume Next
        Kill strFilePath
        On Error GoTo 0
    Next i

    On Error Resume Next
    RmDir TEMP_PATH
    On Error GoTo 0

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
        MsgBox strSummary & vbCrLf & vbCrLf & _
               "NEXT STEP: Run InstallToken to install the API token.", _
               vbInformation, "HMF Brief Approval — Update Successful"
    Else
        MsgBox strSummary, vbExclamation, "HMF Brief Approval — Update Completed with Errors"
    End If

    ' Schedule save after VBA settles — avoids crash on immediate save
    If intSuccess > 0 Then
        Application.OnTime Now + TimeValue("0:00:05"), "PostUpdateSave"
    End If

    InvalidateRibbon
End Sub

'================================================================
' PostUpdateSave
' Scheduled by UpdateFromGitHub via Application.OnTime.
' Runs 5 seconds after update completes giving VBA time to settle.
' Public so Application.OnTime can call it by name.
'================================================================
Public Sub PostUpdateSave()
    On Error GoTo ErrorHandler

    ActiveDocument.Save
    Debug.Print "Document saved successfully after update."

    ' Prompt Admin to install token
    If MsgBox("Document saved successfully." & vbCrLf & vbCrLf & _
              "Would you like to install the API token now?", _
              vbQuestion + vbYesNo, "HMF Brief Approval — Install Token") = vbYes Then
        InstallToken
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Document could not be saved automatically." & vbCrLf & _
           "Please save manually with Ctrl+S.", _
           vbExclamation, "HMF Brief Approval — Save Required"
End Sub

'================================================================
' InstallToken
' Admin only — prompts for the Quickbase API token and
' writes it directly into modQuickbase using in-place
' code replacement. Token never stored in GitHub.
' Run from Immediate Window: InstallToken
'================================================================
Public Sub InstallToken()

    ' Enforce Admin role
    If GetCurrentUserRole() <> ROLE_ADMIN Then
        MsgBox "Token installation is available to Admin users only.", _
               vbExclamation, "HMF Brief Approval — Access Denied"
        Exit Sub
    End If

    ' Verify VBA project object model is accessible
    If Not IsVBAAccessible() Then
        MsgBox "The VBA project object model is not accessible." & vbCrLf & _
               "Enable it in Trust Center Settings first.", _
               vbCritical, "HMF Brief Approval — Token Install Failed"
        Exit Sub
    End If

    ' Prompt for token
    Dim strToken As String
    strToken = InputBox("Enter the Quickbase API User Token:" & vbCrLf & vbCrLf & _
                        "This token will be installed directly into the document." & vbCrLf & _
                        "It will not be saved to GitHub.", _
                        "HMF Brief Approval — Install API Token", "")

    ' User cancelled
    If strToken = "" Then
        MsgBox "Token installation cancelled.", _
               vbInformation, "HMF Brief Approval"
        Exit Sub
    End If

    strToken = Trim(strToken)

    ' Basic validation — QB tokens follow a pattern
    If Len(strToken) < 20 Then
        MsgBox "That doesn't look like a valid Quickbase token." & vbCrLf & _
               "Tokens are typically longer than 20 characters.", _
               vbExclamation, "HMF Brief Approval — Invalid Token"
        Exit Sub
    End If

    ' Write token into modQuickbase using in-place replacement
    If ReplaceTokenInModule(strToken) Then
        ' Save document to persist the token
        ActiveDocument.Save

        MsgBox "Token installed and document saved successfully!" & vbCrLf & vbCrLf & _
               "Run TestConnection in the Immediate Window to verify.", _
               vbInformation, "HMF Brief Approval — Token Installed"

        Debug.Print "Token installed successfully."
    Else
        MsgBox "Token installation failed." & vbCrLf & _
               "Please check the VBA editor and try again.", _
               vbCritical, "HMF Brief Approval — Token Install Failed"
    End If
End Sub

'================================================================
' ReplaceTokenInModule
' Finds the token placeholder line in modQuickbase and
' replaces it with the actual token value.
' Returns True on success.
'================================================================
Private Function ReplaceTokenInModule(strToken As String) As Boolean
    Dim vbProj      As Object
    Dim vbComp      As Object
    Dim codeMod     As Object
    Dim i           As Long
    Dim strLine     As String

    On Error GoTo ErrorHandler

    Set vbProj = ActiveDocument.VBProject
    Set vbComp = vbProj.VBComponents("modQuickbase")
    Set codeMod = vbComp.CodeModule

    ' Search for the placeholder line
    For i = 1 To codeMod.CountOfLines
        strLine = codeMod.Lines(i, 1)

        If InStr(strLine, TOKEN_PLACEHOLDER) > 0 Then
            ' Replace just this line with the real token
            codeMod.ReplaceLine i, _
                "Private Const QB_USER_TOKEN     As String = """ & strToken & """"
            ReplaceTokenInModule = True
            Debug.Print "  Token installed at line " & i
            Exit Function
        End If
    Next i

    ' Placeholder not found — token may already be installed
    ' Check if a non-placeholder token exists
    Debug.Print "  Placeholder not found — token may already be installed"
    ReplaceTokenInModule = True
    Exit Function

ErrorHandler:
    Debug.Print "  Token install error: " & Err.Description
    ReplaceTokenInModule = False
End Function

'================================================================
' DownloadFile
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
'================================================================
Private Function ReimportModule(strModuleName As String, _
                                 strFilePath As String) As Boolean
    Dim vbProj      As Object
    Dim vbComp      As Object
    Dim vbNewComp   As Object
    Dim codeMod     As Object

    On Error GoTo ErrorHandler

    Set vbProj = ActiveDocument.VBProject

    On Error Resume Next
    Set vbComp = vbProj.VBComponents(strModuleName)
    On Error GoTo ErrorHandler

    If Not vbComp Is Nothing Then
        On Error GoTo TryRemove
        Set codeMod = vbComp.CodeModule

        If codeMod.CountOfLines > 0 Then
            codeMod.DeleteLines 1, codeMod.CountOfLines
        End If

        Dim intFile     As Integer
        Dim strLine     As String
        Dim strAll      As String
        intFile = FreeFile
        Open strFilePath For Input As #intFile
        Do While Not EOF(intFile)
            Line Input #intFile, strLine
            strAll = strAll & strLine & vbCrLf
        Loop
        Close #intFile

        codeMod.InsertLines 1, strAll
        vbComp.Name = strModuleName
        ReimportModule = True
        Exit Function
    End If

TryRemove:
    On Error GoTo ErrorHandler

    On Error Resume Next
    If Not vbComp Is Nothing Then
        vbProj.VBComponents.Remove vbComp
        Set vbComp = Nothing
    End If
    On Error GoTo ErrorHandler

    Set vbNewComp = vbProj.VBComponents.Import(strFilePath)

    If Not vbNewComp Is Nothing Then
        vbNewComp.Name = strModuleName
    End If

    ReimportModule = True
    Exit Function

ErrorHandler:
    Debug.Print "  VBA import error: " & Err.Description
    ReimportModule = False
    On Error GoTo 0
    If intFile > 0 Then Close #intFile
End Function

'================================================================
' CreateTempDirectory
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
'================================================================
Private Function IsVBAAccessible() As Boolean
    On Error GoTo NotAccessible

    Dim vbProj As Object
    Set vbProj = ActiveDocument.VBProject

    Dim strTest As String
    strTest = vbProj.Name

    IsVBAAccessible = True
    Exit Function

NotAccessible:
    IsVBAAccessible = False
End Function

'================================================================
' CheckForUpdates
' Run from Immediate Window: CheckForUpdates
'================================================================
Public Sub CheckForUpdates()
    Debug.Print "Checking GitHub connectivity..."

    Dim strTest As String
    strTest = DownloadFile(GITHUB_RAW_BASE & "modUtilities.bas")

    If strTest = "" Then
        Debug.Print "FAILED: Could not reach GitHub"
    Else
        Debug.Print "OK: GitHub is reachable"
        Debug.Print "modUtilities.bas — " & Len(strTest) & " bytes downloaded"
    End If
End Sub