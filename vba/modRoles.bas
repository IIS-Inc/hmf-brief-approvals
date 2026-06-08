'================================================================
' modRoles.bas
' HMF Brief Approval System
' Role management — passcode authentication, access control
'================================================================

Option Explicit

'----------------------------------------------------------------
' Passcodes — update these for production deployment
'----------------------------------------------------------------
Private Const CODE_ADMIN        As String = "Admin"
Private Const CODE_DIGITAL      As String = "Digital"
Private Const CODE_RESEARCH     As String = "Research"
Private Const CODE_EXECUTIVE    As String = "Executive"

' Document property name for cached role
Private Const PROP_USER_ROLE    As String = "HMF_UserRole"

' Role constants — public so other modules can reference them
Public Const ROLE_ADMIN         As String = "Admin"
Public Const ROLE_DIGITAL       As String = "Digital"
Public Const ROLE_RESEARCH      As String = "Research"
Public Const ROLE_EXECUTIVE     As String = "Executive"
Public Const ROLE_NONE          As String = ""

'================================================================
' InitializeUserRole
' Called from InitializeDocument in modUtilities on every open.
' Prompts for passcode, validates, stores role for session.
' Silent — no error messages on wrong code or cancel.
'================================================================
Public Sub InitializeUserRole()
    Dim strRole As String
    strRole = PromptForPasscode()

    ' Silently store whatever role was returned
    ' ROLE_NONE means no approval buttons will appear
    SetStatusProperty PROP_USER_ROLE, strRole
End Sub

'================================================================
' PromptForPasscode
' Displays an input box for the passcode.
' One attempt only — wrong code or cancel = no role assigned.
' No error messages — silent denial.
'================================================================
Private Function PromptForPasscode() As String
    Dim strInput As String

    strInput = InputBox("Enter your approval passcode to access the HMF Brief Approval System." & vbCrLf & vbCrLf & _
                        "Leave blank and click OK to open in read-only mode.", _
                        "HMF Brief Approval — Authentication", "")

    ' Exact match check — case sensitive, one attempt
    Select Case strInput
        Case CODE_ADMIN:        PromptForPasscode = ROLE_ADMIN
        Case CODE_DIGITAL:      PromptForPasscode = ROLE_DIGITAL
        Case CODE_RESEARCH:     PromptForPasscode = ROLE_RESEARCH
        Case CODE_EXECUTIVE:    PromptForPasscode = ROLE_EXECUTIVE
        Case Else:              PromptForPasscode = ROLE_NONE
    End Select
End Function

'================================================================
' GetCurrentUserRole
' Returns the cached role from document properties.
' Called by ribbon callbacks to show/hide buttons.
'================================================================
Public Function GetCurrentUserRole() As String
    GetCurrentUserRole = GetStatusProperty(PROP_USER_ROLE)
End Function

'================================================================
' IsButtonVisible
' Called by ribbon callbacks to determine button visibility.
' Admin sees all buttons, others see only their department.
'================================================================
Public Function IsButtonVisible(strDept As String) As Boolean
    Dim strRole As String
    strRole = GetCurrentUserRole()

    Select Case strRole
        Case ROLE_ADMIN
            IsButtonVisible = True
        Case strDept
            IsButtonVisible = True
        Case Else
            IsButtonVisible = False
    End Select
End Function

'================================================================
' CanUserAct
' Combines role check AND stage gate.
' Returns True only if user has the right role AND
' the workflow is at the correct stage.
'================================================================
Public Function CanUserAct(strDept As String) As Boolean
    Dim strRole As String
    strRole = GetCurrentUserRole()

    ' No role — no access
    If strRole = ROLE_NONE Then
        CanUserAct = False
        Exit Function
    End If

    ' Admin can act if stage allows
    If strRole = ROLE_ADMIN Then
        CanUserAct = CanDeptAct(strDept)
        Exit Function
    End If

    ' Department role must match requested department
    If strRole <> strDept Then
        CanUserAct = False
        Exit Function
    End If

    ' Role matches — check stage gate
    CanUserAct = CanDeptAct(strDept)
End Function

'================================================================
' ResetRole
' Clears the cached role — useful for testing.
' Run from Immediate Window: ResetRole
'================================================================
Public Sub ResetRole()
    SetStatusProperty PROP_USER_ROLE, ""
    MsgBox "Role has been cleared." & vbCrLf & _
           "You will be prompted on next document open.", _
           vbInformation, "HMF Brief Approval"
End Sub

'================================================================
' TestPasscode
' Dev utility — tests passcode flow without reopening document.
' Run from Immediate Window: TestPasscode
'================================================================
Public Sub TestPasscode()
    Debug.Print "Testing passcode authentication..."
    InitializeUserRole
    Debug.Print "  Role assigned: " & GetCurrentUserRole()
End Sub