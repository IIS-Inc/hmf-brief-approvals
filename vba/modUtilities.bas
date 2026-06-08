'================================================================
' modUtilities.bas
' HMF Brief Approval System
' Core utilities: RID harvesting, validation, document properties
'================================================================

Option Explicit

'----------------------------------------------------------------
' Constants
'----------------------------------------------------------------
Private Const PROP_RID          As String = "HMF_RID"
Private Const PROP_STATUS       As String = "HMF_BriefStatus"
Private Const PROP_DIG_STATUS   As String = "HMF_DigitalStatus"
Private Const PROP_RES_STATUS   As String = "HMF_ResearchStatus"
Private Const PROP_EXEC_STATUS  As String = "HMF_ExecutiveStatus"

'================================================================
' GetRID
' Master function — returns the Record ID as a Long.
' Check order:
'   1. Custom Document Property (fastest, most reliable)
'   2. Prompt user to enter manually
' Returns 0 on failure/cancel.
'================================================================
Public Function GetRID() As Long
    Dim lngRID As Long

    ' --- Step 1: Try document property first ---
    lngRID = GetRIDFromProperty()
    If lngRID > 0 Then
        GetRID = lngRID
        Exit Function
    End If

    ' --- Step 2: Prompt user ---
    lngRID = PromptForRID()
    GetRID = lngRID
End Function

'================================================================
' GetRIDFromProperty
' Reads HMF_RID from custom document properties.
' Returns 0 if not found or invalid.
'================================================================
Private Function GetRIDFromProperty() As Long
    Dim prop As Object

    On Error Resume Next
    Set prop = Nothing
    Set prop = ActiveDocument.CustomDocumentProperties(PROP_RID)
    On Error GoTo 0

    If prop Is Nothing Then
        GetRIDFromProperty = 0
        Exit Function
    End If

    Dim val As String
    val = Trim(CStr(prop.Value))

    If IsNumeric(val) And CLng(val) > 0 Then
        GetRIDFromProperty = CLng(val)
    Else
        GetRIDFromProperty = 0
    End If
End Function

'================================================================
' PromptForRID
' Shows an input box asking for the Quickbase Record ID.
' Validates numeric input. Loops until valid or cancelled.
' On success, saves to document property.
' Returns 0 if user cancels.
'================================================================
Private Function PromptForRID() As Long
    Dim strInput    As String
    Dim lngRID      As Long
    Dim strPrompt   As String
    Dim blnValid    As Boolean

    strPrompt = "No HMF Record ID was found for this document." & vbCrLf & vbCrLf & _
                "Please enter the Quickbase Record ID (numeric only):" & vbCrLf & vbCrLf & _
                "You can find this in Quickbase under the Brief record URL" & vbCrLf & _
                "or in the HMF ID field on the Brief form." & vbCrLf & vbCrLf & _
                "Click Cancel to open without an ID (Admin update mode)."

    blnValid = False

    Do While Not blnValid
        strInput = InputBox(strPrompt, "HMF Brief Approval — Record ID Required", "")

        ' User cancelled — allowed, Admin may need to update without RID
        If strInput = "" Then
            PromptForRID = 0
            Exit Function
        End If

        strInput = Trim(strInput)

        ' Validate: must be numeric and positive
        If IsNumeric(strInput) Then
            lngRID = CLng(strInput)
            If lngRID > 0 Then
                blnValid = True
            Else
                MsgBox "The Record ID must be a positive number. Please try again.", _
                       vbExclamation, "Invalid Record ID"
            End If
        Else
            MsgBox "'" & strInput & "' is not a valid Record ID." & vbCrLf & _
                   "Please enter numbers only (e.g. 577).", _
                   vbExclamation, "Invalid Record ID"
        End If
    Loop

    ' Save to document property for future use
    SaveRIDToProperty lngRID
    PromptForRID = lngRID
End Function

'================================================================
' SaveRIDToProperty
' Writes the RID to a custom document property.
' Creates the property if it doesn't exist; updates if it does.
'================================================================
Public Sub SaveRIDToProperty(lngRID As Long)
    Dim prop As Object

    On Error Resume Next
    Set prop = ActiveDocument.CustomDocumentProperties(PROP_RID)
    On Error GoTo 0

    If prop Is Nothing Then
        ActiveDocument.CustomDocumentProperties.Add _
            Name:=PROP_RID, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeNumber, _
            Value:=lngRID
    Else
        prop.Value = lngRID
    End If

    ActiveDocument.Save
End Sub

'================================================================
' GetStatusProperty / SetStatusProperty
' Read and write status document properties.
'================================================================
Public Function GetStatusProperty(strPropName As String) As String
    Dim prop As Object

    On Error Resume Next
    Set prop = ActiveDocument.CustomDocumentProperties(strPropName)
    On Error GoTo 0

    If prop Is Nothing Then
        GetStatusProperty = ""
    Else
        GetStatusProperty = CStr(prop.Value)
    End If
End Function

Public Sub SetStatusProperty(strPropName As String, strValue As String)
    Dim prop As Object

    On Error Resume Next
    Set prop = ActiveDocument.CustomDocumentProperties(strPropName)
    On Error GoTo 0

    If prop Is Nothing Then
        ActiveDocument.CustomDocumentProperties.Add _
            Name:=strPropName, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeString, _
            Value:=strValue
    Else
        prop.Value = strValue
    End If
End Sub

'================================================================
' GetCurrentBriefStatus / SetCurrentBriefStatus
' Convenience wrappers for Brief Approval Status (FID 8)
'================================================================
Public Function GetCurrentBriefStatus() As String
    GetCurrentBriefStatus = GetStatusProperty(PROP_STATUS)
End Function

Public Sub SetCurrentBriefStatus(strStatus As String)
    SetStatusProperty PROP_STATUS, strStatus
End Sub

'================================================================
