'================================================================
' CleanTemplate
' Removes all HMF custom document properties.
' Uses Item().Delete syntax which is reliable across Word versions.
' Forces document dirty flag before save to ensure changes persist.
' Run from Immediate Window or Admin ribbon button: CleanTemplate
'================================================================
Public Sub CleanTemplate()

    If MsgBox("This will remove all HMF stored properties from this document." & vbCrLf & vbCrLf & _
              "Use this before distributing the master template." & vbCrLf & vbCrLf & _
              "Are you sure?", _
              vbQuestion + vbYesNo, "HMF Brief Approval — Clean Template") = vbNo Then
        Exit Sub
    End If

    Dim intRemoved As Integer
    intRemoved = 0

    ' Delete each property individually using Item().Delete
    ' Cannot loop and delete simultaneously in VBA
    ' Must use Item() syntax — direct .Delete is unreliable
    On Error Resume Next

    ActiveDocument.CustomDocumentProperties.Item("HMF_RID").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    ActiveDocument.CustomDocumentProperties.Item("HMF_BriefStatus").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    ActiveDocument.CustomDocumentProperties.Item("HMF_DigitalStatus").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    ActiveDocument.CustomDocumentProperties.Item("HMF_ResearchStatus").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    ActiveDocument.CustomDocumentProperties.Item("HMF_ExecutiveStatus").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    ActiveDocument.CustomDocumentProperties.Item("HMF_UserRole").Delete
    If Err.Number = 0 Then intRemoved = intRemoved + 1
    Err.Clear

    On Error GoTo 0

    ' Force document dirty flag — property deletions don't
    ' automatically trigger Word's save mechanism
    ActiveDocument.Saved = False

    ' Invalidate ribbon to reflect cleared state immediately
    InvalidateRibbon

    ' Save the clean document
    ActiveDocument.Save

    MsgBox "Template cleaned successfully." & vbCrLf & vbCrLf & _
           intRemoved & " properties removed." & vbCrLf & _
           "Document saved.", _
           vbInformation, "HMF Brief Approval — Clean Complete"

    ' Prompt for new RID immediately after clean
    Dim lngRID As Long
    lngRID = GetRID()

    ' Invalidate again to show new RID in ribbon if entered
    If lngRID > 0 Then
        InvalidateRibbon
    End If

End Sub