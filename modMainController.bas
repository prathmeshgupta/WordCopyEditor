Attribute VB_Name = "modMainController"
Option Explicit

' Constants for UserForm modes
Public Const MODE_AUTO As Integer = 1
Public Const MODE_APPROVAL As Integer = 2

' Main procedure for Automatic Edit Mode
Sub RunAutoEditMode()
    Dim startTime As Double
    startTime = Timer
    
    ' Ensure Track Changes is on and set username
    Call modTrackChanges.SetTrackChangesAndUsername
    
    ' Run sequential edits
    Call modBasicEdits.RunBasicEdits(MODE_AUTO)
    
    ' Run citation and abbreviation checks (more complex, run last)
    Call modAdvancedChecks.RunAdvancedChecks(MODE_AUTO)
    
    Dim endTime As Double
    endTime = Timer
    MsgBox "Automatic Edit Mode completed in " & Format(endTime - startTime, "0.00") & " seconds.", vbInformation, "Edit Complete"
End Sub

' Main procedure for Edit with Approval Mode
Sub RunApprovalMode()
    Dim startTime As Double
    startTime = Timer
    
    ' Ensure Track Changes is on and set username
    Call modTrackChanges.SetTrackChangesAndUsername
    
    ' Run sequential edits with user approval
    Call modBasicEdits.RunBasicEdits(MODE_APPROVAL)
    
    ' Run citation and abbreviation checks with user approval
    Call modAdvancedChecks.RunAdvancedChecks(MODE_APPROVAL)
    
    Dim endTime As Double
    endTime = Timer
    MsgBox "Approval Mode completed in " & Format(endTime - startTime, "0.00") & " seconds.", vbInformation, "Edit Complete"
End Sub

' Procedure to set default username (can be called manually if needed)
Sub SetDefaultUsername()
    Call modTrackChanges.SetTrackChangesAndUsername
End Sub

' Placeholder for future GenAI integration
Sub LaunchSelfImprovementModule()
    MsgBox "Self-Improvement/Learning Module Placeholder." & vbCrLf & _
           "This would integrate with a GenAI model for advanced suggestions and learning." & vbCrLf & _
           "Implementation would require API keys and external model access.", vbInformation, "Future Feature"
End Sub
