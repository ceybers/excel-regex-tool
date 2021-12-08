Attribute VB_Name = "modMain"
' Inline Regular Expression (RegEx) tool
' Craig Eybers
' 20 September 2021
'
' Applies a RegEx pattern to a selection of cells in Excel.
' * Preview table to check result of pattern
' * Apply pattern to values in-memory instead of directly to sheet
' * Review results (% match vs non-match, # of blanks, errors, etc.)
' * Save results back to spreadsheet

Option Explicit

Public Type reResults
    Match As Integer
    NonMatch As Integer
    NonText As Integer
    Blanks As Integer
    Errors As Integer
    Total As Integer
End Type

Public Sub InlineRegex()
    With New clsInlineRegex
        .Go
    End With
End Sub
