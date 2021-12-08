VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInlineRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frm As frmInlineRegex
Attribute frm.VB_VarHelpID = -1

Private Type tInlineRegex
    inputRegexObj As VBScript_RegExp_55.RegExp
    outputRegexObj As VBScript_RegExp_55.RegExp
    outReplaceRegexObj As VBScript_RegExp_55.RegExp
    
    targetRange As Range
    previewDataBefore As Variant
    previewDataAfter As Variant
    
    cachedInput As Variant
    cachedOutput As Variant
    typeMask As Variant
    resultMask As Variant
    results As reResults
    
    matchPattern As String
    outputPattern As String
End Type

Private this As tInlineRegex

Private Sub Class_Initialize()
    Set this.inputRegexObj = New VBScript_RegExp_55.RegExp
    Set this.outputRegexObj = New VBScript_RegExp_55.RegExp
    Set this.outReplaceRegexObj = New VBScript_RegExp_55.RegExp
    Set frm = New frmInlineRegex
    
    InitialiseRegexObjects
End Sub

Private Sub InitialiseRegexObjects()
    With this.inputRegexObj
        .Global = True
        .MultiLine = False
        .IgnoreCase = True
    End With
    
    With this.outputRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
    
    With this.outReplaceRegexObj
        .Global = True
        .MultiLine = True
        .IgnoreCase = False
    End With
End Sub

Public Sub Go()
    If frm.Visible = True Then Exit Sub
      
    'Set this.targetRange = ThisWorkbook.Worksheets("Sheet1").listobjects(1).ListColumns(1).DataBodyRange
    'Set this.targetRange = ThisWorkbook.Worksheets("Sheet1").listobjects(1).DataBodyRange
    Set this.targetRange = Selection
    
    If this.targetRange.Cells.Count = 1 Then
        MsgBox "Cannot run on a single cell!", vbCritical + vbOKOnly, "Inline Regex"
        Exit Sub
    End If
    
    Set this.targetRange = Intersect(this.targetRange, Selection.Worksheet.UsedRange)
    
    If this.targetRange Is Nothing Then
        MsgBox "Selection is outside of UsedRange!", vbCritical + vbOKOnly, "Inline Regex"
        Exit Sub
    End If
    
    If this.targetRange.Cells.Count > 1000000 Then
        MsgBox "Choose a smaller selection!", vbCritical + vbOKOnly, "Inline Regex"
        Exit Sub
    End If
    
    If GetPreviewData Then
        frm.Show
    End If
End Sub

Private Function GetPreviewData() As Boolean
    Dim rng As Range
    Dim rowCount As Integer
    
    Set rng = this.targetRange
    rowCount = rng.Rows.Count
    If rowCount > 100 Then rowCount = 100
    Set rng = rng.Resize(rowCount, 1)
    this.previewDataBefore = ArrayFilterTextOnly(RangeToArray(rng))
    
    If IsEmpty(this.previewDataBefore) Then
        MsgBox "No text data found.", vbCritical + vbOKOnly, "Inline Regex"
        GetPreviewData = False
        Exit Function
    End If
    
    frm.LoadBeforeValues this.previewDataBefore
    GetPreviewData = True
End Function

Private Sub Class_Terminate()
    Set this.inputRegexObj = Nothing
    Set this.outputRegexObj = Nothing
    Set this.outReplaceRegexObj = Nothing
End Sub

Private Function Execute(ByVal inputString As String, ByVal pattern As String, ByVal outputPattern As String) As String
    Dim inputMatches As Object
    Dim replaceMatches As Object
    Dim replaceMatch As Object
    Dim replaceNumber As Integer
    
    'outputPattern = "$0"
    
    With this.inputRegexObj
        '.pattern = pattern
        'On Error GoTo RegexFailed
        Set inputMatches = .Execute(inputString)
        'On Error GoTo 0
    End With
    
    If inputMatches.Count = 0 Then
        Execute = False
        Set inputMatches = Nothing
        Exit Function
    End If
    
    ' TODO this can be cached
    this.outputRegexObj.pattern = "\$(\d+)"
    
    Set replaceMatches = this.outputRegexObj.Execute(outputPattern)
    For Each replaceMatch In replaceMatches
        replaceNumber = replaceMatch.SubMatches(0)
        this.outReplaceRegexObj.pattern = "\$" & replaceNumber

        If replaceNumber = 0 Then
            outputPattern = this.outReplaceRegexObj.Replace(outputPattern, inputMatches(0).Value)
        Else
            If replaceNumber > inputMatches(0).SubMatches.Count Then
                Execute = CVErr(xlErrValue)
                Set inputMatches = Nothing
                Exit Function
            Else
                outputPattern = this.outReplaceRegexObj.Replace(outputPattern, inputMatches(0).SubMatches(replaceNumber - 1))
            End If
        End If
    Next replaceMatch
    
    Execute = outputPattern
    Set inputMatches = Nothing
    
    Exit Function
RegexFailed:
    ' Invalid pattern e.g. ^.*{3}
    Execute = False
End Function

Private Function RangeToArray(rng As Range) As Variant
    Dim arr As Variant
    arr = ArrayTrim(ArrayDistinct(rng.Value), 10)
    RangeToArray = arr
End Function

Private Sub frm_DoRun(matchPattern As String, outputPattern As String)
    'this.matchPattern = matchPattern
    'this.outputPattern = outputPattern
    
    GetVarTypes
    CountVarTypes
    ExecuteRegex
    
    Call frm.UpdateResults(ResultsToArray(this.results))
End Sub

Private Function ResultsToArray(results As reResults) As Variant
    Dim arr As Variant
    ReDim arr(1 To 6, 1 To 3)
    
    arr(1, 1) = "Match"
    arr(1, 2) = Format(results.Match / results.Total, "0%")
    arr(1, 3) = results.Match
    arr(2, 1) = "Non-match"
    arr(2, 2) = Format(results.NonMatch / results.Total, "0%")
    arr(2, 3) = results.NonMatch
    arr(3, 1) = "Non-text"
    arr(3, 2) = Format(results.NonText / results.Total, "0%")
    arr(3, 3) = results.NonText
    arr(4, 1) = "Blanks"
    arr(4, 2) = Format(results.Blanks / results.Total, "0%")
    arr(4, 3) = results.Blanks
    arr(5, 1) = "Errors"
    arr(5, 2) = Format(results.Errors / results.Total, "0%")
    arr(5, 3) = results.Errors
    arr(6, 1) = "Total"
    arr(6, 2) = ""
    arr(6, 3) = results.Total
    
    ResultsToArray = arr
End Function

Private Sub GetVarTypes()
    this.cachedInput = this.targetRange.Value
    this.typeMask = this.targetRange.Value

    Dim r As Integer
    Dim c As Integer
    For r = 1 To this.targetRange.Rows.Count
        For c = 1 To this.targetRange.Columns.Count
            this.typeMask(r, c) = VarType(this.cachedInput(r, c))
        Next c
    Next r
End Sub

Private Sub CountVarTypes()
    Dim res As reResults
    Dim r As Integer
    Dim c As Integer
    For r = 1 To this.targetRange.Rows.Count
        For c = 1 To this.targetRange.Columns.Count
            Select Case (this.typeMask(r, c))
                Case vbEmpty
                    res.Blanks = res.Blanks + 1
                Case vbNull
                    res.Blanks = res.Blanks + 1
                Case vbError
                    res.Errors = res.Blanks + 1
                Case vbString
                    'res.Match = res.Match +1
                Case Else
                    res.NonText = res.NonText + 1
            End Select
        Next c
    Next r
    res.Total = this.targetRange.Cells.Count
    res.Match = 0
    res.NonMatch = 0
    this.results = res
End Sub

Private Sub ExecuteRegex()
    this.cachedOutput = this.cachedInput
    this.resultMask = this.cachedInput
    ReDim resultMask(1 To UBound(this.resultMask, 1), 1 To UBound(this.resultMask, 2))
    
    this.inputRegexObj.pattern = this.matchPattern
    
    Dim r As Integer
    Dim c As Integer
    For r = 1 To this.targetRange.Rows.Count
        For c = 1 To this.targetRange.Columns.Count
            this.resultMask(r, c) = 0
            If this.typeMask(r, c) = vbString Then
                this.cachedOutput(r, c) = Execute(CStr(this.cachedInput(r, c)), this.matchPattern, this.outputPattern)
            
                If this.cachedOutput(r, c) = False Then
                    this.results.NonMatch = this.results.NonMatch + 1
                    this.resultMask(r, c) = -1
                Else
                    this.results.Match = this.results.Match + 1
                    this.resultMask(r, c) = 1
                End If
            End If
        Next c
    Next r
End Sub

Private Sub ApplyFormatting(idx As Integer)
    Dim rng As Range
    Dim r As Integer
    Dim c As Integer
    
    For r = 1 To this.targetRange.Rows.Count
        For c = 1 To this.targetRange.Columns.Count
            If this.resultMask(r, c) = idx Then
                If rng Is Nothing Then
                    Set rng = this.targetRange.Cells(r, c)
                Else
                    Set rng = Union(rng, this.targetRange.Cells(r, c))
                End If
            End If
        Next c
    Next r
    
    If rng Is Nothing Then
        Exit Sub
    End If
    
    Select Case idx
        Case 1
            modFormatting.FormatMatch rng
        Case -1
            modFormatting.FormatNonMatch rng
            ' Replace FALSE of non-matches with blank string
            rng.Value = ""
    End Select
End Sub

Private Sub frm_DoSave(matchPattern As String, outputPattern As String)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
  
    this.targetRange.Value = this.cachedOutput
    
    If frm.DoHighlight Then
        Call FormatReset(this.targetRange)
        ApplyFormatting 1
        ApplyFormatting -1
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Data saved to sheet!", vbInformation + vbOKOnly
    
    frm.Hide
    Unload frm
End Sub

Private Sub frm_PreviewRequested(matchPattern As String, outputPattern As String)
    Dim i As Integer
    Dim s As String
    
    this.matchPattern = matchPattern
    this.outputPattern = outputPattern
    this.inputRegexObj.pattern = this.matchPattern
    
    this.previewDataAfter = this.previewDataBefore
    
    For i = LBound(this.previewDataAfter, 1) To UBound(this.previewDataAfter, 1)
        s = CStr(this.previewDataAfter(i, 1))
        this.previewDataAfter(i, 1) = Execute(s, matchPattern, outputPattern)
    Next i
    
    frm.LoadAfterValues this.previewDataAfter
End Sub
