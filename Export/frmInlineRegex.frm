VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInlineRegex 
   Caption         =   "Inline Regex"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "frmInlineRegex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInlineRegex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DoRun(matchPattern As String, outputPattern As String)
Public Event DoSave(matchPattern As String, outputPattern As String)
Public Event PreviewRequested(matchPattern As String, outputPattern As String)
Public Event Cancellled()

Private Type tInlineRegex
    rng As Range
End Type
    
Private this As tInlineRegex

Public Function DoHighlight() As Boolean
    DoHighlight = Me.chkHighlight.Value
End Function

Public Function LoadBeforeValues(arr As Variant)
    LoadArrayToListView arr, Me.lvPreview, 1
End Function

Public Function LoadAfterValues(arr As Variant)
    LoadArrayToListView arr, Me.lvPreview, 2
End Function

Private Sub cmbRun_Click()
    'Me.Hide
    RaiseEvent DoRun(Me.txtExpression, Me.txtResults)
    Me.cmbRun.Enabled = False
    Me.cmbSave.Enabled = True
    Me.cmbSave.Default = True
    Me.cmbSave.SetFocus
    'Unload Me
End Sub

Private Sub cmbPreview_Click()
    Me.cmbPreview.Enabled = False
    RaiseEvent PreviewRequested(Me.txtExpression, Me.txtResults)
    Me.cmbRun.Enabled = True
    Me.cmbRun.Default = True
End Sub

Private Sub cmbCancel_Click()
    Me.Hide
    RaiseEvent Cancellled
    Unload Me
End Sub

Private Sub cmbSave_Click()
    RaiseEvent DoSave(Me.txtExpression, Me.txtResults)
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub txtExpression_Change()
    CheckButtons
    Me.cmbRun.Enabled = False
    Me.cmbSave.Enabled = False
    Me.cmbPreview.Enabled = True
    Me.cmbPreview.Default = True
End Sub

Private Sub CheckButtons()
    Me.cmbPreview.Enabled = Len(Me.txtExpression) > 0
    
    If ((Me.cmbRun.Enabled = False) And Me.cmbPreview.Enabled = True) Then
        Me.cmbRun.Default = False
        Me.cmbPreview.Default = True
    End If
End Sub

Private Sub txtResults_Change()
    CheckButtons
    Me.cmbRun.Enabled = False
    Me.cmbSave.Enabled = False
    Me.cmbPreview.Enabled = True
    Me.cmbPreview.Default = True
End Sub

Private Sub UserForm_Initialize()
    Me.txtExpression = "^.*$"
    Me.txtResults = "$0"
    SetupImages
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent Cancellled
End Sub

Private Sub LoadArrayToListView(arr As Variant, lv As ListView, idx As Integer)
    If idx = 1 Then
        lv.Gridlines = True
        lv.View = lvwReport
        lv.ListItems.Clear
        lv.ColumnHeaders.Clear
    
        lv.ColumnHeaders.Add , , "Before"
        lv.ColumnHeaders.Add , , "After"
    End If
    
    Dim i As Integer
    For i = 1 To ArrayLength(arr)
        If idx = 1 Then
            lv.ListItems.Add , , CStr(arr(i, 1))
        Else
            lv.ListItems(i).ListSubItems.Clear
            lv.ListItems(i).ListSubItems.Add , , arr(i, 1)
            If arr(i, 1) = False Then
                lv.ListItems(i).SmallIcon = "Non-match"
            Else
                lv.ListItems(i).SmallIcon = "Match"
            End If
        End If
    Next i
End Sub

Private Sub LoadResultsToListView(arr As Variant, lv As ListView)
    lv.Gridlines = True
    lv.View = lvwReport

    lv.ListItems.Clear
    lv.ColumnHeaders.Clear

    lv.ColumnHeaders.Add , , "Description"
    lv.ColumnHeaders.Add , , "Percentage"
    lv.ColumnHeaders.Add , , "Value"
    lv.ColumnHeaders(1).Width = 75
    lv.ColumnHeaders(2).Width = 50
    lv.ColumnHeaders(3).Width = 50
    
    Dim i As Integer
    For i = 1 To ArrayLength(arr)
        lv.ListItems.Add , , CStr(arr(i, 1))
        lv.ListItems(i).ListSubItems.Add , , arr(i, 2)
        lv.ListItems(i).ListSubItems.Add , , arr(i, 3)
        lv.ListItems(i).SmallIcon = arr(i, 1)
    Next i
End Sub

Private Sub SetupImages()
    Dim il As ImageList
    Set il = Me.ilImageList
    If il.ListImages.Count > 0 Then
        Exit Sub
    End If
    
    il.ImageWidth = 16
    il.ImageHeight = 16
    il.ListImages.Add 1, "Match", Me.imgPass.Picture
    il.ListImages.Add 2, "Non-match", Me.imgFail.Picture
    il.ListImages.Add 2, "Non-text", Me.imgNonText.Picture
    il.ListImages.Add 3, "Blanks", Me.imgBlank.Picture
    il.ListImages.Add 4, "Errors", Me.imgError.Picture
    il.ListImages.Add 5, "Total", Me.imgTotal.Picture
    
    Me.lvPreview.SmallIcons = il
    Me.lvResults.SmallIcons = il
End Sub

Public Sub UpdateResults(results As Variant)
    Call LoadResultsToListView(results, Me.lvResults)
End Sub
