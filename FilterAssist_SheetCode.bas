' ===========================================================================
' Filter assist column: auto-fill merged cells
'
' Paste this into the WORKSHEET's own code module (right-click the sheet tab
' -> View Code), NOT into a normal module. Save the workbook as .xlsm.
'
' When you type into (or clear) a merged cell in ASSIST_COL, the value is
' copied into the hidden cells underneath the merge, and the merge is kept.
' Excel's filter then sees the value on every row of the merged block.
' ===========================================================================

Private Const ASSIST_COL As String = "W"   ' <-- change to your filter assist column

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim hit As Range
    Dim c As Range
    Dim done As String
    Dim prevSel As Range

    Set hit = Intersect(Target, Me.Columns(ASSIST_COL), Me.UsedRange)
    If hit Is Nothing Then Exit Sub

    If TypeName(Selection) = "Range" Then Set prevSel = Selection

    On Error GoTo CleanUp
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each c In hit.Cells
        If c.MergeCells Then
            ' Process each merged block once, even if several of its cells changed
            If InStr(done, "|" & c.MergeArea.Address & "|") = 0 Then
                done = done & "|" & c.MergeArea.Address & "|"
                FillMergedArea c.MergeArea
            End If
        End If
    Next c

CleanUp:
    Application.CutCopyMode = False
    If Not prevSel Is Nothing Then prevSel.Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then MsgBox "Filter assist fill failed: " & Err.Description, vbExclamation
End Sub

' Copy the top-left value into every cell of the merged block, keeping the merge.
Private Sub FillMergedArea(area As Range)
    Dim scratch As Worksheet
    Dim v As Variant

    v = area.Cells(1, 1).Value
    Set scratch = GetScratchSheet()
    scratch.Cells.Clear

    ' Save the merged layout and formatting, unmerge, fill, then paste the
    ' formats back (like Format Painter) to re-merge without losing the values
    area.Copy scratch.Range("A1")
    area.UnMerge
    area.Value = v
    scratch.Range("A1").Resize(area.Rows.Count, area.Columns.Count).Copy
    area.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

' A very hidden helper sheet used as a scratch area. Created on first use.
Private Function GetScratchSheet() As Worksheet
    On Error Resume Next
    Set GetScratchSheet = Me.Parent.Worksheets("_MergeScratch")
    On Error GoTo 0

    If GetScratchSheet Is Nothing Then
        Set GetScratchSheet = Me.Parent.Worksheets.Add(After:=Me.Parent.Sheets(Me.Parent.Sheets.Count))
        GetScratchSheet.Name = "_MergeScratch"
        GetScratchSheet.Visible = xlSheetVeryHidden
        Me.Activate
    End If
End Function
