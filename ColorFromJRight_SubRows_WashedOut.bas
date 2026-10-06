
' Sheet layout (three-level hierarchy):
'   A:D  - Level 1. One row here can own several EFG rows below it.
'   E:G  - Level 2. One row here can own several "H onwards" rows below it.
'   H+   - Level 3. Column K holds the status. A row with K filled is a "main row";
'          rows with K blank are "sub rows" that inherit the status above them.
'
' Colouring:
'   - Main rows: J to last column, coloured by their own K status.
'   - Sub rows:  V to last column, coloured by the carried-over status.
'   - Fills from H onwards are extra light; E:G uses the stronger washed-out fills.
'   - E:G block: green if any K row in the block is Accepted,
'                otherwise yellow if any is Under Review, otherwise cleared.
'                Text in E:G (and everything left of it) stays black.

Private Const FIRST_DATA_ROW As Long = 2   ' Row 1 is the header row
Private Const MIN_LAST_COL As Long = 22    ' Always format at least to column V

Sub ColorFromJRight_SubRows_WashedOut()
    Dim ws As Worksheet
    Dim lastRowTotal As Long
    Dim lastCol As Long
    Dim i As Long
    Dim statusValue As String
    Dim formatRange As Range
    Dim calcMode As XlCalculation
    Dim blockStart As Long
    Dim hasAccepted As Boolean
    Dim hasUnderReview As Boolean

    ' Set to the active worksheet
    Set ws = ActiveSheet

    ' Find the absolute last row and column in the entire sheet
    ' (Necessary because sub-rows in Column V extend further down than Column K)
    On Error Resume Next
    lastRowTotal = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0

    ' If there is no data below the header, exit
    If lastRowTotal < FIRST_DATA_ROW Then Exit Sub

    If lastCol < MIN_LAST_COL Then lastCol = MIN_LAST_COL

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    statusValue = ""
    blockStart = 0

    For i = FIRST_DATA_ROW To lastRowTotal

        ' A value anywhere in A:D or E:G starts a new EFG block.
        ' Colour the block that just ended, then start tracking the new one.
        If StartsNewGroup(ws.Range(ws.Cells(i, "A"), ws.Cells(i, "G"))) Then
            If blockStart > 0 Then
                ColourEFGBlock ws, blockStart, i - 1, hasAccepted, hasUnderReview
            End If
            blockStart = i
            hasAccepted = False
            hasUnderReview = False
            statusValue = ""   ' Don't carry a status across into a new block
        End If

        ' If Column K has a value, it's a "Main Row". Update the status and format from J to the end.
        If RangeHasValue(ws.Cells(i, "K")) Then
            statusValue = LCase$(Trim$(CStr(ws.Cells(i, "K").Value)))
            Set formatRange = ws.Range(ws.Cells(i, "J"), ws.Cells(i, lastCol))

            If statusValue = "accepted" Then hasAccepted = True
            If statusValue = "under review" Then hasUnderReview = True

        ' If Column K is blank, it's a "Sub Row". Keep the previous status but format ONLY from V to the end.
        Else
            Set formatRange = ws.Range(ws.Cells(i, "V"), ws.Cells(i, lastCol))
        End If

        ApplyStatusColours formatRange, statusValue, True
    Next i

    ' Colour the final EFG block
    If blockStart > 0 Then
        ColourEFGBlock ws, blockStart, lastRowTotal, hasAccepted, hasUnderReview
    End If

CleanUp:
    Application.Calculation = calcMode
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Colouring stopped at row " & i & ": " & Err.Description, vbExclamation
    End If
End Sub

' Colour E:G for every row in the block based on the K statuses found within it.
Private Sub ColourEFGBlock(ws As Worksheet, firstRow As Long, lastRow As Long, _
                           hasAccepted As Boolean, hasUnderReview As Boolean)
    Dim efgRange As Range
    Set efgRange = ws.Range(ws.Cells(firstRow, "E"), ws.Cells(lastRow, "G"))

    If hasAccepted Then
        ApplyStatusColours efgRange, "accepted"
    ElseIf hasUnderReview Then
        ApplyStatusColours efgRange, "under review"
    Else
        ApplyStatusColours efgRange, ""
    End If

    ' Keep text black in columns G and prior; only the fill changes
    efgRange.Font.Color = RGB(0, 0, 0)
End Sub

' Apply washed-out fills and fonts for a status (expects lower-case, trimmed text).
' extraLight = True uses even paler fills (used for columns H onwards).
Private Sub ApplyStatusColours(target As Range, statusValue As String, _
                               Optional extraLight As Boolean = False)
    Select Case statusValue
        Case "accepted"
            target.Interior.Color = IIf(extraLight, RGB(243, 250, 243), RGB(230, 245, 230))
            target.Font.Color = RGB(130, 190, 130)

        Case "under review"
            target.Interior.Color = IIf(extraLight, RGB(255, 252, 238), RGB(255, 248, 220))
            target.Font.Color = RGB(210, 170, 90)

        Case "cancelled"
            target.Interior.Color = IIf(extraLight, RGB(249, 249, 249), RGB(242, 242, 242))
            target.Font.Color = RGB(160, 160, 160)

        Case "rejected"
            target.Interior.Color = IIf(extraLight, RGB(255, 243, 243), RGB(255, 230, 230))
            target.Font.Color = RGB(220, 130, 130)

        Case Else
            target.Interior.ColorIndex = xlNone
            target.Font.ColorIndex = xlAutomatic
    End Select
End Sub

' True if a row starts a new group: a cell in the range has a value and is either
' not merged or is the top-left cell of its merged area. Values copied underneath
' merged cells (see FillUnderMergedCells) are ignored, so they don't split groups.
Private Function StartsNewGroup(target As Range) As Boolean
    Dim c As Range
    For Each c In target.Cells
        If c.MergeCells Then
            If c.Address = c.MergeArea.Cells(1, 1).Address Then
                If RangeHasValue(c) Then
                    StartsNewGroup = True
                    Exit Function
                End If
            End If
        ElseIf RangeHasValue(c) Then
            StartsNewGroup = True
            Exit Function
        End If
    Next c
End Function

' True if any cell in the range holds a non-blank, non-error value.
Private Function RangeHasValue(target As Range) As Boolean
    Dim c As Range
    For Each c In target.Cells
        If Not IsError(c.Value) Then
            If Len(Trim$(CStr(c.Value))) > 0 Then
                RangeHasValue = True
                Exit Function
            End If
        End If
    Next c
End Function

' ---------------------------------------------------------------------------
' Filtering by CRQ-ID group
'   A CRQ-ID group starts on a row with a value in column A and runs down to
'   the row before the next value in column A. These macros show or hide whole
'   groups, so a CRQ-ID always stays together with all of its sub rows.
' ---------------------------------------------------------------------------

' Show only CRQ-ID groups where at least one row has the search text in column Q.
Sub FilterCRQ_ByColumnQ()
    Dim ws As Worksheet
    Dim searchText As String
    Dim lastRowTotal As Long
    Dim i As Long
    Dim groupStart As Long
    Dim groupMatches As Boolean
    Dim rowsToHide As Range
    Dim matchCount As Long

    Set ws = ActiveSheet

    searchText = InputBox("Show CRQ-IDs where column Q contains:", _
                          "Filter CRQ-IDs", "Construction Management Plan")
    If Len(Trim$(searchText)) = 0 Then Exit Sub

    On Error Resume Next
    lastRowTotal = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lastRowTotal < FIRST_DATA_ROW Then Exit Sub

    Application.ScreenUpdating = False

    ' Start from a clean slate so repeated filters don't stack up
    ws.Rows(FIRST_DATA_ROW & ":" & lastRowTotal).Hidden = False

    groupStart = 0
    For i = FIRST_DATA_ROW To lastRowTotal + 1
        ' A value in column A (or running past the last row) ends the current group
        If i > lastRowTotal Or StartsNewGroup(ws.Cells(i, "A")) Then
            If groupStart > 0 Then
                If groupMatches Then
                    matchCount = matchCount + 1
                Else
                    AddToRange rowsToHide, ws.Rows(groupStart & ":" & i - 1)
                End If
            End If
            groupStart = i
            groupMatches = False
        End If

        If i <= lastRowTotal And groupStart > 0 And Not groupMatches Then
            If Not IsError(ws.Cells(i, "Q").Value) Then
                If InStr(1, CStr(ws.Cells(i, "Q").Value), searchText, vbTextCompare) > 0 Then
                    groupMatches = True
                End If
            End If
        End If
    Next i

    If Not rowsToHide Is Nothing Then rowsToHide.EntireRow.Hidden = True

    Application.ScreenUpdating = True
    MsgBox matchCount & " CRQ-ID(s) contain """ & searchText & """ in column Q.", vbInformation
End Sub

' Unhide every data row (clears the CRQ-ID filter).
Sub ShowAllCRQ()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Rows(FIRST_DATA_ROW & ":" & ws.Rows.Count).Hidden = False
End Sub

Private Sub AddToRange(ByRef target As Range, addition As Range)
    If target Is Nothing Then
        Set target = addition
    Else
        Set target = Union(target, addition)
    End If
End Sub

' ---------------------------------------------------------------------------
' "Fix 2" for merged cells: copy each merged cell's value into the hidden cells
' underneath it, keeping the merges. Excel's filter then treats every row of a
' merged block as having the parent value. Re-run after editing merged values.
'   FillUnderAllMergedCells - every merged cell on the active sheet
'   FillUnderMergedCells    - merged cells in the data rows of columns A:G only
' ---------------------------------------------------------------------------
Sub FillUnderAllMergedCells()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    FillMergedIn ws, ws.UsedRange
End Sub

Sub FillUnderMergedCells()
    Dim ws As Worksheet
    Dim lastRowTotal As Long

    Dim firstRow As Long

    Set ws = ActiveSheet
    firstRow = 2   ' First data row (row 1 is the header)

    On Error Resume Next
    lastRowTotal = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If lastRowTotal < firstRow Then Exit Sub

    FillMergedIn ws, ws.Range(ws.Cells(firstRow, "A"), ws.Cells(lastRowTotal, "G"))
End Sub

Private Sub FillMergedIn(ws As Worksheet, src As Range)
    Dim tmp As Worksheet
    Dim c As Range
    Dim area As Range
    Dim topCell As Range
    Dim areaAddresses As Collection
    Dim addr As Variant

    ' Record every merged area before unmerging
    Set areaAddresses = New Collection
    For Each c In src.Cells
        If c.MergeCells Then
            If c.Address = c.MergeArea.Cells(1, 1).Address Then
                areaAddresses.Add c.MergeArea.Address
            End If
        End If
    Next c
    If areaAddresses.Count = 0 Then
        MsgBox "No merged cells found in " & src.Address(False, False) & ".", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    ' Keep a copy of the merged layout, at the same addresses, on a temporary sheet
    Set tmp = ws.Parent.Worksheets.Add
    src.Copy tmp.Range(src.Cells(1, 1).Address)

    ' Unmerge each block and copy its top-left value into the other cells
    For Each addr In areaAddresses
        Set area = ws.Range(addr)
        area.UnMerge
        Set topCell = area.Cells(1, 1)
        For Each c In area.Cells
            If c.Address <> topCell.Address Then c.Value = topCell.Value
        Next c
    Next addr

    ' Paste each block's formats back (like Format Painter): this re-merges
    ' the cells but keeps the values now stored underneath
    ws.Activate
    For Each addr In areaAddresses
        tmp.Range(addr).Copy
        ws.Range(addr).PasteSpecial xlPasteFormats
    Next addr

CleanUp:
    Application.CutCopyMode = False
    If Not tmp Is Nothing Then
        Application.DisplayAlerts = False
        tmp.Delete
        Application.DisplayAlerts = True
    End If
    ws.Activate
    ws.Range("A1").Select
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Stopped with an error: " & Err.Description, vbExclamation
    Else
        MsgBox areaAddresses.Count & " merged block(s) filled. You can now filter on these columns.", vbInformation
    End If
End Sub
