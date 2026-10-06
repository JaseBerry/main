' Sheet layout (three-level hierarchy):
'   A    - Filter assist column (Y), merged like the CRQ-ID. Not used for grouping.
'   B:E  - Level 1 (CRQ-ID in B). One row here can own several F:H rows below it.
'   F:H  - Level 2. One row here can own several "I onwards" rows below it.
'   I+   - Level 3. Column L holds the status. A row with L filled is a "main row";
'          rows with L blank are "sub rows" that inherit the status above them.
'
' Colouring:
'   - Main rows: K to last column, coloured by their own status.
'   - Sub rows:  W to last column, coloured by the carried-over status.
'   - Fills from K onwards are extra light; F:H uses the stronger washed-out fills.
'   - F:H block: green if any status row in the block is Accepted,
'                otherwise yellow if any is Under Review, otherwise cleared.
'                Text in F:H (and everything left of it) stays black.
'
' If you insert or move columns, update the letters below - nothing else.

Private Const FIRST_DATA_ROW As Long = 2   ' Row 1 is the header row
Private Const COL_CRQ As String = "B"          ' CRQ-ID, first column of level 1
Private Const COL_L2_FIRST As String = "F"     ' Level 2 block (coloured green/yellow)
Private Const COL_L2_LAST As String = "H"
Private Const COL_MAIN_FROM As String = "K"    ' Main rows are coloured from here
Private Const COL_STATUS As String = "L"       ' Accepted / Under Review / ...
Private Const COL_SUB_FROM As String = "W"     ' Sub rows are coloured from here
Private Const COL_SEARCH As String = "R"       ' Searched by FilterCRQ_BySearchText

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
    ' (Necessary because sub-rows extend further down than the status column)
    On Error Resume Next
    lastRowTotal = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find(What:="*", After:=ws.Range("A1"), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    On Error GoTo 0

    ' If there is no data below the header, exit
    If lastRowTotal < FIRST_DATA_ROW Then Exit Sub

    ' Always format at least to the sub-row start column
    If lastCol < ws.Range(COL_SUB_FROM & "1").Column Then lastCol = ws.Range(COL_SUB_FROM & "1").Column

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo CleanUp

    statusValue = ""
    blockStart = 0

    For i = FIRST_DATA_ROW To lastRowTotal

        ' A value anywhere in level 1 or level 2 starts a new level 2 block.
        ' Colour the block that just ended, then start tracking the new one.
        If StartsNewGroup(ws.Range(ws.Cells(i, COL_CRQ), ws.Cells(i, COL_L2_LAST))) Then
            If blockStart > 0 Then
                ColourEFGBlock ws, blockStart, i - 1, hasAccepted, hasUnderReview
            End If
            blockStart = i
            hasAccepted = False
            hasUnderReview = False
            statusValue = ""   ' Don't carry a status across into a new block
        End If

        ' If the status column has a value, it's a "Main Row". Update the status and format from COL_MAIN_FROM.
        If RangeHasValue(ws.Cells(i, COL_STATUS)) Then
            statusValue = LCase$(Trim$(CStr(ws.Cells(i, COL_STATUS).Value)))
            Set formatRange = ws.Range(ws.Cells(i, COL_MAIN_FROM), ws.Cells(i, lastCol))

            If statusValue = "accepted" Then hasAccepted = True
            If statusValue = "under review" Then hasUnderReview = True

        ' If the status is blank, it's a "Sub Row". Keep the previous status but format ONLY from COL_SUB_FROM.
        Else
            Set formatRange = ws.Range(ws.Cells(i, COL_SUB_FROM), ws.Cells(i, lastCol))
        End If

        ApplyStatusColours formatRange, statusValue, True
    Next i

    ' Colour the final level 2 block
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

' Colour the level 2 columns for every row in the block based on the statuses found within it.
Private Sub ColourEFGBlock(ws As Worksheet, firstRow As Long, lastRow As Long, _
                           hasAccepted As Boolean, hasUnderReview As Boolean)
    Dim efgRange As Range
    Set efgRange = ws.Range(ws.Cells(firstRow, COL_L2_FIRST), ws.Cells(lastRow, COL_L2_LAST))

    If hasAccepted Then
        ApplyStatusColours efgRange, "accepted"
    ElseIf hasUnderReview Then
        ApplyStatusColours efgRange, "under review"
    Else
        ApplyStatusColours efgRange, ""
    End If

    ' Keep text black in the level 2 columns; only the fill changes
    efgRange.Font.Color = RGB(0, 0, 0)
End Sub

' Apply washed-out fills and fonts for a status (expects lower-case, trimmed text).
' extraLight = True uses even paler fills (used for the status rows).
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
' merged cells (see FillUnderAllMergedCells) are ignored, so they don't split groups.
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
'   A CRQ-ID group starts on a row with a value in the CRQ-ID column and runs
'   down to the row before the next CRQ-ID. These macros show or hide whole
'   groups, so a CRQ-ID always stays together with all of its sub rows.
' ---------------------------------------------------------------------------

' Show only CRQ-ID groups where at least one row has the search text in COL_SEARCH.
Sub FilterCRQ_BySearchText()
    Dim ws As Worksheet
    Dim searchText As String
    Dim lastRowTotal As Long
    Dim i As Long
    Dim groupStart As Long
    Dim groupMatches As Boolean
    Dim rowsToHide As Range
    Dim matchCount As Long

    Set ws = ActiveSheet

    searchText = InputBox("Show CRQ-IDs where column " & COL_SEARCH & " contains:", _
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
        ' A new CRQ-ID (or running past the last row) ends the current group
        If i > lastRowTotal Or StartsNewGroup(ws.Cells(i, COL_CRQ)) Then
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
            If Not IsError(ws.Cells(i, COL_SEARCH).Value) Then
                If InStr(1, CStr(ws.Cells(i, COL_SEARCH).Value), searchText, vbTextCompare) > 0 Then
                    groupMatches = True
                End If
            End If
        End If
    Next i

    If Not rowsToHide Is Nothing Then rowsToHide.EntireRow.Hidden = True

    Application.ScreenUpdating = True
    MsgBox matchCount & " CRQ-ID(s) contain """ & searchText & """ in column " & COL_SEARCH & ".", vbInformation
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
' ---------------------------------------------------------------------------
Sub FillUnderAllMergedCells()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    FillMergedIn ws, ws.UsedRange
End Sub

Private Sub FillMergedIn(ws As Worksheet, src As Range)
    Dim tmp As Worksheet
    Dim c As Range
    Dim area As Range
    Dim topCell As Range
    Dim areaAddresses As Collection
    Dim addr As Variant
    Dim n As Long
    Dim total As Long
    Dim calcMode As XlCalculation
    Dim finished As Boolean

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler   ' Esc jumps to CleanUp
    On Error GoTo CleanUp

    ' Step 1: record every merged area before unmerging
    Set areaAddresses = New Collection
    total = src.Cells.Count
    For Each c In src.Cells
        n = n + 1
        If n Mod 5000 = 0 Then ShowProgress "Step 1 of 3: scanning cells", n, total
        If c.MergeCells Then
            If c.Address = c.MergeArea.Cells(1, 1).Address Then
                areaAddresses.Add c.MergeArea.Address
            End If
        End If
    Next c

    If areaAddresses.Count = 0 Then
        finished = True
        GoTo CleanUp
    End If

    ' Keep a copy of the merged layout, at the same addresses, on a temporary sheet
    Application.StatusBar = "Copying layout to a temporary sheet..."
    Set tmp = ws.Parent.Worksheets.Add
    src.Copy tmp.Range(src.Cells(1, 1).Address)

    ' Step 2: unmerge each block and copy its top-left value into the other cells
    total = areaAddresses.Count
    n = 0
    For Each addr In areaAddresses
        n = n + 1
        If n Mod 100 = 0 Then ShowProgress "Step 2 of 3: filling block", n, total
        Set area = ws.Range(addr)
        area.UnMerge
        Set topCell = area.Cells(1, 1)
        If topCell.HasFormula Then
            For Each c In area.Cells
                If c.Address <> topCell.Address Then c.Value = topCell.Value
            Next c
        Else
            area.Value = topCell.Value
        End If
    Next addr

    ' Step 3: paste each block's formats back (like Format Painter): this
    ' re-merges the cells but keeps the values now stored underneath
    ws.Activate
    n = 0
    For Each addr In areaAddresses
        n = n + 1
        If n Mod 100 = 0 Then ShowProgress "Step 3 of 3: re-merging block", n, total
        tmp.Range(addr).Copy
        ws.Range(addr).PasteSpecial xlPasteFormats
    Next addr

    finished = True

CleanUp:
    Application.CutCopyMode = False
    If Not tmp Is Nothing Then
        Application.DisplayAlerts = False
        tmp.Delete
        Application.DisplayAlerts = True
    End If
    ws.Activate
    ws.Range("A1").Select
    Application.StatusBar = False
    Application.Calculation = calcMode
    Application.EnableEvents = True
    Application.EnableCancelKey = xlInterrupt
    Application.ScreenUpdating = True

    If Not finished Then
        MsgBox "Stopped before finishing" & IIf(Err.Number <> 0, ": " & Err.Description, ".") & vbCrLf & _
               "Some merged cells may now be unmerged. Close WITHOUT saving and reopen your copy.", vbExclamation
    ElseIf areaAddresses.Count = 0 Then
        MsgBox "No merged cells found in " & src.Address(False, False) & ".", vbInformation
    Else
        MsgBox areaAddresses.Count & " merged block(s) filled. You can now filter on these columns.", vbInformation
    End If
End Sub

' Show progress in Excel's status bar (bottom-left) and keep Excel responsive.
Private Sub ShowProgress(stepText As String, done As Long, total As Long)
    Application.StatusBar = stepText & " " & Format(done, "#,##0") & " of " & _
                            Format(total, "#,##0") & " (" & Format(done / total, "0%") & ")  - press Esc to cancel"
    DoEvents
End Sub
