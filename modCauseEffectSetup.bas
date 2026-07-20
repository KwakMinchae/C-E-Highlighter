Attribute VB_Name = "modCauseEffectSetup"
Option Explicit

'==============================================================
' CAUSE & EFFECT MATRIX — SETUP (direct-coloring version)
'
' Four facts define everything:
'   1. Cause rows start at row 61          — FIXED constant
'   2. Effect columns start at column AW   — FIXED constant
'   3. Where the cause rows END            — auto-detected
'   4. Where the effect columns END        — auto-detected
'
' No Conditional Formatting here — that approach hit unpredictable
' Excel-specific formula behavior. This version just stores the
' four boundary numbers; the runtime module colors cells directly.
'==============================================================

Public Const CAUSE_DATA_START As Long = 61
Public Const EFFECT_COL_START As Long = 49   ' AW

' Run this to find out EXACTLY which cell is causing the "last column" scan to jump
' further right than expected — e.g.: DiagnoseWidth "Sheet2", 2, 200
Public Sub DiagnoseWidth(sheetName As String, rowFrom As Long, rowTo As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim r As Long, c As Long, worst As Long, worstRow As Long
    worst = 0
    For r = rowFrom To rowTo
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > worst Then
            worst = c
            worstRow = r
        End If
    Next r
    Debug.Print "Furthest-right content in rows " & rowFrom & "-" & rowTo & ":"
    Debug.Print "  Row " & worstRow & ", column " & worst & " (" & ws.Cells(worstRow, worst).Address(False, False) & _
                ") = '" & ws.Cells(worstRow, worst).Value & "'"
End Sub

Public Sub SetupOneSheet(sheetName As String)
    SetupSheetAnchors ThisWorkbook.Worksheets(sheetName)
    MsgBox "Setup complete for '" & sheetName & "'.", vbInformation
End Sub

Public Sub RunSetupAllSheets()
    Dim ws As Worksheet
    Dim failedSheets As String, okCount As Long
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Err.Clear
        SetupSheetAnchors ws
        If Err.Number <> 0 Then
            failedSheets = failedSheets & "  - " & ws.Name & ": " & Err.Description & vbCrLf
        Else
            okCount = okCount + 1
        End If
        On Error GoTo 0
    Next ws
    If Len(failedSheets) > 0 Then
        MsgBox "Setup finished." & vbCrLf & okCount & " sheet(s) OK." & vbCrLf & failedSheets, vbExclamation
    Else
        MsgBox "Setup completed successfully on all " & okCount & " sheets.", vbInformation
    End If
End Sub

Public Sub SetupSheetAnchors(ws As Worksheet)

    ' ---- End of cause rows: scan every column in the cause block, take the furthest down ----
    Dim causeDataEnd As Long
    causeDataEnd = FindLastRowAcross(ws, CAUSE_DATA_START, 1, EFFECT_COL_START - 1)
    If causeDataEnd < CAUSE_DATA_START Then
        Err.Raise vbObjectError + 1, , "No data found anywhere in columns A to " & (EFFECT_COL_START - 1) & _
                  " at or below row " & CAUSE_DATA_START & "."
    End If
    Dim paddedDataEnd As Long
    paddedDataEnd = causeDataEnd   ' no buffer — recalculated fresh every setup run anyway

    ' ---- End of effect columns: scan every row in the relevant block, take the furthest right ----
    ' Row 1 is deliberately excluded — that's where our own scratch/metadata cell lives,
    ' and including it would let leftover metadata contaminate this detection.
    Dim lastEffectCol As Long
    lastEffectCol = FindLastColAcross(ws, EFFECT_COL_START, 2, causeDataEnd)
    If lastEffectCol < EFFECT_COL_START Then
        Err.Raise vbObjectError + 2, , "No data found anywhere from column " & EFFECT_COL_START & _
                  " (AW) onward, in rows 1 to " & causeDataEnd & "."
    End If
    Dim paddedLastEffectCol As Long
    paddedLastEffectCol = lastEffectCol   ' no buffer — recalculated fresh every setup run anyway

    ' Hard safety caps — never let a detection glitch produce an out-of-range reference
    Const MAX_EXCEL_COL As Long = 16384
    Const MAX_EXCEL_ROW As Long = 1048576
    If paddedLastEffectCol > MAX_EXCEL_COL Then paddedLastEffectCol = MAX_EXCEL_COL
    If paddedDataEnd > MAX_EXCEL_ROW - 10 Then paddedDataEnd = MAX_EXCEL_ROW - 10

    ' ---- Reset any leftover coloring from a previous run across the padded block ----
    ws.Range(ws.Cells(1, 1), ws.Cells(paddedDataEnd, paddedLastEffectCol)).Interior.ColorIndex = xlNone

    ' ---- Store the four numbers as one compact metadata cell ----
    ' Placed at the FAR EDGE of the sheet (not just past the effect columns) so it can
    ' never be picked up by our own detection scans on a future re-run — that's exactly
    ' what caused the boundaries to drift after repeated testing.
    Dim scratchCol As Long
    scratchCol = ws.Columns.Count - 5
    Dim metaCell As Range
    Set metaCell = ws.Cells(1, scratchCol)
    metaCell.Value = CAUSE_DATA_START & "|" & paddedDataEnd & "|" & EFFECT_COL_START & "|" & paddedLastEffectCol
    metaCell.Font.Color = RGB(200, 200, 200)
    metaCell.Font.Size = 7
    DefineName ws, "CEMeta", metaCell

    ' ---- Minimal summary block: row number + active effect count ----
    Dim notesCol As Long
    notesCol = paddedLastEffectCol + 2
    SetupSummaryBlock ws, paddedDataEnd, notesCol

End Sub

' Scans every column in [colFrom, colTo] and returns the furthest-down populated row
' at or below startRow — bounded to a generous but FINITE range below startRow, so
' unrelated content far down the sheet (or in an unrelated section) can't be mistaken
' for the end of THIS table. Also skips rows that look like our own summary output.
Private Function FindLastRowAcross(ws As Worksheet, startRow As Long, colFrom As Long, colTo As Long) As Long
    Const MAX_ROWS_BELOW As Long = 500   ' generous headroom — no real sheet needs more causes than this
    Dim searchFloor As Long
    searchFloor = startRow + MAX_ROWS_BELOW

    Dim c As Long, r As Long, best As Long
    best = startRow - 1
    For c = colFrom To colTo
        r = ws.Cells(searchFloor, c).End(xlUp).Row
        Do
            If r < startRow Then Exit Do
            Dim cellVal As String
            cellVal = Trim(ws.Cells(r, c).Value & "")
            If cellVal <> "" And Left(cellVal, 2) <> "> " Then Exit Do   ' genuine, non-blank data — stop here
            r = r - 1   ' blank, or our own summary marker — keep looking further up
        Loop
        If r >= startRow And r > best Then best = r
    Next c
    FindLastRowAcross = best
End Function

' Scans every row in [rowFrom, rowTo] and returns the furthest-right populated column
' at or after startCol — bounded to a generous but FINITE range past startCol, so
' unrelated content elsewhere on the sheet (like a reference-documents box, or a stray
' cell far to the right) can't be mistaken for the end of THIS effect matrix.
Private Function FindLastColAcross(ws As Worksheet, startCol As Long, rowFrom As Long, rowTo As Long) As Long
    Const MAX_COLS_RIGHT As Long = 300   ' generous headroom — no real sheet needs more effects than this
    Dim searchEdge As Long
    searchEdge = startCol + MAX_COLS_RIGHT

    Dim r As Long, c As Long, best As Long
    best = startCol - 1
    For r = rowFrom To rowTo
        c = ws.Cells(r, searchEdge).End(xlToLeft).Column
        If c >= startCol And c > best And c <= searchEdge Then best = c
    Next r
    FindLastColAcross = best
End Function

Private Sub DefineName(ws As Worksheet, nm As String, rng As Range)
    On Error Resume Next
    ws.Names(nm).Delete
    On Error GoTo 0
    ws.Names.Add Name:=nm, RefersTo:=rng
End Sub

Private Sub SetupSummaryBlock(ws As Worksheet, dataEnd As Long, notesCol As Long)
    Dim summaryStartRow As Long
    summaryStartRow = dataEnd + 3

    Const MAX_SUMMARY_ROWS As Long = 60

    Dim r As Long
    Application.DisplayAlerts = False
    For r = summaryStartRow To summaryStartRow + MAX_SUMMARY_ROWS
        On Error Resume Next
        ws.Range(ws.Cells(r, 1), ws.Cells(r, notesCol)).UnMerge
        On Error GoTo 0
        ws.Range(ws.Cells(r, 1), ws.Cells(r, notesCol)).Merge
        ws.Cells(r, 1).Value = ""
        ws.Cells(r, 1).Interior.ColorIndex = xlNone
        ws.Rows(r).RowHeight = 18
    Next r
    Application.DisplayAlerts = True

    DefineName ws, "SummaryAnchorCell", ws.Cells(summaryStartRow, 1)
End Sub
