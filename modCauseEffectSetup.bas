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

' Effect-side header rows (confirmed directly against the live sheet — these are the
' TOP row of each merged label block, since that's where a merged cell's value lives)
Public Const ROW_REF_DOC As Long = 12
Public Const ROW_VOTING As Long = 23
Public Const ROW_DESCRIPTION As Long = 24
Public Const ROW_ACTION As Long = 44
Public Const ROW_TAG_NUMBER As Long = 50

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

' Lightweight refresh — call this from Worksheet_Change so adding new cause rows or
' effect columns updates the stored boundaries automatically, without needing to
' manually re-run SetupOneSheet. Only re-measures and re-stores the four numbers —
' does NOT clear colors or rebuild the summary block, so it stays fast even if it
' fires on every edit.
Public Sub RefreshBoundaries(ws As Worksheet, Optional Target As Range)
    If Not NameExists(ws, "CEMeta") Then Exit Sub   ' full setup hasn't run yet — nothing to refresh

    ' Cheap short-circuit: if we know which cell(s) were edited and they're safely
    ' inside the already-known table bounds, nothing could have changed the edges —
    ' skip the expensive scan entirely. This is the common case (filling in existing
    ' cells), so this avoids most of the cost on most edits.
    If Not Target Is Nothing Then
        Dim curMeta() As String
        curMeta = Split(ws.Names("CEMeta").RefersToRange.Value, "|")
        If UBound(curMeta) >= 3 Then
            Dim curDataEnd As Long, curEffEnd As Long
            curDataEnd = CLng(curMeta(1))
            curEffEnd = CLng(curMeta(3))
            Dim targetBottom As Long, targetRight As Long
            targetBottom = Target.Row + Target.Rows.Count - 1
            targetRight = Target.Column + Target.Columns.Count - 1
            If targetBottom < curDataEnd And targetRight < curEffEnd Then Exit Sub
        End If
    End If

    Dim causeDataEnd As Long
    causeDataEnd = FindLastRowAcross(ws, CAUSE_DATA_START, 1, EFFECT_COL_START - 1)
    If causeDataEnd < CAUSE_DATA_START Then Exit Sub

    Dim lastEffectCol As Long
    lastEffectCol = FindLastColAcross(ws, EFFECT_COL_START, 2, causeDataEnd)
    If lastEffectCol < EFFECT_COL_START Then Exit Sub

    ' Writing to metaCell below is itself a "cell changed" event — without turning
    ' events off first, it would immediately re-trigger this same sub on this same
    ' sheet, risking infinite recursion now that every sheet has Worksheet_Change.
    On Error GoTo CleanFail
    Application.EnableEvents = False
    Dim metaCell As Range
    Set metaCell = ws.Names("CEMeta").RefersToRange
    metaCell.Value = CAUSE_DATA_START & "|" & causeDataEnd & "|" & EFFECT_COL_START & "|" & lastEffectCol
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
End Sub

Private Function NameExists(ws As Worksheet, nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names(nm)
    NameExists = (Err.Number = 0) And Not (n Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

' Run this to wipe out old summary debris scattered from previous setup runs at
' different row positions. Automatically starts right after your real cause data ends
' (read from CEMeta) so it can never touch real data, and only touches column A —
' e.g.: CleanupOldSummaryDebris "Sheet2", 300
Public Sub CleanupOldSummaryDebris(sheetName As String, rowTo As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)

    If Not NameExists(ws, "CEMeta") Then
        MsgBox "Run SetupOneSheet on this sheet first — I need to know where your real data ends before I can clean up safely.", vbExclamation
        Exit Sub
    End If

    Dim meta() As String
    meta = Split(ws.Names("CEMeta").RefersToRange.Value, "|")
    Dim dataEnd As Long
    dataEnd = CLng(meta(1))

    Dim rowFrom As Long
    rowFrom = dataEnd + 1   ' safely below all real cause data — this can never touch it

    Dim r As Long
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    For r = rowFrom To rowTo
        On Error Resume Next
        ws.Cells(r, 1).UnMerge
        ws.Cells(r, 1).ClearContents
        ws.Cells(r, 1).Interior.ColorIndex = xlNone
        On Error GoTo 0
    Next r
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    MsgBox "Cleaned column A, rows " & rowFrom & " to " & rowTo & ".", vbInformation
End Sub

' Injects the Worksheet_SelectionChange / Worksheet_Change stub into EVERY sheet
' automatically, skipping any sheet that already has one. Requires "Trust access to
' the VBA project object model" enabled once (File > Options > Trust Center >
' Trust Center Settings > Macro Settings).
Public Sub InjectStubIntoAllSheets()
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim stubSelChange As String, stubChange As String
    stubSelChange = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & _
                    "    modCauseEffectHighlighter.HandleCauseSelection Me, Target" & vbCrLf & _
                    "End Sub"
    stubChange = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf & _
                 "    modCauseEffectSetup.RefreshBoundaries Me" & vbCrLf & _
                 "End Sub"

    Dim ws As Worksheet, comp As Object, codeMod As Object, existingCode As String
    Dim addedCount As Long, skippedCount As Long, addedSomething As Boolean

    For Each ws In ThisWorkbook.Worksheets
        Set comp = vbProj.VBComponents(ws.CodeName)
        Set codeMod = comp.CodeModule
        If codeMod.CountOfLines > 0 Then
            existingCode = codeMod.Lines(1, codeMod.CountOfLines)
        Else
            existingCode = ""
        End If
        addedSomething = False

        If InStr(1, existingCode, "Worksheet_SelectionChange", vbTextCompare) = 0 Then
            codeMod.InsertLines codeMod.CountOfLines + 1, stubSelChange
            addedSomething = True
        End If
        If InStr(1, existingCode, "Worksheet_Change", vbTextCompare) = 0 Then
            codeMod.InsertLines codeMod.CountOfLines + 1, stubChange
            addedSomething = True
        End If

        If addedSomething Then addedCount = addedCount + 1 Else skippedCount = skippedCount + 1
    Next ws

    MsgBox "Stub added to " & addedCount & " sheet(s)." & vbCrLf & _
           "Skipped " & skippedCount & " sheet(s) that already had matching code.", vbInformation
End Sub

' Upgrades the Worksheet_Change stub on every sheet to pass Target through, enabling
' the short-circuit optimization in RefreshBoundaries above. Safe to run any time —
' skips sheets that don't have the old exact line, so it won't touch anything else.
Public Sub UpgradeChangeStubsForPerformance()
    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject

    Dim ws As Worksheet, comp As Object, codeMod As Object
    Dim updatedCount As Long, skippedCount As Long, i As Long, lineText As String

    For Each ws In ThisWorkbook.Worksheets
        Set comp = vbProj.VBComponents(ws.CodeName)
        Set codeMod = comp.CodeModule
        Dim found As Boolean
        found = False
        For i = 1 To codeMod.CountOfLines
            lineText = codeMod.Lines(i, 1)
            If InStr(1, lineText, "modCauseEffectSetup.RefreshBoundaries Me", vbTextCompare) > 0 _
               And InStr(1, lineText, "Target", vbTextCompare) = 0 Then
                codeMod.ReplaceLine i, "    modCauseEffectSetup.RefreshBoundaries Me, Target"
                found = True
                Exit For
            End If
        Next i
        If found Then updatedCount = updatedCount + 1 Else skippedCount = skippedCount + 1
    Next ws

    MsgBox "Upgraded " & updatedCount & " sheet(s)." & vbCrLf & _
           skippedCount & " sheet(s) skipped (already upgraded, or no matching line found).", vbInformation
End Sub

Public Sub SetupOneSheet(sheetName As String)
    SetupSheetAnchors ThisWorkbook.Worksheets(sheetName)
    MsgBox "Setup complete for '" & sheetName & "'.", vbInformation
End Sub

Public Sub RunSetupAllSheets()
    Dim ws As Worksheet
    Dim failedSheets As String, okCount As Long

    ' Screen redraw and recalculation on every single cell write across 274 sheets is
    ' by far the slowest part of this — disabling both here cuts runtime dramatically.
    Dim prevCalc As XlCalculation
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

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

    Application.Calculation = prevCalc
    Application.ScreenUpdating = True

    If Len(failedSheets) > 0 Then
        MsgBox "Setup finished." & vbCrLf & okCount & " sheet(s) OK." & vbCrLf & failedSheets, vbExclamation
    Else
        MsgBox "Setup completed successfully on all " & okCount & " sheets.", vbInformation
    End If
End Sub

Public Sub SetupSheetAnchors(ws As Worksheet)
    ' This sub writes to many cells while building the table (colors, metadata, the
    ' summary block) — with Worksheet_Change now on every sheet, those writes would
    ' otherwise re-trigger RefreshBoundaries mid-setup. Guard against that here, and
    ' guarantee it's restored even if one of this sub's own validation errors fires.
    On Error GoTo CleanFail
    Application.EnableEvents = False

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

    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Dim errNum As Long, errDesc As String, errSrc As String
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source
    Application.EnableEvents = True
    Err.Raise errNum, errSrc, errDesc   ' re-raise so callers (e.g. RunSetupAllSheets) still see the real error
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

    ' Wipe out any OLD summary block left at a previous position before building the
    ' new one — otherwise every re-run leaves another patch of stray text behind.
    ' Only touches column A, cell by cell — never a wide range, which risks colliding
    ' with unrelated merged cells elsewhere on the sheet.
    If NameExists(ws, "SummaryAnchorCell") Then
        Dim oldRow As Long
        oldRow = ws.Names("SummaryAnchorCell").RefersToRange.Row
        Dim wr As Long
        Application.DisplayAlerts = False
        For wr = oldRow To oldRow + 200
            On Error Resume Next
            ws.Cells(wr, 1).UnMerge
            ws.Cells(wr, 1).ClearContents
            ws.Cells(wr, 1).Interior.ColorIndex = xlNone
            On Error GoTo 0
        Next wr
        Application.DisplayAlerts = True
    End If

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
