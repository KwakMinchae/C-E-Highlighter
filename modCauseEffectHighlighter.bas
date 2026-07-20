Attribute VB_Name = "modCauseEffectHighlighter"
Option Explicit

'==============================================================
' CAUSE & EFFECT MATRIX — RUNTIME HANDLER (direct-coloring version)
'
'   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'       modCauseEffectHighlighter.HandleCauseSelection Me, Target
'   End Sub
'
' No Conditional Formatting, no formulas — this directly sets
' .Interior.Color on the whole selected row and every matching
' effect column, every click. Simple, predictable, easy to verify.
'==============================================================

Public Sub HandleCauseSelection(ws As Worksheet, Target As Range)

    ' Guard against absurdly large selections (e.g. clicking a row/column header) —
    ' never touches .MergeArea at all, so it can't hit the 1004 error that caused before.
    ' Normal single clicks, merged or not, just use Target.Row directly — this is exactly
    ' what the original working code did, and merging never actually broke that.
    If Target.Cells.Count > 1000 Then Exit Sub
    Dim effRow As Long
    effRow = Target.Row

    If Not NameExists(ws, "CEMeta") Then Exit Sub

    Dim meta() As String
    meta = Split(ws.Names("CEMeta").RefersToRange.Value, "|")
    If UBound(meta) < 3 Then Exit Sub

    Dim dataStart As Long, dataEnd As Long, effStart As Long, effEnd As Long
    dataStart = CLng(meta(0))
    dataEnd = CLng(meta(1))
    effStart = CLng(meta(2))
    effEnd = CLng(meta(3))

    If effRow < dataStart Or effRow > dataEnd Then Exit Sub

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- 1. Clear everything in the relevant block first ----
    ws.Range(ws.Cells(1, 1), ws.Cells(dataEnd, effEnd)).Interior.ColorIndex = xlNone

    ' ---- 2. Highlight the ENTIRE selected row, whole width, no exceptions ----
    ws.Range(ws.Cells(effRow, 1), ws.Cells(effRow, effEnd)).Interior.Color = RGB(255, 220, 50)

    ' ---- 3. Highlight the ENTIRE column (top to bottom) for every effect column with a mark ----
    Dim j As Long, effCount As Long
    effCount = 0
    For j = effStart To effEnd
        If Trim(ws.Cells(effRow, j).Value & "") <> "" Then
            ws.Range(ws.Cells(1, j), ws.Cells(dataEnd, j)).Interior.Color = RGB(255, 220, 50)
            effCount = effCount + 1
        End If
    Next j

    ' ---- Minimal summary: row number + active effect count ----
    If NameExists(ws, "SummaryAnchorCell") Then
        Dim summaryAnchor As Range
        Set summaryAnchor = ws.Names("SummaryAnchorCell").RefersToRange
        With ws.Cells(summaryAnchor.Row, 1)
            .Value = "> Row " & effRow & "  |  ACTIVE EFFECTS: " & effCount
            .Font.Bold = True
            .Font.Size = 10
            .Font.Color = RGB(31, 78, 121)
            .Interior.Color = RGB(189, 215, 238)
        End With
    End If

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Function NameExists(ws As Worksheet, nm As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ws.Names(nm)
    NameExists = (Err.Number = 0) And Not (n Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function
