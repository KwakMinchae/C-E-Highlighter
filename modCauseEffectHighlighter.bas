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

    If effRow < dataStart Or effRow > dataEnd + 1 Then Exit Sub

    ' If this click is on that one extra "+1" row, extend the working bottom bound to
    ' cover it too — otherwise the column highlight and next-click clear would stop
    ' one row short of the row actually being highlighted.
    Dim effectiveEnd As Long
    effectiveEnd = dataEnd
    If effRow > effectiveEnd Then effectiveEnd = effRow

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ---- 1. Clear everything in the relevant block first ----
    ws.Range(ws.Cells(1, 1), ws.Cells(effectiveEnd, effEnd)).Interior.ColorIndex = xlNone

    ' ---- 2. Highlight the ENTIRE selected row, whole width, no exceptions ----
    ws.Range(ws.Cells(effRow, 1), ws.Cells(effRow, effEnd)).Interior.Color = RGB(255, 220, 50)

    ' ---- 3. Highlight the ENTIRE column (top to bottom) for every effect column with a mark ----
    ' Also collect readable text for each active effect, since the header labels are
    ' rotated/vertical and hard to read — this writes them out horizontally below instead.
    Dim j As Long, effCount As Long
    effCount = 0
    Dim summaryAnchor As Range
    Dim haveSummaryAnchor As Boolean
    haveSummaryAnchor = NameExists(ws, "SummaryAnchorCell")
    If haveSummaryAnchor Then Set summaryAnchor = ws.Names("SummaryAnchorCell").RefersToRange

    For j = effStart To effEnd
        If Trim(ws.Cells(effRow, j).Value & "") <> "" Then
            ws.Range(ws.Cells(1, j), ws.Cells(effectiveEnd, j)).Interior.Color = RGB(255, 220, 50)
            effCount = effCount + 1

            If haveSummaryAnchor Then
                Dim effNo As Long
                effNo = j - effStart + 1
                Dim outRow As Long
                outRow = summaryAnchor.Row + effCount
                With ws.Cells(outRow, 1)
                    .Value = "  Effect " & effNo & _
                             "   |   Tag: " & Trim(ws.Cells(modCauseEffectSetup.ROW_TAG_NUMBER, j).Value & "") & _
                             "   |   Action: " & Trim(ws.Cells(modCauseEffectSetup.ROW_ACTION, j).Value & "") & _
                             "   |   Desc: " & Trim(ws.Cells(modCauseEffectSetup.ROW_DESCRIPTION, j).Value & "") & _
                             "   |   V: " & Trim(ws.Cells(modCauseEffectSetup.ROW_VOTING, j).Value & "") & _
                             "   |   REF: " & Trim(ws.Cells(modCauseEffectSetup.ROW_REF_DOC, j).Value & "")
                    .Font.Name = "Arial"
                    .Font.Size = 10
                    .Font.Bold = False
                    .Font.Color = RGB(0, 0, 0)
                    .Interior.Color = RGB(198, 239, 206)
                End With
            End If
        End If
    Next j

    ' ---- Summary banner + clear any leftover rows from a previous click with more effects ----
    If haveSummaryAnchor Then
        With ws.Cells(summaryAnchor.Row, 1)
            .Value = "> Row " & effRow & "  |  ACTIVE EFFECTS: " & effCount
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Interior.Color = RGB(0, 176, 80)
        End With

        Dim clearRow As Long
        For clearRow = summaryAnchor.Row + effCount + 1 To summaryAnchor.Row + 60
            If Trim(ws.Cells(clearRow, 1).Value & "") = "" Then Exit For
            ws.Cells(clearRow, 1).Value = ""
            ws.Cells(clearRow, 1).Interior.ColorIndex = xlNone
        Next clearRow
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
