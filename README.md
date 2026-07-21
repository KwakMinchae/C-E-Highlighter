# Cause & Effect Matrix Highlighter

VBA tool for Excel Cause & Effect (C&E) matrix sheets. Click any cause row and the
entire row plus every matching effect column highlights automatically, with a
readable text summary of active effects written below the table.

## What it does

- Click a cause row → highlights the **whole row** and **every effect column** with
  a mark in that row (full column height, no exceptions)
- Writes a readable summary strip below the table. There is one line per active effect,
  since the effect header labels are rotated/vertical and hard to read at a glance
- Auto-adapts as rows/columns are added. No manual reconfiguration needed for
  edits made *inside* the existing table bounds, and a live auto-refresh handles
  edits at the edges too

## Files

| File | Purpose |
|---|---|
| `modCauseEffectSetup.bas` | Run once per sheet. Detects table boundaries and prepares the sheet. |
| `modCauseEffectHighlighter.bas` | Runs on every click. Does the actual highlighting. |

## How it works

Four facts define the whole table. Two are fixed (same on every sheet, since they
share one template); two are auto-detected per sheet:

| | Value | How it's determined |
|---|---|---|
| Cause rows start | Row 61 | Fixed constant |
| Effect columns start | Column AW (49) | Fixed constant |
| Cause rows end | *(varies)* | Auto-detected — scans across the cause block's columns, bounded to a sane range, ignoring its own leftover output |
| Effect columns end | *(varies)* | Auto-detected — scans across the effect block's rows, bounded to a sane range |

Highlighting is done with direct cell coloring (`.Interior.Color`), not Conditional
Formatting — CF formulas hit unpredictable Excel-specific behavior in early
versions and were dropped in favor of something simpler to verify.

## Installation (per sheet)

1. Open the VBA editor: **Alt+F11**
2. Import both modules: right-click the workbook in the Project tree →
   **Import File...** → select `modCauseEffectSetup.bas`. Repeat for
   `modCauseEffectHighlighter.bas`.
3. Double-click the target sheet in the Project tree and paste in:

   ```vba
   Private Sub Worksheet_SelectionChange(ByVal Target As Range)
       modCauseEffectHighlighter.HandleCauseSelection Me, Target
   End Sub

   Private Sub Worksheet_Change(ByVal Target As Range)
       modCauseEffectSetup.RefreshBoundaries Me
   End Sub
   ```

4. In the Immediate window (**Ctrl+G**), click the line and press **Enter**:

   ```
   SetupOneSheet "SheetNameHere"
   ```

5. Click a real cause row to test — the row and its matching effect column(s)
   should highlight yellow, and a summary line should appear below the table.

### Rolling out to all sheets

Repeat steps 3–4 for each sheet, or call `RunSetupAllSheets()` from the Immediate
window to run setup on every sheet in the workbook at once (step 3's stub still
needs to be added to each sheet individually).

## Maintenance

- **Adding causes/effects inside the existing table** — no action needed;
  `Worksheet_Change` refreshes the boundaries automatically.
- **Adding causes/effects past the current known edge** — one extra row past the
  edge works automatically (a small buffer is built in); beyond that, re-run
  `SetupOneSheet "SheetName"` once.
- **Cleaning up stray summary text** left over from earlier testing/edits:

  ```
  CleanupOldSummaryDebris "SheetName", 300
  ```

  (adjust `300` to a row comfortably below your table; this only ever touches
  column A and only below your real data, so it can't affect real content)

## Customization

Formatting and header-row constants live at the top of `modCauseEffectSetup.bas`:

```vba
Public Const CAUSE_DATA_START As Long = 61
Public Const EFFECT_COL_START As Long = 49   ' AW
Public Const ROW_REF_DOC As Long = 12
Public Const ROW_VOTING As Long = 23
Public Const ROW_DESCRIPTION As Long = 24
Public Const ROW_ACTION As Long = 44
Public Const ROW_TAG_NUMBER As Long = 50
```

Highlight color and summary text styling (font, size, colors) are set in
`modCauseEffectHighlighter.bas`, inside `HandleCauseSelection`.
