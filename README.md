# Cause & Effect Matrix — Installation & Setup Guide

## What this does

Click any cause row → the whole row and every matching effect column highlight
yellow. A readable summary of active effects is written below the table (the
effect header labels are rotated/vertical in the original sheet — this writes
them out horizontally instead). Boundaries auto-adjust as rows/columns are added.

## Files

- `modCauseEffectSetup.bas` — one-time setup logic, plus the tools to roll it out
  across all 274 sheets at once
- `modCauseEffectHighlighter.bas` — runs on every click, does the highlighting

---

## Step 1 — Get the file to actually open with macros enabled

Skip this section if your file already opens without security warnings.

1. If Windows blocked the file entirely (**"Microsoft has blocked macros from
   running because the source of this file is untrusted"**, no Enable button):
   close Excel, right-click the file in File Explorer → **Properties** → check
   **Unblock** at the bottom of the General tab → OK → reopen.
   - If there's no Unblock checkbox, that's an IT-managed policy — contact your
     IT/helpdesk to allowlist the file, this can't be fixed from Excel itself.
2. Once open, if you see a yellow **"SECURITY WARNING — Macros have been
   disabled"** bar with an **Enable Content** button, click it.
3. If no bar appears at all: **File → Options → Trust Center → Trust Center
   Settings → Macro Settings → "Disable VBA macros with notification"** → OK →
   close and fully reopen the file → click Enable Content when the bar appears.

## Step 2 — Enable one more setting (needed for the automated rollout in Step 4)

**File → Options → Trust Center → Trust Center Settings → Macro Settings** →
check **"Trust access to the VBA project object model"** → OK.

## Step 3 — Import the two modules

1. **Alt+F11** to open the VBA editor.
2. Right-click the workbook in the Project tree → **Import File...** → select
   `modCauseEffectSetup.bas`. Repeat for `modCauseEffectHighlighter.bas`.

## Step 4 — Roll out to all 274 sheets

In the Immediate window (**Ctrl+G**), run each line separately (click the line,
press Enter, wait for its popup before running the next):

```
InjectStubIntoAllSheets
```
Adds the click-handler hookup to every sheet automatically. Popup confirms how
many sheets were updated vs. already had it.

```
RunSetupAllSheets
```
Detects each sheet's table boundaries and builds its summary block. Takes a
couple of minutes across 274 sheets (mostly the summary block formatting) —
this is expected, not a hang. Popup confirms success or lists any sheet that
failed and why.

## Step 5 — Test

Click any real cause row on any sheet. The row and its matching effect column(s)
should highlight yellow, and a summary line should appear below the table.

---

## Maintenance

- **Adding causes/effects inside the existing table bounds** — nothing needed,
  updates live via the auto-refresh hook.
- **Adding causes/effects right past the current known edge** — one extra row
  past the edge still works via a small built-in buffer; beyond that, re-run
  setup on that one sheet:
  ```
  SetupOneSheet "SheetNameHere"
  ```
- **Stray leftover summary text** from testing/edits — safe to run any time,
  only ever touches column A below your real data:
  ```
  CleanupOldSummaryDebris "SheetNameHere", 300
  ```

## Troubleshooting

| Symptom | Likely cause |
|---|---|
| "Compile error: Sub or Function not defined" | One of the two modules failed to import correctly, or there's a duplicate/corrupted module. Debug → Compile VBAProject to find the exact broken line. |
| Clicking does nothing, on one specific sheet only | That sheet is missing its stub — re-run `InjectStubIntoAllSheets`, or check the sheet's code module directly. |
| Clicking does nothing, on every sheet | Events may be stuck disabled from an earlier error. Run `Application.EnableEvents = True` in the Immediate window. |
| Setup errors on a specific sheet | That sheet's layout may not match the shared template — check the exact error message, it names the row/column it expected. |

## Customization

Fixed positions and effect-header row constants — top of `modCauseEffectSetup.bas`:

```vba
Public Const CAUSE_DATA_START As Long = 61
Public Const EFFECT_COL_START As Long = 49   ' AW
Public Const ROW_REF_DOC As Long = 12
Public Const ROW_VOTING As Long = 23
Public Const ROW_DESCRIPTION As Long = 24
Public Const ROW_ACTION As Long = 44
Public Const ROW_TAG_NUMBER As Long = 50
```

Highlight color and summary text styling — inside `HandleCauseSelection` in
`modCauseEffectHighlighter.bas`.
