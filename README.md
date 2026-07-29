**Language / Язык:**  English  ·  [Русский](README.ru.md)

# MXL Column Editor

A desktop tool for bulk-editing register columns in **1C:Enterprise spreadsheet
documents** (`.mxl`), with a pixel-accurate preview rendered by 1C itself.

Built for maintaining expected-result templates: when a register gains, loses or
renames a column in the configuration, every `.mxl` template in a test suite has to
follow. Doing that by hand across a hundred files is slow and easy to get wrong.

- Add, rename or delete a column in any register, across a whole folder at once
- See exactly what each file would gain or lose before writing anything
- Preview the printform as 1C draws it, embedded in the window
- Every run is backed up and verified byte-for-byte

Pure Python standard library plus Tkinter — no dependencies required for the core
editing. A single-file Windows `.exe` can be built for colleagues who have no
Python at all.

---

## Contents

- [Quick start](#quick-start)
- [Requirements](#requirements)
- [Using the editor](#using-the-editor)
  - [1. Working folder](#1-working-folder)
  - [2. Registers](#2-registers)
  - [3. Modifications](#3-modifications)
  - [4. Files](#4-files)
  - [Applying changes](#applying-changes)
- [Previewing a printform](#previewing-a-printform)
- [Setting up the 1C renderer](#setting-up-the-1c-renderer)
  - [Building MxlToHtmlService.epf](#building-mxltohtmlserviceepf)
- [Building a Windows .exe](#building-a-windows-exe)
- [Troubleshooting](#troubleshooting)
- [How it works](#how-it-works)
- [Repository layout](#repository-layout)
- [Credits](#credits)
- [License](#license)

---

## Quick start

**With Python:**

```bash
git clone https://github.com/SergiosZZz/mxl-column-editor.git
cd mxl-column-editor
python mxl_column_editor.py
```

**Without Python** — download the release, unzip, run `MXL Column Editor.exe`.

Everything except the 1C preview works immediately. The preview needs a local
1C:Enterprise installation; see [Setting up the 1C renderer](#setting-up-the-1c-renderer).

---

## Requirements

| | Needed for |
|---|---|
| Python 3.6+ with Tkinter | Running from source. The official Windows and macOS installers include Tkinter; on Debian/Ubuntu install `python3-tk`. |
| `pip install tkinterweb` | *Optional.* Embeds the 1C render in the window. Without it, previews open in your browser. |
| 1C:Enterprise (licensed) | *Optional.* Any preview at all. Column editing is unaffected. |
| PyInstaller | *Optional.* Only to build the `.exe`. |

---

## Using the editor

<img width="1916" height="990" alt="image" src="https://github.com/user-attachments/assets/08361130-df17-42e9-9602-35aa3472c935" />



The window is a sidebar of four pages, with a status bar and **Apply changes**
always visible at the bottom. Each page shows a count, so you can see at a glance
how much is loaded.

### 1. Working folder

*Browse…* selects the folder holding your `.mxl` files and scans it immediately.
Subfolders are not searched. After a successful scan the app moves to Registers.

### 2. Registers

Every register found across the folder, with how many files contain it and how many
distinct column names it has. Select one to list its columns on the right.

**Double-click a column** to jump straight to *Add modification* with that register
and column pre-filled as the anchor and the cursor waiting in the New column box —
usually the fastest way to build a rule.

### 3. Modifications

The queue of changes. *+ Add modification* opens a dialog with three tabs.

**Add column**

| Field | Meaning |
|---|---|
| Register | Which register table to change, e.g. `Inventory and expenses` |
| Position | `after` or `before` the anchor column |
| this column | The anchor, e.g. `Posting content` |
| New column | Header caption to add |
| Value in data rows | Constant written into every data row (blank allowed) |

The dialog shows the **resulting column order** as you type, and warns if the name
already exists or the anchor cannot be found.

**Rename column** — change a caption. Only the header cell is touched: column order,
widths, numbering and every data row are left exactly as they were, and the byte
diff is literally just the caption text.

**Delete column** — remove a column and every cell under it; the remaining columns
close up.

All three appear in the queue with an Action column and can be mixed in one run.
Registers and columns are editable comboboxes, so you can type a name the scan did
not find — useful when only some files have it yet.

> Adding a column then deleting it restores the file byte-for-byte, as does renaming
> a column and renaming it back. Both round trips are part of the test suite.

### 4. Files

Every `.mxl` with its size, register count, widest register, and what your
modifications would do to it — colour-coded: green for columns that would be added
or removed, amber for a missing anchor, grey for "register not present", red for a
file that could not be parsed.

The **Use** column controls scope. *Apply to all files* is ticked by default.
Untick it and nothing is included to begin with; the list becomes multi-select and
you pick files deliberately:

- **Ctrl-click** to add files to the highlight, **Shift-click** for a range
- **Right-click** for *Include in the run*, *Exclude from the run*, *Toggle*,
  *Include every eligible file*, *Include none*, and the preview commands
- **Space** toggles everything highlighted
- clicking a single **Use** cell toggles just that file

Files where nothing would change show an en dash and can never be included.

### Applying changes

**Apply changes** confirms the scope, then rewrites the files in place with a
progress bar and per-file status. The log panel at the bottom expands for detail and
opens automatically when a run starts.

Three safety options, all on by default:

- **Back up originals** — copies every file being processed into
  `_mxl_backup_<timestamp>\` inside the working folder first.
- **Skip if column already exists** — re-running never creates duplicates.
- **Verify after writing** — re-parses each written file against its backup and
  checks the row count is unchanged, the new column sits in the requested position
  exactly once, every data row holds the requested value, no existing cell text was
  altered, and the document preamble and trailer are byte-identical.

---

## Previewing a printform

Select a file and press **Preview printform** (or double-click it). The window has
two tabs:

- **1C preview** — opens first and starts rendering at once: the document is handed
  to 1C:Enterprise and drawn exactly as the platform draws it, embedded in the
  window. *Open in browser* sends the same render to your default browser.
- **Grid** — a built-in text grid of the whole document. Instant, needs nothing
  installed, and tints newly inserted cells green.

With modifications queued, both views offer **With modifications** and **Original**,
so you can flip between them and see exactly what your change did. Switching
re-renders automatically. Your modifications are applied to a *temporary copy*, so
the file on disk is never touched by previewing.

---

## Setting up the 1C renderer

The preview is not a reimplementation of MXL layout — the document is handed to the
platform, which does the work:

```bsl
SpreadsheetDocument = New SpreadsheetDocument;
SpreadsheetDocument.Read(InputFileName);
SpreadsheetDocument.Write(OutputFileName, SpreadsheetDocumentFileType.HTML);
```

`SpreadsheetDocument.Read` is unavailable in the thin client, so this runs in server
context and therefore needs an infobase. On first use the tool creates a small
service infobase from the bundled `onec\MxlRendererTemplate.dt` and reuses it
afterwards.

**Point the editor at your client:** *Tools → 1C:Enterprise renderer…* — choose the
thin client `1cv8c.exe`. `1cv8.exe` must sit beside it, as it is used once to create
the service infobase. *Auto-detect* searches the usual
`Program Files\1cv8\<version>\bin` locations.

### Building MxlToHtmlService.epf

Out of the box every preview starts a fresh 1C process: a few seconds and a window
flash each time. The **resident** processor is launched once and then serves a queue,
so later previews are near-instant. It has to be built once in Designer:

1. Create a new external data processor.
2. Add a **managed form**.
3. Set that form as the data processor's **default form** (`DefaultForm`).
4. Open the form's **Module** tab — the *form* module, not the object module — and
   paste the whole of [`onec/MxlToHtmlService.bsl`](onec/MxlToHtmlService.bsl).
5. **Wire the OnOpen event.** A procedure merely *named* `OnOpen` is not the form's
   handler; without this the form opens, nothing runs, and there is no error
   anywhere. Select the form's root element, press **F4**, and in **Events** click
   the magnifier beside **OnOpen** so the field reads `OnOpen`.
6. Save as `MxlToHtmlService.epf` in the `onec\` folder.

Then run *Tools → 1C renderer diagnostics…*, which walks the whole path and prints
where it stops. Full details in [`onec/README.md`](onec/README.md).

---

## Building a Windows .exe

On a Windows machine with Python, double-click **`build_windows_exe.bat`**. It
installs PyInstaller, bundles the `onec\` assets and produces:

```
dist\MXL Column Editor.exe
dist\onec\                     rebuildable .epf files
```

Hand over the whole `dist\` folder. The `.exe` alone works too — it carries its own
copy of `onec\` — but shipping the folder means an `.epf` can be rebuilt in Designer
and dropped in without rebuilding the program: assets beside the `.exe` take
precedence over the bundled ones.

Install `tkinterweb` **before** building if you want the embedded preview in the
`.exe`; a frozen program cannot pip-install anything later.

> Expect an antivirus false positive the first time — PyInstaller executables often
> trip heuristics.

---

## Troubleshooting

Start with *Tools → 1C renderer diagnostics…*. It reports the client, both
processors, the infobase, starts a throwaway resident session, renders one real file,
and prints 1C's own error text if anything fails.

| Symptom | Cause |
|---|---|
| Previews open in the browser instead of the window | `tkinterweb` is missing, or installed into a different Python. The log prints the interpreter path at startup and the exact command to fix it. |
| Diagnostics: "the resident processor never reported ready", no `startup.log` | The `OnOpen` event is not wired (step 5 above), or the form is not the default form. |
| Diagnostics: `startup.log` stops partway | A later step raised; the trace names the error. |
| A 1C window sits on screen | Normal for the resident session — it is minimised automatically. *Tools → Keep 1C running between previews* turns it off. |
| Nothing happens when applying | Check the Files page: files showing an en dash have nothing to change. |
| A file "could not be parsed" | It is not a MOXCEL document, or is truncated. |

---

## How it works

`.mxl` is not zipped XML — it is a brace-nested text format that starts with the
magic bytes `MOXCEL`. Rows and cells are encoded as:

```
<row tuple> = {rowFormat} , rowNum , height , nCells , flag
<cell>      = {cellData}  , columnNumber
```

A row owns exactly `nCells` cells. The first `nCells - 1` carry an explicit column
number; the **last cell's number is elided**, and the run holding that last cell is
simultaneously the next row's tuple. That overlap is what makes the format
confusing, and it is why the parser is driven purely by the declared counts rather
than by guessing where a row ends.

Inserting a column therefore has to bump the explicit column numbers after the
insert point, increment `nCells`, and — when inserting at the very end — promote the
old implicit last cell to an explicit one. Deleting does the mirror. Renaming swaps
the caption *inside* the existing header cell, so nothing else can shift.

Registers are discovered by their title row (`Accumulation register "Name"`), then
the header row and the contiguous data rows of the same width beneath it.

---

## Repository layout

```
mxl_column_editor.py        the whole program: parser, edit engine, renderer, UI
build_windows_exe.bat       one-click PyInstaller build
render_check.py / .bat      standalone harness: renders files original-vs-modified
                            through 1C for side-by-side comparison
onec/
  MxlToHtml.epf             one-shot renderer (fallback)
  MxlToHtml.bsl             its form module
  MxlToHtmlService.epf      resident renderer (build it yourself, see above)
  MxlToHtmlService.bsl      its form module
  MxlRendererTemplate.dt    minimal service infobase
  README.md                 building and debugging the processors
```

The minimum to run the program is `mxl_column_editor.py` plus `onec/` with the two
`.epf` files and the `.dt`.

---

## Credits

The MXL-to-HTML rendering approach — handing the document to the platform and
exporting HTML — along with the `MxlToHtml.epf` data processor and the
`MxlRendererTemplate.dt` service infobase come from
[**@alexiosus**](https://github.com/alexiosus), and are used here with his kind
permission. The resident `MxlToHtmlService` variant is an extension of that same
approach.

---

## License

[MIT](LICENSE).

The bundled 1C artifacts (`MxlToHtml.epf`, `MxlRendererTemplate.dt`) originate with
[@alexiosus](https://github.com/alexiosus) and are redistributed here with his
permission.
