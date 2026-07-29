# 1C renderer assets

Two external data processors live here. Both do the same conversion in server
context, because `SpreadsheetDocument.Read` is unavailable in the thin client:

```bsl
SpreadsheetDocument = New SpreadsheetDocument;
SpreadsheetDocument.Read(InputFileName);
SpreadsheetDocument.Write(OutputFileName, SpreadsheetDocumentFileType.HTML);
```

| File | Role |
|---|---|
| `MxlToHtml.epf` | One-shot: renders a job and exits. Launched with `/C job.json`. |
| `MxlToHtml.bsl` | Its form module. |
| `MxlToHtmlService.bsl` | Form module for the **resident** renderer (not yet built). |
| `MxlRendererTemplate.dt` | Minimal service infobase, restored once per user and platform version. |

`MxlToHtml.epf` and `MxlRendererTemplate.dt` come from
[@alexiosus](https://github.com/alexiosus) and are redistributed with his
permission. `MxlToHtmlService.bsl` extends the same approach to a resident
process.

## Building MxlToHtmlService.epf

Without it the editor starts a fresh 1C process for every preview: a few seconds
and a window flash each time. The resident processor is launched once and then
serves a queue, so later previews are near-instant.

1. Open the English Designer and create a new external data processor.
2. Add a **managed form** to it.
3. In the data processor's properties, set that form as the **default form**
   (`DefaultForm`). Without this the platform opens nothing and the module never
   runs.
4. Open the form, switch to its **Module** tab — the *form* module, not the data
   processor's object module — and replace the whole contents with
   `MxlToHtmlService.bsl`.
5. **Wire the OnOpen event.** This is the step that is easy to miss, and nothing
   works without it: a procedure merely *named* `OnOpen` is not the form's event
   handler. The form opens, no code runs, and there is no error anywhere.
   - switch to the form's **Form** tab
   - select the form's **root element** in the elements tree (the top node — the
     form itself, not a control)
   - open the properties palette with **F4**
   - in the **Events** group, click the magnifier beside **OnOpen**
   - 1C links to the existing `OnOpen` procedure; the field should now read
     `OnOpen`
6. Save as `MxlToHtmlService.epf` **in this folder**, next to the others.

### Checking the build

Run *Tools → 1C renderer diagnostics…* in the editor. It starts a throwaway
session and prints what happened, including the processor's own trace.

The module writes `startup.log` into the session folder as its very first act, so:

| Symptom | Meaning |
|---|---|
| no `startup.log` at all | the module never ran — almost always the OnOpen event is not wired (step 5), sometimes the default form is not set (step 3) |
| `startup.log` stops after "OnOpen reached" | a later line raised; the trace names it |
| `ready.json` present | the build is good |

`jobs\` and `done\` are created by the processor, not by the editor, so their
presence is itself evidence that the module ran.

That is the only trigger: the editor uses persistent mode when this exact
filename is present, and one launch per render when it is not. It is deliberately
a separate file from `MxlToHtml.epf`, because the two take different launch
parameters — a session folder rather than a job file — and handing a folder to
the one-shot processor would leave a stuck window on screen.

## The queue protocol

The processor receives a session folder as `LaunchParameter` and:

- writes `ready.json` once its idle loop is running, which is how the editor
  knows the build is good;
- polls `jobs\*.json`, each holding `{"items": [{name, inputPath, outputPath}]}`;
- writes the outcome to `done\<same name>` as `{"success": ..., "error": ...}`,
  then deletes the job;
- exits when a file named `stop` appears.

Job files are written to a staging name and moved into `jobs\` so the processor
never reads a half-written file. If `ready.json` does not appear within two
minutes the editor kills the process, logs it, and quietly uses one-shot mode for
the rest of the session.
