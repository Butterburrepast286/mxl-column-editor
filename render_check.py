#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
1C render check
===============

Renders .mxl files through 1C:Enterprise twice -- once exactly as they are on
disk, once with a set of column modifications applied to a temporary copy --
and drops both HTML files plus a machine-readable report into an output folder.
Both renders share a single 1C launch, so the platform starts up once per file.

The point is to make the platform's own rendering inspectable: run this, then
open (or hand over) the output folder and compare `<name>__original.html`
against `<name>__modified.html`.

Usage
-----
Double-click `render_check.bat`, or:

    python render_check.py                    # uses the settings below
    python render_check.py FILE.mxl [FILE...] # override the file list

Everything it needs is configured in the CONFIG block. It reuses the renderer,
parser and rule engine from mxl_column_editor.py sitting next to it, so there is
no second copy of the MXL logic to keep in sync.
"""

import io
import os
import sys
import json
import shutil
import datetime
import tempfile
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# --------------------------------------------------------------------------
#  CONFIG -- edit this block
# --------------------------------------------------------------------------

# Files to render. Absolute paths, or names relative to SOURCE_FOLDER.
FILES = [
    "000014156_17_TMPLT.mxl",
]

# Where those files live.
SOURCE_FOLDER = r"C:\Claude work folder\_mxl_backup_20260727_184510"

# Where to write the HTML and the report. Must be a folder Claude can read.
OUTPUT_FOLDER = r"C:\Claude work folder\render_check"

# The modifications to apply for the "modified" render.
# (register, anchor column, "after"/"before", new column, value)
MODIFICATIONS = [
    ("Inventory and expenses", "Posting content", "after",
     "(not used) Transfer to POS (RIM)", "No"),
    ("Landed costs", "Corr. warehouse", "after",
     "(not used) Transfer to RIM", "No"),
]

# Path to 1cv8c.exe. Leave as None to use the editor's saved setting or
# auto-detection.
ONEC_CLIENT = None

# --------------------------------------------------------------------------


def banner(text):
    print("\n" + "=" * 70)
    print(text)
    print("=" * 70)


def inspect(M, path, entry):
    """Record what the file looks like before anything is rendered."""
    runs = M.runs_of(M.load(path)[2])
    rows, _ = M.rows_of(runs)
    tables = M.discover(runs, rows)
    entry["registers"] = [
        {"name": t.name,
         "columns": len(t.columns),
         "header_row": rows[t.header_idx]["rownum"],
         "data_rows": len(t.data_idx)}
        for t in tables]
    entry["widest_before"] = max((r["ncells"] for r in rows), default=0)
    print("registers: %d, widest %d columns" % (len(tables), entry["widest_before"]))


def main():
    try:
        import mxl_column_editor as M
    except ImportError:
        print("ERROR: mxl_column_editor.py must sit next to this script.")
        return 2

    files = sys.argv[1:] or FILES
    resolved = []
    for name in files:
        path = name if os.path.isabs(name) else os.path.join(SOURCE_FOLDER, name)
        if os.path.isfile(path):
            resolved.append(path)
        else:
            print("SKIP (not found): %s" % path)
    if not resolved:
        print("Nothing to render.")
        return 1

    client = ONEC_CLIENT or M.load_settings().get("onec_client") or M.find_onec_client()
    if not client or not os.path.isfile(client):
        print("ERROR: 1C client not found. Set ONEC_CLIENT at the top of this script,")
        print("       or configure it once in the editor (Tools > 1C:Enterprise renderer).")
        return 3

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    rules = [M.Rule(reg, anchor, pos, new, val)
             for reg, anchor, pos, new, val in MODIFICATIONS]

    banner("1C render check")
    print("client : %s" % client)
    print("output : %s" % OUTPUT_FOLDER)
    print("files  : %d" % len(resolved))
    print("batch  : %s" % ("yes" if M.epf_supports_batch() else
                           "no (legacy data processor: one launch per document)"))
    for rule in rules:
        print("rule   : %s" % rule.describe())

    report = {
        "generated": datetime.datetime.now().isoformat(timespec="seconds"),
        "client": client,
        "batch_capable": M.epf_supports_batch(),
        "modifications": [r.describe() for r in rules],
        "files": [],
    }

    for path in resolved:
        name = os.path.basename(path)
        stem = os.path.splitext(name)[0]
        entry = {"file": name, "source": path, "renders": {}, "errors": {}}
        banner(name)

        try:
            inspect(M, path, entry)
        except Exception as exc:
            entry["errors"]["parse"] = str(exc)
            print("parse failed: %s" % exc)
            report["files"].append(entry)
            continue

        work = tempfile.mkdtemp(prefix="render-check-")
        try:
            staged = os.path.join(work, name)
            applied = M.write_modified_copy(path, rules, True, staged)
            entry["applied"] = [r.describe() for r in applied]
            rows2, _ = M.rows_of(M.runs_of(M.load(staged)[2]))
            entry["widest_after"] = max((r["ncells"] for r in rows2), default=0)
            print("applied %d modification(s); widest %d -> %d columns"
                  % (len(applied), entry["widest_before"], entry["widest_after"]))

            keep = os.path.join(OUTPUT_FOLDER, stem + "__modified.mxl")
            shutil.copy2(staged, keep)
            entry["staged_mxl"] = keep

            jobs = [("original", path,
                     os.path.join(OUTPUT_FOLDER, stem + "__original.html")),
                    ("modified", staged,
                     os.path.join(OUTPUT_FOLDER, stem + "__modified.html"))]
            try:
                rendered = M.render_mxl_batch(jobs, client,
                                              log=lambda m: print("   " + m))
                for side in sorted(rendered):
                    out = rendered[side]
                    entry["renders"][side] = out
                    print("   %-9s -> %s (%d bytes)"
                          % (side, os.path.basename(out), os.path.getsize(out)))
            except Exception as exc:
                entry["errors"]["render"] = str(exc)
                print("   FAILED: %s" % exc)
        except Exception:
            entry["errors"]["run"] = traceback.format_exc()
            print(traceback.format_exc())
        finally:
            shutil.rmtree(work, ignore_errors=True)

        report["files"].append(entry)

    path = os.path.join(OUTPUT_FOLDER, "render_report.json")
    with io.open(path, "w", encoding="utf-8") as fh:
        json.dump(report, fh, indent=2, ensure_ascii=False)

    banner("Done")
    print("Report: %s" % path)
    print("Open the output folder and compare the __original / __modified pairs.")
    return 0


if __name__ == "__main__":
    code = main()
    try:
        input("\nPress Enter to close...")
    except (EOFError, KeyboardInterrupt):
        pass
    sys.exit(code)
