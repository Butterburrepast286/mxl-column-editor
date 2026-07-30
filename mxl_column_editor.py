#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MXL Column Editor
=================

Adds columns to register tables inside 1C:Enterprise spreadsheet documents
(*.mxl -- the "MOXCEL" tabular document format), in bulk, across a folder.

Typical use: a register in the configuration gained a new dimension/resource,
and every expected-result template (.mxl) in a test suite now needs that column
inserted next to an existing one, with a constant value in every data row.

Run it:

    python mxl_column_editor.py

No third-party packages required -- standard library + Tkinter only.

--------------------------------------------------------------------------
FILE FORMAT NOTES (how the parser works)
--------------------------------------------------------------------------
An .mxl file is:

    b'MOXCEL\\x00\\x08\\x00\\x01\\x00\\x0c\\x00'  header
    optional UTF-8 BOM
    one big brace-nested, comma-separated UTF-8 structure

Inside that structure the document body is a flat sequence of "runs", where a
run is a {...} group followed by zero or more bare numbers.  Rows and cells are
encoded like this:

    <row tuple>   = {rowFormat} , rowNum , height , nCells , flag
    <cell>        = {cellData}  , columnNumber

A row owns exactly nCells cells.  The first (nCells - 1) of them carry an
explicit column number; the LAST cell's column number is elided, and the run
holding that last cell is simultaneously the next row's row-tuple group.  That
overlap is what makes the format confusing to eyeball, and it is why this
parser is driven purely by the declared counts rather than by guessing where
a row ends.

Consequences for inserting a column, all handled below:
  * cells after the insert point need their explicit column numbers bumped;
  * the row's nCells must be incremented;
  * inserting at the very end promotes the old implicit last cell to an
    explicit one and makes the new cell implicit.
"""

import os
import re
import sys
import glob
import time
import atexit
import json
import shutil
import queue
import hashlib
import tempfile
import threading
import traceback
import datetime
import subprocess
import webbrowser

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


APP_TITLE = "MXL Column Editor"
MAGIC = b"MOXCEL"
BOM = b"\xef\xbb\xbf"


# ==========================================================================
#  Low level: load / save
# ==========================================================================

def load(path):
    """Return (header_bytes, bom_bytes, text)."""
    raw = open(path, "rb").read()
    if not raw.startswith(MAGIC):
        raise ValueError("not an MXL file (missing MOXCEL signature)")
    hdr, body = raw[:12], raw[12:]
    bom = b""
    if body.startswith(BOM):
        bom, body = body[:3], body[3:]
    return hdr, bom, body.decode("utf-8")


def save(path, hdr, bom, text):
    with open(path, "wb") as fh:
        fh.write(hdr + bom + text.encode("utf-8"))


# ==========================================================================
#  Tokenizer / structure
# ==========================================================================

def tokens(s):
    """Yield (kind, start, end, value); kind in {'{','}',',','atom','str'}."""
    out, i, n = [], 0, len(s)
    while i < n:
        c = s[i]
        if c in "\r\n \t":
            i += 1
            continue
        if c in "{},":
            out.append((c, i, i + 1, c))
            i += 1
            continue
        if c == '"':
            j = i + 1
            while j < n:
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        j += 2
                        continue
                    j += 1
                    break
                j += 1
            out.append(("str", i, j, s[i:j]))
            i = j
            continue
        j = i
        while j < n and s[j] not in '{},"\r\n':
            j += 1
        out.append(("atom", i, j, s[i:j].strip()))
        i = j
    return out


def runs_of(text):
    """Flatten the master group into runs: {'gs','ge','g','nums':[{'s','e','v'}]}."""
    tk_ = tokens(text)
    start = next(i for i, x in enumerate(tk_) if x[0] == "{")
    runs, cur = [], None
    depth, gstart, i = 0, None, start + 1
    while i < len(tk_):
        kind, s, e, v = tk_[i]
        if depth == 0:
            if kind == "{":
                gstart, depth = s, 1
            elif kind == "}":
                break
            elif kind == "atom" and cur is not None:
                cur["nums"].append({"s": s, "e": e, "v": v})
        else:
            if kind == "{":
                depth += 1
            elif kind == "}":
                depth -= 1
                if depth == 0:
                    cur = {"gs": gstart, "ge": e, "g": text[gstart:e], "nums": []}
                    runs.append(cur)
        i += 1
    return runs


_TXT = re.compile(r'"#","((?:[^"]|"")*)"')


def celltext(group):
    """Displayed text of a cell group ('' when the cell is empty)."""
    m = _TXT.findall(group)
    return m[-1].replace('""', '"') if m else ""


def cellstyle(group):
    """Style index of a cell group, e.g. 7 for '{16,7,...}'."""
    m = re.match(r"\{\s*\d+\s*,\s*(\d+)", group)
    return m.group(1) if m else "9"


def rows_of(runs):
    """
    Parse the row region using the declared counts only.

    Returns (rows, index_after_last_cell) where each row is
    {'tup', 'nums', 'rownum', 'ncells', 'cells':[{'i','col'}]}.
    The final cell of every row has col=None (implicit).
    """
    cand = [i for i, r in enumerate(runs) if len(r["nums"]) == 7]
    if not cand:
        raise ValueError("no document preamble marker found")
    start = cand[0]
    nrows = int(runs[start]["nums"][2]["v"])
    out, i = [], start
    while True:
        nums = runs[i]["nums"][-4:]
        ncells = int(nums[2]["v"])
        cells, k = [], i + 1
        for _ in range(ncells - 1):
            if k >= len(runs) or len(runs[k]["nums"]) != 1:
                raise ValueError("malformed cell near row %s" % nums[0]["v"])
            cells.append({"i": k, "col": int(runs[k]["nums"][0]["v"])})
            k += 1
        if k >= len(runs):
            raise ValueError("file truncated at row %s" % nums[0]["v"])
        cells.append({"i": k, "col": None})
        out.append({"tup": i, "nums": nums, "rownum": int(nums[0]["v"]),
                    "ncells": ncells, "cells": cells})
        i = k
        if len(out) >= nrows:
            break
    return out, i + 1


# ==========================================================================
#  Register discovery
# ==========================================================================

_REG = re.compile(r'^(?:Accumulation|Information|Accounting)\s+register\s+"(.+)"$')


class Table(object):
    """One register table: its title row, header row, columns and data rows."""

    def __init__(self, name, kind, title_idx, header_idx, columns, data_idx):
        self.name = name
        self.kind = kind
        self.title_idx = title_idx
        self.header_idx = header_idx
        self.columns = columns          # list of header captions, cell order
        self.data_idx = data_idx        # row indices of data rows

    @property
    def width(self):
        return len(self.columns)


def discover(runs, rows):
    """Find every register table in a parsed document."""
    tables = []
    for k, row in enumerate(rows):
        title = None
        for c in row["cells"]:
            t = celltext(runs[c["i"]]["g"]).strip()
            m = _REG.match(t)
            if m:
                title = m
                break
        if not title:
            continue
        header_idx = None
        for j in range(k + 1, min(k + 5, len(rows))):
            if rows[j]["ncells"] > 1:
                header_idx = j
                break
        if header_idx is None:
            continue
        hdr = rows[header_idx]
        cols = [celltext(runs[c["i"]]["g"]).strip() for c in hdr["cells"]]
        data = []
        j = header_idx + 1
        while j < len(rows) and rows[j]["ncells"] == hdr["ncells"]:
            data.append(j)
            j += 1
        tables.append(Table(title.group(1), title.group(0).split()[0],
                            k, header_idx, cols, data))
    return tables


# ==========================================================================
#  The edit itself
# ==========================================================================

class Rule(object):
    """
    One modification to one register.

    action "add"    -- insert `new_col` before/after `anchor`, filling data rows
                       with `value`.
    action "delete" -- remove the column named `anchor` and all of its cells.
    action "rename" -- change the caption of `anchor` to `new_col`; only the
                       header cell is touched, the data underneath is untouched.
    """

    def __init__(self, register, anchor, position="after", new_col="", value="",
                 action="add"):
        self.register = register
        self.anchor = anchor            # for "delete" this IS the doomed column
        self.position = position        # 'after' | 'before'
        self.new_col = new_col
        self.value = value
        self.action = action            # 'add' | 'delete'

    @property
    def is_delete(self):
        return self.action == "delete"

    @property
    def is_rename(self):
        return self.action == "rename"

    def describe(self):
        if self.is_delete:
            return 'register "%s": delete column "%s"' % (self.register, self.anchor)
        if self.is_rename:
            return ('register "%s": rename "%s" to "%s"'
                    % (self.register, self.anchor, self.new_col))
        return ('register "%s": add "%s" %s "%s" = "%s"'
                % (self.register, self.new_col, self.position, self.anchor, self.value))


def _mk_cell(style, text):
    if text == "":
        return "{16,%s,\r\n{1,0},0}" % style
    return '{16,%s,\r\n{1,1,\r\n{"#","%s"}\r\n},0}' % (style, text.replace('"', '""'))


def _target_rows(tab):
    return [tab.header_idx] + list(tab.data_idx)


def _plan_add(runs, rows, tab, rule):
    """Patches that insert one column into one register table."""
    patches = []
    p = tab.columns.index(rule.anchor)
    q = p + 1 if rule.position == "after" else p       # 0-based insert slot

    for n, ridx in enumerate(_target_rows(tab)):
        row = rows[ridx]
        cells = row["cells"]
        style = cellstyle(runs[cells[p]["i"]]["g"])
        body = _mk_cell(style, rule.new_col if n == 0 else rule.value)

        # bump explicit column numbers at or after the insert slot
        for c in cells[q:]:
            if c["col"] is not None:
                tok = runs[c["i"]]["nums"][0]
                patches.append((tok["s"], tok["e"], str(c["col"] + 1)))

        if q < len(cells):
            gs = runs[cells[q]["i"]]["gs"]
            patches.append((gs, gs, "%s,%d,\r\n" % (body, q + 1)))
        else:
            # appending past the implicit last cell: promote it, the new cell
            # becomes the implicit one
            ge = runs[cells[-1]["i"]]["ge"]
            patches.append((ge, ge, ",%d,\r\n%s" % (len(cells), body)))

        nc = row["nums"][2]
        patches.append((nc["s"], nc["e"], str(row["ncells"] + 1)))

    return patches, ("ok", "inserted at position %d, %d data row(s)"
                           % (q + 1, len(tab.data_idx)))


def _plan_delete(runs, rows, tab, rule):
    """
    Patches that remove one column from one register table.

    The mirror of _plan_add, including the same wrinkle: the last cell of a row
    carries no column number, so deleting it means the cell before it must give
    up its own number and become the implicit one.
    """
    patches = []
    p = tab.columns.index(rule.anchor)

    for ridx in _target_rows(tab):
        row = rows[ridx]
        cells = row["cells"]
        if row["ncells"] < 2:
            continue                       # nothing sensible to remove
        if p >= len(cells):
            continue                       # ragged row: leave it alone

        if p < len(cells) - 1:
            # drop "{cell},COL,\r\n" up to the next cell's group
            gs = runs[cells[p]["i"]]["gs"]
            nxt = runs[cells[p + 1]["i"]]["gs"]
            patches.append((gs, nxt, ""))
        else:
            # dropping the implicit last cell: also drop the previous cell's
            # column number, which makes it implicit in turn
            prev_ge = runs[cells[p - 1]["i"]]["ge"]
            ge = runs[cells[p]["i"]]["ge"]
            patches.append((prev_ge, ge, ""))

        # pull the surviving explicit column numbers back by one
        for c in cells[p + 1:]:
            if c["col"] is not None:
                tok = runs[c["i"]]["nums"][0]
                patches.append((tok["s"], tok["e"], str(c["col"] - 1)))

        nc = row["nums"][2]
        patches.append((nc["s"], nc["e"], str(row["ncells"] - 1)))

    return patches, ("ok", "removed from position %d, %d data row(s)"
                           % (p + 1, len(tab.data_idx)))


def _plan_rename(runs, rows, tab, rule):
    """
    Patches that change one column's caption.

    Only the header cell moves; widths, column numbers and every data row stay
    exactly as they were. The caption is swapped inside the existing cell rather
    than rebuilding it, so whatever else the cell carries survives untouched.
    """
    p = tab.columns.index(rule.anchor)
    header = rows[tab.header_idx]
    run = runs[header["cells"][p]["i"]]
    escaped = rule.new_col.replace('"', '""')

    matches = list(_TXT.finditer(run["g"]))
    if matches:
        last = matches[-1]
        patch = (run["gs"] + last.start(1), run["gs"] + last.end(1), escaped)
    else:
        # the header cell was empty: give it text, keeping its style
        patch = (run["gs"], run["ge"], _mk_cell(cellstyle(run["g"]), rule.new_col))

    return [patch], ("ok", "renamed at position %d, %d data row(s) untouched"
                           % (p + 1, len(tab.data_idx)))


def plan_rule(text, rule, skip_existing=True):
    """
    Work out the byte patches this rule implies for one document.

    Returns (patches, outcomes) where patches is a list of (start, end, replacement)
    and outcomes is a list of (status, detail) pairs, one per matching table.
    """
    runs = runs_of(text)
    rows, _ = rows_of(runs)
    tables = [t for t in discover(runs, rows) if t.name == rule.register]
    if not tables:
        return [], [("skip", "register not present")]

    patches, outcomes = [], []
    for tab in tables:
        if rule.is_rename:
            if rule.anchor not in tab.columns:
                outcomes.append(("skip", "already renamed"
                                 if rule.new_col in tab.columns else
                                 'column "%s" not in this register' % rule.anchor))
                continue
            if tab.columns.count(rule.anchor) > 1:
                outcomes.append(("skip", 'column "%s" appears more than once; '
                                         'refusing to guess' % rule.anchor))
                continue
            if rule.new_col in tab.columns:
                outcomes.append(("skip", 'a column named "%s" already exists here'
                                         % rule.new_col))
                continue
            made, outcome = _plan_rename(runs, rows, tab, rule)
        elif rule.is_delete:
            if rule.anchor not in tab.columns:
                outcomes.append(("skip", 'column "%s" not in this register' % rule.anchor))
                continue
            if tab.columns.count(rule.anchor) > 1:
                outcomes.append(("skip", 'column "%s" appears more than once; '
                                         'refusing to guess' % rule.anchor))
                continue
            made, outcome = _plan_delete(runs, rows, tab, rule)
        else:
            if skip_existing and rule.new_col in tab.columns:
                outcomes.append(("skip", "column already present"))
                continue
            if rule.anchor not in tab.columns:
                outcomes.append(("skip", 'anchor column "%s" not in this register'
                                         % rule.anchor))
                continue
            made, outcome = _plan_add(runs, rows, tab, rule)
        patches.extend(made)
        outcomes.append(outcome)
    return patches, outcomes


def apply_patches(text, patches):
    """Apply (start, end, replacement) triples right-to-left."""
    for s, e, new in sorted(patches, key=lambda x: (-x[0], -x[1])):
        text = text[:s] + new + text[e:]
    return text


def process_file(path, rules, skip_existing=True, write=False):
    """Plan (and optionally write) every rule for one file. Returns list of (rule, outcomes)."""
    hdr, bom, text = load(path)
    report = []
    for rule in rules:
        patches, outcomes = plan_rule(text, rule, skip_existing)
        report.append((rule, outcomes))
        if patches:
            text = apply_patches(text, patches)
    if write and any(o[0] == "ok" for _, outs in report for o in outs):
        save(path, hdr, bom, text)
    return report


# ==========================================================================
#  Verification
# ==========================================================================

def _fingerprint(text):
    runs = runs_of(text)
    rows, end = rows_of(runs)
    grid = [(r["rownum"], r["ncells"],
             tuple((c["col"], celltext(runs[c["i"]]["g"])) for c in r["cells"]))
            for r in rows]
    sparse = {r["rownum"] for r in rows
              if [c["col"] for c in r["cells"] if c["col"] is not None]
              != list(range(1, r["ncells"]))}
    pre_end = runs[8]["ge"] if len(runs) > 8 else 0
    return grid, sparse, text[runs[end]["gs"]:] if end < len(runs) else "", text[:pre_end]


def _is_subsequence(small, big):
    """True when every item of `small` appears in `big`, in order."""
    it = iter(big)
    return all(any(x == y for y in it) for x in small)


def verify_file(new_path, old_path, rules):
    """Compare a written file against its backup. Returns list of problem strings."""
    problems = []
    try:
        new_txt = load(new_path)[2]
        old_txt = load(old_path)[2]
    except Exception as exc:
        return ["unreadable: %s" % exc]

    try:
        g_new, sp_new, tail_new, pre_new = _fingerprint(new_txt)
        g_old, sp_old, tail_old, pre_old = _fingerprint(old_txt)
    except Exception as exc:
        return ["does not re-parse after write: %s" % exc]

    adds = [r for r in rules if r.action == "add"]
    dels = [r for r in rules if r.is_delete]
    renames = [r for r in rules if r.is_rename]
    lo, hi = -len(dels), len(adds)

    if len(g_new) != len(g_old):
        problems.append("row count changed %d -> %d" % (len(g_old), len(g_new)))
    else:
        for a, b in zip(g_old, g_new):
            if a == b:
                continue
            delta = b[1] - a[1]
            if not lo <= delta <= hi:
                problems.append("row %s width delta %+d (expected %+d..%+d)"
                                % (a[0], delta, lo, hi))
            elif delta == 0 and not (renames or (adds and dels)):
                problems.append("row %s changed but its width did not" % a[0])
            if renames:
                continue          # a rename legitimately rewrites a header caption
            row_old = [t for _, t in a[2]]
            row_new = [t for _, t in b[2]]
            if delta > 0 and not dels:
                if not _is_subsequence(row_old, row_new):
                    problems.append("row %s: existing cell contents were altered, "
                                    "not just extended" % a[0])
            elif delta < 0 and not adds:
                if not _is_subsequence(row_new, row_old):
                    problems.append("row %s: surviving cell contents were altered, "
                                    "not just shortened" % a[0])
    if sp_new != sp_old:
        problems.append("sparse-column rows changed")
    if tail_new != tail_old:
        problems.append("document trailer changed")
    if pre_new != pre_old:
        problems.append("document preamble changed")

    runs = runs_of(new_txt)
    rows, _ = rows_of(runs)
    tables = discover(runs, rows)
    old_runs = runs_of(old_txt)
    old_rows, _ = rows_of(old_runs)
    old_tables = discover(old_runs, old_rows)

    for rule in rules:
        for index, tab in enumerate(tables):
            if tab.name != rule.register:
                continue
            if rule.is_rename:
                was = old_tables[index].columns if index < len(old_tables) else []
                if rule.anchor not in was:
                    continue                      # nothing to rename in this table
                if rule.anchor in tab.columns:
                    problems.append('%s: column "%s" was not renamed'
                                    % (tab.name, rule.anchor))
                elif tab.columns.count(rule.new_col) != 1:
                    problems.append('%s: "%s" appears %d times after the rename'
                                    % (tab.name, rule.new_col,
                                       tab.columns.count(rule.new_col)))
                elif len(tab.columns) != len(was):
                    problems.append('%s: rename changed the column count %d -> %d'
                                    % (tab.name, len(was), len(tab.columns)))
                elif tab.columns.index(rule.new_col) != was.index(rule.anchor):
                    problems.append('%s: renamed column moved from position %d to %d'
                                    % (tab.name, was.index(rule.anchor) + 1,
                                       tab.columns.index(rule.new_col) + 1))
                continue
            if rule.is_delete:
                if rule.anchor in tab.columns:
                    problems.append('%s: column "%s" is still present after a delete'
                                    % (tab.name, rule.anchor))
                continue
            if rule.new_col not in tab.columns:
                problems.append('%s: column "%s" missing' % (tab.name, rule.new_col))
                continue
            if tab.columns.count(rule.new_col) != 1:
                problems.append('%s: column "%s" appears %d times'
                                % (tab.name, rule.new_col, tab.columns.count(rule.new_col)))
                continue
            if rule.anchor in tab.columns:
                i_new, i_anc = tab.columns.index(rule.new_col), tab.columns.index(rule.anchor)
                want = i_anc + 1 if rule.position == "after" else i_anc - 1
                if i_new != want:
                    problems.append('%s: "%s" at position %d, expected %d'
                                    % (tab.name, rule.new_col, i_new + 1, want + 1))
            i = tab.columns.index(rule.new_col)
            for ridx in tab.data_idx:
                got = celltext(runs[rows[ridx]["cells"][i]["i"]]["g"]).strip()
                if got != rule.value:
                    problems.append('%s: row %d holds "%s", expected "%s"'
                                    % (tab.name, rows[ridx]["rownum"], got, rule.value))
                    break
    return problems


# ==========================================================================
#  Folder scan
# ==========================================================================

def scan_folder(folder):
    """Return (per_file_tables, registers, errors)."""
    per_file, registers, errors = {}, {}, []
    for path in sorted(glob.glob(os.path.join(folder, "*.mxl"))):
        name = os.path.basename(path)
        try:
            runs = runs_of(load(path)[2])
            rows, _ = rows_of(runs)
            tabs = discover(runs, rows)
        except Exception as exc:
            errors.append((name, str(exc)))
            continue
        per_file[name] = tabs
        for t in tabs:
            entry = registers.setdefault(t.name, {"files": set(), "columns": []})
            entry["files"].add(name)
            for col in t.columns:
                if col and col not in entry["columns"]:
                    entry["columns"].append(col)
    return per_file, registers, errors


# ==========================================================================
#  Preview rendering
# ==========================================================================

def modified_preview(path, rules, skip_existing=True):
    """
    Apply the rules to a document IN MEMORY and return everything needed to
    draw it: (runs, rows, highlight, applied).

    `highlight` maps a row index to the set of cell positions that this run of
    rules actually introduced, so the preview can tint only genuinely new cells
    (a register that already had the column is left unhighlighted).
    """
    text = load(path)[2]
    runs0 = runs_of(text)
    rows0, _ = rows_of(runs0)
    before = [t.columns for t in discover(runs0, rows0)]

    applied = []
    for rule in rules:
        patches, _outcomes = plan_rule(text, rule, skip_existing)
        if patches:
            text = apply_patches(text, patches)
            applied.append(rule)

    runs = runs_of(text)
    rows, _ = rows_of(runs)
    tables = discover(runs, rows)

    highlight = {}
    for i, tab in enumerate(tables):
        was = before[i] if i < len(before) else []
        for rule in applied:
            if tab.name != rule.register:
                continue
            if rule.new_col in tab.columns and rule.new_col not in was:
                pos = tab.columns.index(rule.new_col)
                for ridx in [tab.header_idx] + list(tab.data_idx):
                    highlight.setdefault(ridx, set()).add(pos)
    return runs, rows, highlight, applied


def grid_lines(runs, rows, highlight, maxw=22):
    """
    Lay the document out as a fixed-width grid.

    Returns (lines, spans) where spans is a list of
    (line_number, start_char, end_char, tag).
    """
    grid = []
    for r in rows:
        grid.append([celltext(runs[c["i"]]["g"]).replace("\r", " ").replace("\n", " ")
                     for c in r["cells"]])
    ncols = max((len(g) for g in grid), default=0)
    widths = [3] * ncols
    for vals in grid:
        for j, v in enumerate(vals):
            widths[j] = max(widths[j], min(len(v), maxw))

    def cut(v, w):
        return v if len(v) <= w else v[:w - 1] + "~"

    lines, spans = [], []
    head = "  row | " + " | ".join(str(j + 1).ljust(widths[j]) for j in range(ncols))
    lines.append(head)
    lines.append("-" * len(head))

    for ridx, r in enumerate(rows):
        vals = grid[ridx]
        hot = highlight.get(ridx, set())
        buf = "%5s | " % r["rownum"]
        for j in range(len(vals)):
            start = len(buf)
            buf += cut(vals[j], widths[j]).ljust(widths[j])
            if j in hot:
                spans.append((len(lines), start, len(buf), "new"))
            if j < len(vals) - 1:
                buf += " | "
        lines.append(buf.rstrip())
    return lines, spans


def html_widget_status():
    """
    (HtmlFrame, message) for the embedded HTML view.

    Imported lazily and behind BaseException: tkinterweb calls sys.exit() when
    its Tkhtml binary is missing, which would otherwise take the whole editor
    down just because someone opened a preview tab.
    """
    try:
        import tkinterweb
        from tkinterweb import HtmlFrame
        return HtmlFrame, "tkinterweb %s" % getattr(tkinterweb, "__version__", "?")
    except BaseException as exc:
        return None, "%s: %s" % (type(exc).__name__, exc or "no detail")


def html_widget_class():
    return html_widget_status()[0]


class PreviewWindow(tk.Toplevel):
    """Full-document grid preview with inserted columns tinted."""

    def __init__(self, master, path, rules, skip_existing=True, on_exact=None):
        tk.Toplevel.__init__(self, master, bg=BG)
        self.path = path
        self.rules = rules
        self.skip_existing = skip_existing
        self.on_exact = on_exact
        self.mode = tk.StringVar(value="modified" if rules else "original")
        self.rendering = False
        self._render_token = 0
        self._pending_mode = None
        self.title("Preview - " + os.path.basename(path))
        self.geometry("1150x700")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        bar = ttk.Frame(self, padding=(8, 8, 8, 4))
        bar.grid(row=0, column=0, sticky="ew")
        if rules:
            ttk.Radiobutton(bar, text="With modifications", value="modified",
                            variable=self.mode, style="TCheckbutton",
                            command=self.on_mode_change).pack(side="left")
            ttk.Radiobutton(bar, text="Original", value="original",
                            variable=self.mode, style="TCheckbutton",
                            command=self.on_mode_change).pack(side="left", padx=(10, 16))
        self.info = ttk.Label(bar, text="")
        self.info.pack(side="left")
        ttk.Button(bar, text="Close", command=self.destroy).pack(side="right")

        self.nb = ttk.Notebook(self)
        self.nb.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self._build_rendered_tab()

        wrap = ttk.Frame(self.nb, padding=6)
        self.nb.add(wrap, text="  Grid  ")
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        self.text = tk.Text(wrap, wrap="none", font="TkFixedFont", bd=0,
                            highlightthickness=0, background=LOGBG,
                            foreground=TEXT, insertbackground=TEXT,
                            padx=8, pady=6)
        self.text.grid(row=0, column=0, sticky="nsew")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.text.yview)
        vs.grid(row=0, column=1, sticky="ns")
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=self.text.xview)
        hs.grid(row=1, column=0, sticky="ew")
        self.text.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self.text.tag_configure("new", background=HIGHLIGHT, foreground=TEXT)
        self.text.tag_configure("head", foreground=MUTED)

        self.render()

    # ------------------------------------------------- rendered (1C) tab

    def _build_rendered_tab(self):
        page = ttk.Frame(self.nb, padding=6)
        self.nb.add(page, text="  1C preview  ")
        page.columnconfigure(0, weight=1)
        page.rowconfigure(1, weight=1)

        bar = ttk.Frame(page)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Button(bar, text="Open in browser",
                   command=self._exact).pack(side="left")
        self.html_status = ttk.Label(bar, text="", style="Muted.TLabel")
        self.html_status.pack(side="left", padx=10)

        self.html_host = ttk.Frame(page)
        self.html_host.grid(row=1, column=0, sticky="nsew")
        self.html_host.columnconfigure(0, weight=1)
        self.html_host.rowconfigure(0, weight=1)
        self.html_view = None
        self.html_rendered_for = None
        self._html_placeholder()

    def _html_placeholder(self, message=None):
        for child in self.html_host.winfo_children():
            child.destroy()
        self.html_view = None
        if message is None:
            widget, detail = html_widget_status()
            if widget is None:
                if getattr(sys, "frozen", False):
                    message = ("This build does not include the embedded HTML view.\n\n"
                               "It has to be bundled when the .exe is built: install\n"
                               "tkinterweb first, then run build_windows_exe.bat again.\n\n"
                               "Import result:\n    %s\n\n"
                               "'Open in browser' still shows the same render." % detail)
                else:
                    message = ("An embedded HTML view needs the tkinterweb package.\n\n"
                               "Install it into THIS Python:\n\n"
                               "    \"%s\" -m pip install tkinterweb\n\n"
                               "Import result:\n    %s\n\n"
                               "Without it, 'Open in browser' still shows the same render."
                               % (sys.executable, detail))
            else:
                message = ("Handing this document to 1C:Enterprise...\n\n"
                           "The platform takes a few seconds to start.\n\n"
                           "Using %s" % detail)
        holder = tk.Frame(self.html_host, bg=PANEL, highlightbackground=BORDER,
                          highlightthickness=1, bd=0)
        holder.grid(row=0, column=0, sticky="nsew")
        tk.Label(holder, text=message, bg=PANEL, fg=MUTED, justify="left",
                 padx=18, pady=16).pack(anchor="nw")

    def show_rendered(self):
        """Bring the 1C preview tab forward and start rendering."""
        try:
            self.nb.select(0)
        except Exception:
            pass
        self._render_html()

    def _render_html(self):
        app = self.master
        if not hasattr(app, "render_html_file"):
            self._html_placeholder("This preview cannot reach the renderer.")
            return
        if html_widget_class() is None:
            self._html_placeholder()
            return

        mode = self.mode.get()
        if self.rendering:
            # a render is already in flight; queue this one instead of racing 1C
            self._pending_mode = mode
            self.html_status.configure(text="Rendering... (queued the new view)")
            return

        self.rendering = True
        self._render_token += 1
        token = self._render_token
        self._html_placeholder("Rendering with 1C:Enterprise...\n\n"
                               "The platform takes a few seconds to start.")
        self.html_status.configure(text="Rendering with 1C:Enterprise...")

        def work():
            try:
                out = app.render_html_file(self.path, mode == "modified")
            except Exception as exc:
                self.after(0, lambda: self._render_failed(exc, token))
                return
            self.after(0, lambda: self._render_done(out, mode, token))

        threading.Thread(target=work, daemon=True).start()

    def reclaim_focus(self):
        """Come back to the front after 1C has had the screen."""
        try:
            self.deiconify()
            self.lift()
            self.attributes("-topmost", True)
            self.after(200, lambda: self.attributes("-topmost", False))
            self.focus_force()
        except Exception:
            pass

    def _render_failed(self, exc, token=None):
        if token is not None and token != self._render_token:
            return
        self.rendering = False
        self._pending_mode = None
        self.reclaim_focus()
        self.html_status.configure(text="")
        self._html_placeholder("1C could not render this document:\n\n%s" % exc)

    def _render_done(self, path, mode, token=None):
        if token is not None and token != self._render_token:
            return
        self.rendering = False
        self.reclaim_focus()
        widget = html_widget_class()
        for child in self.html_host.winfo_children():
            child.destroy()
        try:
            view = widget(self.html_host, messages_enabled=False)
        except TypeError:
            view = widget(self.html_host)
        except BaseException as exc:
            self._render_failed(exc)
            return
        view.grid(row=0, column=0, sticky="nsew")
        try:
            view.load_file(path, force=True)
        except BaseException as exc:
            self._render_failed(exc)
            return
        self.html_view = view
        self.html_rendered_for = mode
        self.html_status.configure(
            text="Rendered %s" % ("with modifications" if mode == "modified"
                                  else "from the file on disk"))
        pending, self._pending_mode = self._pending_mode, None
        if pending is not None and pending != mode:
            self._render_html()          # the user switched while 1C was busy

    def _exact(self):
        if self.on_exact:
            self.on_exact(self.path, self.mode.get() == "modified")

    def on_mode_change(self):
        """Original / With modifications: redraw the grid and re-render through 1C."""
        self.render()
        if html_widget_class() is not None:
            self._render_html()

    def render(self):
        """(Re)draw the grid for whichever mode is selected."""
        rules = self.rules if self.mode.get() == "modified" else []
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        try:
            runs, rows, highlight, applied = modified_preview(
                self.path, rules, self.skip_existing)
        except Exception as exc:
            self.text.insert("1.0", "Could not read this file:\n\n%s" % exc)
            self.text.configure(state="disabled")
            return

        lines, spans = grid_lines(runs, rows, highlight)
        self.text.insert("1.0", "\n".join(lines))
        self.text.tag_add("head", "1.0", "3.0")
        for ln, a, b, tag in spans:
            self.text.tag_add(tag, "%d.%d" % (ln + 1, a), "%d.%d" % (ln + 1, b))
        self.text.configure(state="disabled")

        ncells = max((r["ncells"] for r in rows), default=0)
        n_new = sum(len(v) for v in highlight.values())
        if self.mode.get() == "original":
            note = "the file exactly as it is on disk"
        elif applied:
            note = ("%d modification(s) applied in this preview, %d cell(s) added "
                    "(shown highlighted)" % (len(applied), n_new))
        elif self.rules:
            note = "no modification applies to this file - showing it unchanged"
        else:
            note = "no modifications defined - showing the file as it is"
        self.info.configure(text="%d rows, widest register %d columns.  %s"
                                 % (len(rows), ncells, note))


# ==========================================================================
#  Exact preview via the 1C:Enterprise platform
# ==========================================================================
#
#  Approach borrowed from the mxl_merge tool: rather than reimplementing MXL
#  layout, hand the file to the platform itself.  The bundled external data
#  processor onec/MxlToHtml.epf does, in server context:
#
#       SpreadsheetDocument = New SpreadsheetDocument;
#       SpreadsheetDocument.Read(InputFileName);
#       SpreadsheetDocument.Write(OutputFileName, SpreadsheetDocumentFileType.HTML);
#
#  SpreadsheetDocument.Read is unavailable in the thin client, so the call needs
#  an infobase.  onec/MxlRendererTemplate.dt is a minimal service infobase that
#  gets created and restored once per user and platform version, then reused.
#
#  The processor is driven by a JSON job file and reports back through a status
#  file, which is what makes an interactive platform run usable as a subprocess.

ONEC_USER = "KOTStartupService"
ONEC_PASSWORD = ""
ONEC_TIMEOUT = 180
INFOBASE_MARKER = "1Cv8.1CD"
ONEC_STATE_FILE = ".mxl-column-editor-renderer.json"


class OneCError(RuntimeError):
    """The platform renderer could not produce HTML."""


def app_dir():
    """
    The folder the program lives in.

    Frozen by PyInstaller this is the folder holding the .exe, NOT the temporary
    extraction directory, so an `onec\\` folder placed beside the .exe wins over
    whatever was bundled at build time. That matters because the .epf files get
    rebuilt in Designer from time to time and nobody wants to rebuild the .exe
    just for that.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def bundled_dir():
    """Where PyInstaller unpacked bundled data, or the source folder."""
    return getattr(sys, "_MEIPASS", app_dir())


def onec_file(name):
    """An asset from onec\\, preferring one shipped beside the program."""
    beside = os.path.join(app_dir(), "onec", name)
    if os.path.isfile(beside):
        return beside
    bundled = os.path.join(bundled_dir(), "onec", name)
    return bundled if os.path.isfile(bundled) else beside


def onec_assets():
    """(epf, dt) paths for the one-shot renderer and the infobase template."""
    return onec_file("MxlToHtml.epf"), onec_file("MxlRendererTemplate.dt")


def _settings_path():
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "MxlColumnEditor", "settings.json")


def load_settings():
    try:
        with open(_settings_path(), encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(data):
    path = _settings_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(data, fh, indent=2)
        return True
    except OSError:
        return False


def _version_key(text):
    return [int(x) for x in re.findall(r"\d+", text)] or [0]


def find_onec_client():
    """Best-effort search for 1cv8c.exe (thin client), newest version first."""
    env = os.environ.get("MXL_ONEC_CLIENT", "").strip()
    if env and os.path.isfile(env):
        return env
    roots = [os.environ.get("PROGRAMFILES", r"C:\Program Files"),
             os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)"),
             r"C:\Program Files", r"C:\Program Files (x86)"]
    found = []
    for root in roots:
        if not root:
            continue
        for exe in ("1cv8c.exe", "1cv8.exe"):
            found.extend(glob.glob(os.path.join(root, "1cv8", "*", "bin", exe)))
    if not found:
        return None
    found.sort(key=lambda pth: (_version_key(os.path.basename(os.path.dirname(os.path.dirname(pth)))),
                                os.path.basename(pth) == "1cv8c.exe"))
    return found[-1]


def designer_exe(client_exe):
    """1cv8.exe sitting next to the configured client; needed for CREATEINFOBASE."""
    folder, name = os.path.split(client_exe)
    designer = os.path.join(folder, "1cv8.exe" if name.lower().endswith(".exe") else "1cv8")
    if not os.path.isfile(designer):
        raise OneCError("1cv8.exe (Designer) was not found next to the client:\n%s" % designer)
    return designer


def _platform_version(client_exe):
    # split on both separators so a Windows path is read correctly anywhere
    parts = [x for x in re.split(r"[\\/]+", str(client_exe)) if x]
    name = ""
    if len(parts) >= 2:
        name = parts[-3] if len(parts) >= 3 and parts[-2].lower() == "bin" else parts[-2]
    return re.sub(r"[^A-Za-z0-9._-]+", "-", name).strip("-.") or "default"


def default_infobase(client_exe):
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "MxlColumnEditor", "renderer",
                        _platform_version(client_exe), "ib")


def _file_conn(infobase):
    value = str(infobase).replace('"', '""')
    if re.search(r'[\s;"]', str(infobase)):
        value = '"%s"' % value
    return "File=%s;" % value


def _has_marker(infobase):
    try:
        return any(n.lower() == INFOBASE_MARKER.lower() for n in os.listdir(infobase))
    except OSError:
        return False


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


ONEC_BACKGROUND = True          # start the platform minimised and unfocused
ONEC_PERSISTENT = True          # keep one 1C process alive for the session


def quiet_windows(pid):
    """
    Minimise any top-level window the given process owns (Windows only).

    STARTUPINFO governs only the first ShowWindow call, and the platform shows
    its main window itself afterwards. For the one-shot renderer that window is
    gone in a moment; a resident session would otherwise sit on screen for the
    rest of the day.
    """
    if os.name != "nt" or not ONEC_BACKGROUND:
        return 0
    try:
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        found = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def visit(hwnd, _lparam):
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if owner.value == pid and user32.IsWindowVisible(hwnd):
                found.append(hwnd)
            return True

        user32.EnumWindows(visit, 0)
        for hwnd in found:
            user32.ShowWindow(hwnd, 6)          # SW_MINIMIZE
        return len(found)
    except Exception:
        return 0


def _startupinfo():
    """
    On Windows, ask the launched process to come up minimised and NOT activated.

    1cv8c.exe is a GUI application: without this it grabs the foreground and
    pushes the preview window behind it, and nothing hands focus back when it
    exits. SW_SHOWMINNOACTIVE is used rather than SW_HIDE so that anything the
    platform genuinely needs to show - a licence prompt, say - is still
    reachable from the taskbar instead of being invisible.
    """
    if os.name != "nt" or not ONEC_BACKGROUND:
        return None
    try:
        info = subprocess.STARTUPINFO()
        info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        info.wShowWindow = 7            # SW_SHOWMINNOACTIVE
        return info
    except Exception:
        return None


def _run(command, timeout, what):
    try:
        return subprocess.run(command, capture_output=True, text=True,
                              timeout=timeout, check=False,
                              startupinfo=_startupinfo())
    except subprocess.TimeoutExpired:
        raise OneCError("Timed out while %s." % what)
    except OSError as exc:
        raise OneCError("Could not start 1C while %s:\n%s" % (what, exc))


def ensure_infobase(client_exe, infobase, log=None):
    """Create the service infobase from the bundled template, once."""
    _epf, template = onec_assets()
    if not os.path.isfile(template):
        raise OneCError("Renderer infobase template is missing:\n%s" % template)
    digest = _sha256(template)
    state = os.path.join(os.path.dirname(infobase),
                         os.path.basename(infobase) + ONEC_STATE_FILE)

    if _has_marker(infobase):
        try:
            with open(state, encoding="utf-8") as fh:
                if json.load(fh).get("template") == digest:
                    return infobase
        except (OSError, ValueError):
            return infobase          # created by hand or by an older build: leave it alone
        shutil.rmtree(infobase, ignore_errors=True)

    if os.path.isdir(infobase) and os.listdir(infobase):
        raise OneCError("Renderer infobase folder is not empty and holds no %s:\n%s"
                        % (INFOBASE_MARKER, infobase))

    if log:
        log("Preparing the 1C service infobase (first run only)...")
    designer = designer_exe(client_exe)
    os.makedirs(infobase, exist_ok=True)
    work = tempfile.mkdtemp(prefix="mxl-onec-")
    try:
        out = os.path.join(work, "create.log")
        done = _run([designer, "CREATEINFOBASE", _file_conn(infobase), "/Out", out],
                    ONEC_TIMEOUT, "creating the renderer infobase")
        if done.returncode != 0 or not _has_marker(infobase):
            raise OneCError("1C could not create the renderer infobase (exit %s):\n%s"
                            % (done.returncode, _text_of(out) or done.stderr or done.stdout))

        out = os.path.join(work, "restore.log")
        done = _run([designer, "DESIGNER", "/DisableStartupDialogs", "/DisableStartupMessages",
                     "/IBConnectionString", _file_conn(infobase),
                     "/RestoreIB", template, "/Out", out],
                    ONEC_TIMEOUT, "restoring the renderer infobase")
        if done.returncode != 0:
            raise OneCError("1C could not restore the renderer infobase (exit %s):\n%s"
                            % (done.returncode, _text_of(out) or done.stderr or done.stdout))
    finally:
        shutil.rmtree(work, ignore_errors=True)

    try:
        os.makedirs(os.path.dirname(state), exist_ok=True)
        with open(state, "w", encoding="utf-8") as fh:
            json.dump({"template": digest}, fh)
    except OSError:
        pass
    return infobase


def _text_of(path):
    try:
        with open(path, encoding="utf-8-sig", errors="replace") as fh:
            return fh.read().strip()
    except OSError:
        return ""


# The data processor shipped originally could only handle one document per
# launch. The batch-capable rebuild accepts an "items" array; anything whose
# digest differs from the legacy build is assumed to be the newer one.
LEGACY_SINGLE_RENDER_EPF_SHA256 = (
    "aa894caf035962974c1834fa8ae9e123a0f3f89182ce53dfc9ad8d1eae0a1e56"
)


def epf_supports_batch(epf=None):
    try:
        return _sha256(epf or onec_assets()[0]) != LEGACY_SINGLE_RENDER_EPF_SHA256
    except OSError:
        return False


def _run_onec_job(payload, outputs, client_exe, infobase, epf, log, what):
    """Launch the platform once for one JSON job and validate what came back."""
    work = tempfile.mkdtemp(prefix="mxl-render-")
    try:
        job = os.path.join(work, "job.json")
        status = os.path.join(work, "status.json")
        onec_log = os.path.join(work, "1c.log")
        payload = dict(payload)
        payload["statusPath"] = status
        with open(job, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, ensure_ascii=False)

        command = [client_exe, "ENTERPRISE", "/IBConnectionString", _file_conn(infobase)]
        if ONEC_USER:
            command += ["/N", ONEC_USER, "/P", ONEC_PASSWORD]
        command += ["/DisableStartupDialogs", "/DisableStartupMessages",
                    "/Execute", epf, "/C", job, "/Out", onec_log]

        if log:
            log(what)
        done = _run(command, ONEC_TIMEOUT, "rendering")

        if os.path.isfile(status):
            try:
                with open(status, encoding="utf-8-sig") as fh:
                    result = json.load(fh)
            except (OSError, ValueError) as exc:
                raise OneCError("Renderer status file could not be read: %s" % exc)
            if result.get("success") is not True:
                raise OneCError(str(result.get("error") or "1C reported a failure"))
        else:
            raise OneCError("1C did not report a result (exit %s).\n%s"
                            % (done.returncode,
                               _text_of(onec_log) or done.stderr or done.stdout
                               or "no log was produced"))

        missing = [os.path.basename(o) for o in outputs if not os.path.isfile(o)]
        if missing:
            raise OneCError("1C reported success but produced no HTML for: %s"
                            % ", ".join(missing))
    finally:
        shutil.rmtree(work, ignore_errors=True)


# --------------------------------------------------------------------------
#  Resident renderer
# --------------------------------------------------------------------------
#
#  Starting the platform costs seconds and flashes a window, which is painful
#  when every preview pays it. MxlToHtmlService.epf stays running instead and
#  serves a queue of jobs from a session folder.
#
#  Persistent mode is used only when that .epf is present next to this script.
#  It is a separate file from the one-shot MxlToHtml.epf on purpose: the two
#  take different launch parameters (a session folder vs a job file), and
#  handing a folder to the old processor would leave a stuck window behind.

ONEC_READY_TIMEOUT = 45         # seconds to wait for the worker to report ready
_WORKERS = {}
_WORKER_REFUSED = set()


def service_epf():
    return onec_file("MxlToHtmlService.epf")


class OneCWorker(object):
    """A 1C:Enterprise process kept alive to render many documents."""

    def __init__(self, client_exe, infobase, log=None):
        self.client_exe = client_exe
        self.infobase = infobase
        self.log = log
        self.folder = None
        self.process = None
        self.failure = ""
        self._counter = 0
        self._lock = threading.Lock()

    # ---------------------------------------------------------------- paths

    @property
    def jobs(self):
        return os.path.join(self.folder, "jobs")

    @property
    def done(self):
        return os.path.join(self.folder, "done")

    # --------------------------------------------------------------- launch

    def start(self):
        epf = service_epf()
        if not os.path.isfile(epf):
            return False
        self.folder = tempfile.mkdtemp(prefix="mxl-onec-session-")
        # the processor creates jobs/ and done/ itself: if they never appear we
        # know the module did not run, rather than guessing

        command = [self.client_exe, "ENTERPRISE",
                   "/IBConnectionString", _file_conn(self.infobase)]
        if ONEC_USER:
            command += ["/N", ONEC_USER, "/P", ONEC_PASSWORD]
        command += ["/DisableStartupDialogs", "/DisableStartupMessages",
                    "/Execute", epf, "/C", self.folder,
                    "/Out", os.path.join(self.folder, "1c.log")]
        try:
            self.process = subprocess.Popen(
                command, startupinfo=_startupinfo(),
                stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except OSError as exc:
            self._cleanup()
            raise OneCError("Could not start the 1C renderer:\n%s" % exc)

        if self.log:
            self.log("Starting the resident 1C renderer...")
            self.log("   session folder: %s" % self.folder)

        ready = os.path.join(self.folder, "ready.json")
        deadline = time.time() + ONEC_READY_TIMEOUT
        while time.time() < deadline:
            if os.path.isfile(ready):
                if self.log:
                    self.log("   ready after %.1fs; later previews reuse this session."
                             % (ONEC_READY_TIMEOUT - (deadline - time.time())))
                quiet_windows(self.process.pid)
                return True
            if self.process.poll() is not None:
                break                       # it gave up before reporting ready
            time.sleep(0.2)

        self.failure = self._why_not_ready()
        if self.log:
            for line in self.failure.splitlines():
                self.log("   " + line)
        self.stop()
        return False

    def _why_not_ready(self):
        """Everything we know about a worker that never said hello."""
        bits = []
        code = self.process.poll() if self.process else None
        bits.append("no ready.json after %ds" % ONEC_READY_TIMEOUT
                    if code is None else
                    "1C exited with code %s before reporting ready" % code)

        trace = _text_of(os.path.join(self.folder, "startup.log"))
        if trace:
            bits.append("the processor's own trace:")
            for line in trace.splitlines():
                bits.append("   " + line)
        else:
            bits.append("startup.log was never written, so the form module never ran.")
            bits.append("Check that MxlToHtmlService.bsl is the module of the")
            bits.append("data processor's DEFAULT MANAGED FORM, not the object module,")
            bits.append("and that the form is set as the default form.")
        onec_log = _text_of(os.path.join(self.folder, "1c.log"))
        if onec_log:
            bits.append("1C log: " + onec_log.replace("\n", " | ")[:500])
        try:
            if self.process and self.process.poll() is not None:
                out, err = self.process.communicate(timeout=2)
                for stream, label in ((err, "stderr"), (out, "stdout")):
                    text = (stream or b"").decode("utf-8", "replace").strip()
                    if text:
                        bits.append("%s: %s" % (label, text[:300]))
        except Exception:
            pass
        try:
            listing = sorted(os.listdir(self.folder))
            bits.append("session folder holds: %s" % (", ".join(listing) or "nothing"))
        except OSError:
            pass
        return "\n".join(bits)

    @property
    def alive(self):
        return (self.process is not None and self.process.poll() is None
                and self.folder is not None and os.path.isdir(self.folder))

    # --------------------------------------------------------------- render

    def render(self, prepared):
        """Queue one job and wait for it. `prepared` is [(name, src, dst), ...]."""
        with self._lock:
            if not self.alive:
                raise OneCError("The resident 1C renderer is no longer running.")
            self._counter += 1
            job_id = "job%04d.json" % self._counter
            payload = {"items": [{"name": n, "inputPath": i, "outputPath": o}
                                 for n, i, o in prepared]}

            # write beside the queue, then move in, so 1C never sees a half file
            os.makedirs(self.jobs, exist_ok=True)
            staging = os.path.join(self.folder, job_id + ".tmp")
            with open(staging, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, ensure_ascii=False)
            os.replace(staging, os.path.join(self.jobs, job_id))

            quiet_windows(self.process.pid)
            result = os.path.join(self.done, job_id)
            deadline = time.time() + ONEC_TIMEOUT
            while time.time() < deadline:
                if os.path.isfile(result):
                    break
                if self.process.poll() is not None:
                    raise OneCError("The 1C renderer stopped while rendering.\n%s"
                                    % (_text_of(os.path.join(self.folder, "1c.log"))
                                       or "no log was produced"))
                time.sleep(0.1)
            else:
                raise OneCError("Timed out waiting for the 1C renderer.")

            try:
                with open(result, encoding="utf-8-sig") as fh:
                    status = json.load(fh)
            except (OSError, ValueError) as exc:
                raise OneCError("Renderer status file could not be read: %s" % exc)
            finally:
                try:
                    os.remove(result)
                except OSError:
                    pass

            if status.get("success") is not True:
                raise OneCError(str(status.get("error") or "1C reported a failure"))

        missing = [os.path.basename(o) for _, _, o in prepared if not os.path.isfile(o)]
        if missing:
            raise OneCError("1C reported success but produced no HTML for: %s"
                            % ", ".join(missing))

    # ----------------------------------------------------------------- stop

    def stop(self):
        if self.folder and os.path.isdir(self.folder):
            try:
                open(os.path.join(self.folder, "stop"), "w").close()
            except OSError:
                pass
        if self.process is not None:
            for _ in range(30):             # give it three seconds to go quietly
                if self.process.poll() is not None:
                    break
                time.sleep(0.1)
            if self.process.poll() is None:
                try:
                    self.process.terminate()
                except OSError:
                    pass
        self._cleanup()

    def _cleanup(self):
        if self.folder:
            shutil.rmtree(self.folder, ignore_errors=True)
        self.folder = None
        self.process = None


def acquire_worker(client_exe, infobase, log=None):
    """A live resident renderer, or None when persistent mode is unavailable."""
    if not ONEC_PERSISTENT or not os.path.isfile(service_epf()):
        return None
    key = (client_exe, infobase)
    if key in _WORKER_REFUSED:
        return None
    worker = _WORKERS.get(key)
    if worker is not None and worker.alive:
        return worker

    worker = OneCWorker(client_exe, infobase, log)
    try:
        started = worker.start()
    except OneCError:
        started = False
    if not started:
        _WORKER_REFUSED.add(key)            # do not retry all session
        if log:
            log("Resident 1C renderer did not start; using one launch per render.")
            if getattr(worker, "failure", ""):
                log("   reason: %s" % worker.failure.splitlines()[0])
        return None
    _WORKERS[key] = worker
    return worker


def shutdown_workers():
    for worker in list(_WORKERS.values()):
        try:
            worker.stop()
        except Exception:
            pass
    _WORKERS.clear()


atexit.register(shutdown_workers)


def render_mxl_batch(items, client_exe, infobase=None, log=None):
    """
    Render several .mxl files in ONE platform launch.

    `items` is a sequence of (name, mxl_path, html_path). Starting 1C costs a
    few seconds, so batching matters as soon as there is more than one document.
    Falls back to one launch per document if the installed data processor is the
    legacy single-document build.

    Returns {name: html_path}.
    """
    epf, _dt = onec_assets()
    if not os.path.isfile(epf):
        raise OneCError("Renderer data processor is missing:\n%s" % epf)
    if not client_exe or not os.path.isfile(client_exe):
        raise OneCError("1C client executable is not configured.")

    prepared = []
    for name, source, target in items:
        source, target = os.path.abspath(source), os.path.abspath(target)
        if not os.path.isfile(source):
            raise OneCError("Input MXL was not found:\n%s" % source)
        folder = os.path.dirname(target)
        if folder:
            os.makedirs(folder, exist_ok=True)
        prepared.append((str(name), source, target))
    if not prepared:
        raise OneCError("Nothing to render.")

    infobase = infobase or default_infobase(client_exe)
    ensure_infobase(client_exe, infobase, log)

    worker = acquire_worker(client_exe, infobase, log)
    if worker is not None:
        if log:
            log("Rendering %d document(s) in the resident 1C session..."
                % len(prepared))
        try:
            worker.render(prepared)
            return dict((n, o) for n, _, o in prepared)
        except OneCError:
            if worker.alive:
                raise
            # the session died: drop it and fall through to a fresh launch
            _WORKERS.pop((client_exe, infobase), None)
            if log:
                log("The resident session ended; falling back to a single launch.")

    def one(name, source, target, note):
        _run_onec_job({"inputPath": source, "outputPath": target},
                      [target], client_exe, infobase, epf, log, note)

    if len(prepared) == 1:
        name, source, target = prepared[0]
        one(name, source, target, "Rendering with 1C:Enterprise...")
    elif not epf_supports_batch(epf):
        if log:
            log("Data processor renders one document per launch; "
                "running %d launches." % len(prepared))
        for index, (name, source, target) in enumerate(prepared, 1):
            one(name, source, target,
                "Rendering %d of %d with 1C:Enterprise..." % (index, len(prepared)))
    else:
        payload = {"items": [{"name": n, "inputPath": i, "outputPath": o}
                             for n, i, o in prepared]}
        _run_onec_job(payload, [o for _, _, o in prepared], client_exe, infobase, epf,
                      log, "Rendering %d documents in one 1C:Enterprise launch..."
                           % len(prepared))
    return dict((n, o) for n, _, o in prepared)


def render_mxl_to_html(mxl_path, html_path, client_exe, infobase=None, log=None):
    """Render one .mxl to standalone HTML using the 1C platform."""
    render_mxl_batch([("result", mxl_path, html_path)], client_exe, infobase, log)
    return html_path


def write_modified_copy(path, rules, skip_existing, dest):
    """Write the document with the rules applied to `dest`. Returns rules applied."""
    hdr, bom, text = load(path)
    applied = []
    for rule in rules:
        patches, _outcomes = plan_rule(text, rule, skip_existing)
        if patches:
            text = apply_patches(text, patches)
            applied.append(rule)
    save(dest, hdr, bom, text)
    return applied


# ==========================================================================
#  Look and feel
# ==========================================================================

THEMES = {
    "dark": {
        "BG": "#1b2027", "CARD": "#232a33", "SIDEBAR": "#161b21", "PANEL": "#1e242b",
        "BORDER": "#323b46", "TEXT": "#e4e9ef", "MUTED": "#94a0ad",
        "ACCENT": "#4f8ff7", "ACCENT_DK": "#3d7ae4", "ACCENT_LT": "#2b3d59",
        "ACCENT_BTN": "#2f6fe0", "ACCENT_BTN_DK": "#2559bd",
        "ACCENT_DIS": "#2f4c7a", "ON_ACCENT": "#ffffff",
        "OK_FG": "#63d18c", "WARN_FG": "#e8b061", "ERR_FG": "#ef7c7f",
        "HIGHLIGHT": "#1e4a33", "HEAD_BG": "#2b333d", "HOVER": "#1f262e",
        "BTN_BG": "#2c333d", "BTN_HOVER": "#39424e",
        "DIS_BG": "#252b33", "DIS_FG": "#5d6774",
        "FIELD": "#1c222a", "LOGBG": "#191f26", "TROUGH": "#2b333d",
    },
    "light": {
        "BG": "#f3f5f7", "CARD": "#ffffff", "SIDEBAR": "#e9edf2", "PANEL": "#f7f9fb",
        "BORDER": "#dde2e8", "TEXT": "#1f2933", "MUTED": "#6b7280",
        "ACCENT": "#2563eb", "ACCENT_DK": "#1d4ed8", "ACCENT_LT": "#dbe6fe",
        "ACCENT_BTN": "#2563eb", "ACCENT_BTN_DK": "#1d4ed8",
        "ACCENT_DIS": "#a9c0f2", "ON_ACCENT": "#ffffff",
        "OK_FG": "#15803d", "WARN_FG": "#b45309", "ERR_FG": "#b91c1c",
        "HIGHLIGHT": "#d8f0dc", "HEAD_BG": "#eef1f5", "HOVER": "#dfe5ec",
        "BTN_BG": "#ffffff", "BTN_HOVER": "#eef2f7",
        "DIS_BG": "#f6f7f9", "DIS_FG": "#a5adb8",
        "FIELD": "#ffffff", "LOGBG": "#fbfcfd", "TROUGH": "#e4e8ee",
    },
}

THEME_NAME = "dark"
_FAMILY = "TkDefaultFont"


def apply_theme(name):
    """Publish one palette as module-level names, so widgets can just use them."""
    global THEME_NAME
    THEME_NAME = name if name in THEMES else "dark"
    globals().update(THEMES[THEME_NAME])


apply_theme(THEME_NAME)


def setup_style(root):
    """One theme with a single accent colour, applied to ttk and tk alike."""
    global _FAMILY
    import tkinter.font as tkfont

    available = set(tkfont.families(root))
    for candidate in ("Segoe UI", "Inter", "Helvetica Neue", "DejaVu Sans", "Arial"):
        if candidate in available:
            _FAMILY = candidate
            break

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    root.configure(bg=BG)

    # widgets that are not ttk read these defaults
    for pattern, value in (("*Menu.background", CARD), ("*Menu.foreground", TEXT),
                           ("*Menu.activeBackground", ACCENT), ("*Menu.activeForeground", ON_ACCENT),
                           ("*Menu.relief", "flat"),
                           ("*TCombobox*Listbox.background", CARD),
                           ("*TCombobox*Listbox.foreground", TEXT),
                           ("*TCombobox*Listbox.selectBackground", ACCENT),
                           ("*TCombobox*Listbox.selectForeground", ON_ACCENT)):
        try:
            root.option_add(pattern, value)
        except tk.TclError:
            pass

    style.configure(".", font=(_FAMILY, 10), background=BG, foreground=TEXT)
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=CARD)
    style.configure("TLabel", background=BG, foreground=TEXT)
    style.configure("Card.TLabel", background=CARD, foreground=TEXT)
    style.configure("Muted.TLabel", background=BG, foreground=MUTED)
    style.configure("CardMuted.TLabel", background=CARD, foreground=MUTED)
    style.configure("H1.TLabel", background=BG, foreground=TEXT, font=(_FAMILY, 16, "bold"))
    style.configure("H2.TLabel", background=CARD, foreground=TEXT, font=(_FAMILY, 11, "bold"))
    style.configure("CardOk.TLabel", background=CARD, foreground=OK_FG)
    style.configure("CardWarn.TLabel", background=CARD, foreground=WARN_FG)

    style.configure("TButton", padding=(13, 7), relief="flat", borderwidth=1,
                    background=BTN_BG, foreground=TEXT,
                    bordercolor=BORDER, lightcolor=BTN_BG, darkcolor=BTN_BG,
                    focuscolor=BG)
    style.map("TButton",
              background=[("active", BTN_HOVER), ("disabled", DIS_BG)],
              lightcolor=[("active", BTN_HOVER)], darkcolor=[("active", BTN_HOVER)],
              foreground=[("disabled", DIS_FG)])
    style.configure("Accent.TButton", padding=(18, 8), relief="flat", borderwidth=1,
                    background=ACCENT_BTN, foreground=ON_ACCENT, bordercolor=ACCENT_BTN,
                    lightcolor=ACCENT_BTN, darkcolor=ACCENT_BTN,
                    font=(_FAMILY, 10, "bold"))
    style.map("Accent.TButton",
              background=[("active", ACCENT_BTN_DK), ("disabled", ACCENT_DIS)],
              lightcolor=[("active", ACCENT_BTN_DK)],
              darkcolor=[("active", ACCENT_BTN_DK)],
              bordercolor=[("active", ACCENT_BTN_DK), ("disabled", ACCENT_DIS)],
              foreground=[("disabled", ON_ACCENT)])

    style.configure("Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TEXT, rowheight=24, borderwidth=0, relief="flat")
    style.configure("Treeview.Heading", background=HEAD_BG, foreground=MUTED,
                    relief="flat", borderwidth=0, padding=(8, 7),
                    font=(_FAMILY, 9, "bold"))
    style.map("Treeview.Heading", background=[("active", HOVER)])
    style.map("Treeview",
              background=[("selected", ACCENT_LT)],
              foreground=[("selected", TEXT)])

    style.configure("TCheckbutton", background=CARD, foreground=TEXT,
                    indicatorcolor=FIELD, focuscolor=CARD)
    style.map("TCheckbutton",
              background=[("active", CARD)],
              indicatorcolor=[("selected", ACCENT), ("pressed", ACCENT_DK)])
    style.configure("TRadiobutton", background=CARD, foreground=TEXT,
                    indicatorcolor=FIELD, focuscolor=CARD)
    style.map("TRadiobutton",
              background=[("active", CARD)],
              indicatorcolor=[("selected", ACCENT), ("pressed", ACCENT_DK)])

    style.configure("TEntry", padding=6, fieldbackground=FIELD, foreground=TEXT,
                    insertcolor=TEXT, bordercolor=BORDER, lightcolor=BORDER,
                    darkcolor=BORDER)
    style.configure("TCombobox", padding=5, fieldbackground=FIELD, background=BTN_BG,
                    foreground=TEXT, arrowcolor=MUTED, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER)
    style.map("TCombobox", fieldbackground=[("readonly", FIELD)],
              foreground=[("readonly", TEXT)])

    style.configure("TScrollbar", background=BTN_BG, troughcolor=BG,
                    bordercolor=BG, arrowcolor=MUTED, relief="flat")
    style.map("TScrollbar", background=[("active", BTN_HOVER)])
    style.configure("Thin.Horizontal.TProgressbar", background=ACCENT,
                    troughcolor=TROUGH, bordercolor=TROUGH,
                    lightcolor=ACCENT, darkcolor=ACCENT, borderwidth=0, thickness=5)
    style.configure("TSeparator", background=BORDER)

    style.configure("TNotebook", background=CARD, borderwidth=0,
                    bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                    tabmargins=(0, 4, 0, 0))
    style.configure("TNotebook.Tab", background=HEAD_BG, foreground=MUTED,
                    padding=(18, 9), borderwidth=0, focuscolor=CARD,
                    lightcolor=HEAD_BG, darkcolor=HEAD_BG, bordercolor=BORDER,
                    font=(_FAMILY, 10))
    style.map("TNotebook.Tab",
              background=[("selected", CARD), ("active", HOVER)],
              foreground=[("selected", ACCENT), ("active", TEXT)],
              lightcolor=[("selected", CARD)],
              darkcolor=[("selected", CARD)],
              font=[("selected", (_FAMILY, 10, "bold"))])

    enable_clipboard_shortcuts(root)
    return style


def enable_clipboard_shortcuts(root):
    """
    Make Ctrl+C/V/X/A work whatever the keyboard layout is.

    Tk binds the clipboard events by keysym, so with a Cyrillic layout Ctrl+V
    arrives as Cyrillic_em and the built-in <<Paste>> binding never fires --
    pasting silently does nothing until you switch to a Latin layout. Matching
    on the hardware keycode instead is layout-agnostic.

    When the keysym IS the Latin letter we do nothing and let Tk's own binding
    handle it, otherwise the event would be delivered twice and paste twice.
    """
    actions = {86: "<<Paste>>", 67: "<<Copy>>", 88: "<<Cut>>", 65: "select-all"}

    def on_control_key(event):
        if event.keysym.lower() in ("v", "c", "x", "a"):
            return None                      # Latin layout: Tk already handles it
        action = actions.get(getattr(event, "keycode", None))
        if action is None:
            return None
        widget = event.widget
        if action == "select-all":
            try:
                widget.select_range(0, "end")
                widget.icursor("end")
            except Exception:
                try:
                    widget.tag_add("sel", "1.0", "end-1c")
                except Exception:
                    return None
            return "break"
        try:
            widget.event_generate(action)
        except Exception:
            return None
        return "break"

    def popup(event):
        widget = event.widget
        menu = tk.Menu(widget, tearoff=0)
        menu.add_command(label="Cut",
                         command=lambda: widget.event_generate("<<Cut>>"))
        menu.add_command(label="Copy",
                         command=lambda: widget.event_generate("<<Copy>>"))
        menu.add_command(label="Paste",
                         command=lambda: widget.event_generate("<<Paste>>"))
        menu.add_separator()

        def select_all():
            try:
                widget.select_range(0, "end")
                widget.icursor("end")
            except Exception:
                widget.tag_add("sel", "1.0", "end-1c")

        menu.add_command(label="Select all", command=select_all)
        try:
            widget.focus_set()
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
        return "break"

    for klass in ("TEntry", "Entry", "TCombobox", "Text"):
        try:
            root.bind_class(klass, "<Control-KeyPress>", on_control_key, add="+")
            root.bind_class(klass, "<Button-3>", popup, add="+")
        except tk.TclError:
            pass


def card(parent, **kwargs):
    """A white panel with a hairline border."""
    frame = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                     highlightcolor=BORDER, highlightthickness=1, bd=0, **kwargs)
    return frame


def hint(parent, text, **grid):
    label = ttk.Label(parent, text=text, style="CardMuted.TLabel", justify="left")
    if grid:
        label.grid(**grid)
    return label


# ==========================================================================
#  Dialogs
# ==========================================================================

class ModalDialog(tk.Toplevel):
    """Shared plumbing: centred on the parent, modal, Esc closes."""

    def __init__(self, master, title):
        tk.Toplevel.__init__(self, master, bg=BG)
        self.title(title)
        self.resizable(False, False)
        self.transient(master.winfo_toplevel())
        self.bind("<Escape>", lambda e: self.destroy())

    def centre(self):
        self.update_idletasks()
        parent = self.master.winfo_toplevel()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
        self.geometry("+%d+%d" % (max(x, 0), max(y, 0)))
        self.grab_set()


class RuleDialog(ModalDialog):
    """Build one modification -- add a column, or delete one -- with a live preview."""

    def __init__(self, master, registers, on_add, initial=None, anchor=None):
        ModalDialog.__init__(self, master, "Add modification")
        self.registers = registers
        self.on_add = on_add

        body = card(self)
        body.pack(fill="both", expand=True, padx=14, pady=14)
        outer = ttk.Frame(body, style="Card.TFrame", padding=14)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)

        self.tabs = ttk.Notebook(outer)
        self.tabs.grid(row=0, column=0, sticky="ew")
        self.tabs.bind("<<NotebookTabChanged>>", lambda e: self._tab_changed())

        self.add_tab = ttk.Frame(self.tabs, style="Card.TFrame", padding=14)
        self.del_tab = ttk.Frame(self.tabs, style="Card.TFrame", padding=14)
        self.ren_tab = ttk.Frame(self.tabs, style="Card.TFrame", padding=14)
        self.tabs.add(self.add_tab, text="  Add column  ")
        self.tabs.add(self.ren_tab, text="  Rename column  ")
        self.tabs.add(self.del_tab, text="  Delete column  ")
        self._build_add(self.add_tab)
        self._build_rename(self.ren_tab)
        self._build_delete(self.del_tab)

        prev = tk.Frame(outer, bg=PANEL, highlightbackground=BORDER,
                        highlightthickness=1, bd=0)
        prev.grid(row=1, column=0, sticky="ew", pady=(14, 0))
        self.preview = tk.Label(prev, text="", bg=PANEL, fg=MUTED,
                                font=("TkFixedFont", 9), justify="left",
                                anchor="w", padx=10, pady=8, wraplength=560)
        self.preview.pack(fill="x")

        self.err = ttk.Label(outer, text="", style="CardWarn.TLabel")
        self.err.grid(row=2, column=0, sticky="w", pady=(10, 0))

        btns = ttk.Frame(outer, style="Card.TFrame")
        btns.grid(row=3, column=0, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        self.ok = ttk.Button(btns, text="Add modification",
                             style="Accent.TButton", command=self._submit)
        self.ok.pack(side="right")

        if initial:
            self.v_reg.set(initial)
            self.d_reg.set(initial)
            self.r_reg.set(initial)
            self._register_changed()
            self._del_register_changed()
            self._ren_register_changed()
        if anchor:
            # arrived here by double-clicking a column: use it everywhere it fits
            self.v_anchor.set(anchor)
            self.d_col.set(anchor)
            self.r_col.set(anchor)
            self._ren_column_changed()
            self.tabs.select(0)
            self._tab_changed()
            self.after(50, self.e_new.focus_set)
        self._refresh()
        self.centre()

    # ---------------------------------------------------------------- tabs

    def _build_add(self, box):
        box.columnconfigure(1, weight=1)
        hint(box, "Insert a column next to an existing one and fill every data row "
                  "with a fixed value.").grid(row=0, column=0, columnspan=2,
                                              sticky="w", pady=(0, 12))

        ttk.Label(box, text="Register", style="Card.TLabel")\
            .grid(row=1, column=0, sticky="w", pady=4)
        self.v_reg = tk.StringVar()
        self.c_reg = ttk.Combobox(box, textvariable=self.v_reg,
                                  values=sorted(self.registers), width=44)
        self.c_reg.grid(row=1, column=1, sticky="ew", pady=4)
        self.c_reg.bind("<<ComboboxSelected>>", lambda e: self._register_changed())
        self.v_reg.trace_add("write", lambda *a: self._refresh())

        ttk.Label(box, text="Position", style="Card.TLabel")\
            .grid(row=2, column=0, sticky="w", pady=4)
        pos = ttk.Frame(box, style="Card.TFrame")
        pos.grid(row=2, column=1, sticky="w", pady=4)
        self.v_pos = tk.StringVar(value="after")
        for value in ("after", "before"):
            ttk.Radiobutton(pos, text=value, value=value, variable=self.v_pos,
                            command=self._refresh).pack(side="left", padx=(0, 12))

        ttk.Label(box, text="this column", style="Card.TLabel")\
            .grid(row=3, column=0, sticky="w", pady=4)
        self.v_anchor = tk.StringVar()
        self.c_anchor = ttk.Combobox(box, textvariable=self.v_anchor, width=44)
        self.c_anchor.grid(row=3, column=1, sticky="ew", pady=4)
        self.v_anchor.trace_add("write", lambda *a: self._refresh())

        ttk.Separator(box).grid(row=4, column=0, columnspan=2, sticky="ew", pady=12)

        ttk.Label(box, text="New column", style="Card.TLabel")\
            .grid(row=5, column=0, sticky="w", pady=4)
        self.v_new = tk.StringVar()
        self.e_new = ttk.Entry(box, textvariable=self.v_new, width=44)
        self.e_new.grid(row=5, column=1, sticky="ew", pady=4)
        self.v_new.trace_add("write", lambda *a: self._refresh())

        ttk.Label(box, text="Value in data rows", style="Card.TLabel")\
            .grid(row=6, column=0, sticky="w", pady=4)
        self.v_val = tk.StringVar()
        ttk.Entry(box, textvariable=self.v_val, width=44)\
            .grid(row=6, column=1, sticky="ew", pady=4)
        hint(box, "Leave empty to insert a blank cell.")\
            .grid(row=7, column=1, sticky="w")

    def _build_delete(self, box):
        box.columnconfigure(1, weight=1)
        hint(box, "Remove a column and every cell under it, in each file that has "
                  "this register.").grid(row=0, column=0, columnspan=2,
                                         sticky="w", pady=(0, 12))

        ttk.Label(box, text="Register", style="Card.TLabel")\
            .grid(row=1, column=0, sticky="w", pady=4)
        self.d_reg = tk.StringVar()
        self.dc_reg = ttk.Combobox(box, textvariable=self.d_reg,
                                   values=sorted(self.registers), width=44)
        self.dc_reg.grid(row=1, column=1, sticky="ew", pady=4)
        self.dc_reg.bind("<<ComboboxSelected>>", lambda e: self._del_register_changed())
        self.d_reg.trace_add("write", lambda *a: self._refresh())

        ttk.Label(box, text="Column to delete", style="Card.TLabel")\
            .grid(row=2, column=0, sticky="w", pady=4)
        self.d_col = tk.StringVar()
        self.dc_col = ttk.Combobox(box, textvariable=self.d_col, width=44)
        self.dc_col.grid(row=2, column=1, sticky="ew", pady=4)
        self.d_col.trace_add("write", lambda *a: self._refresh())

        hint(box, "Files that do not have this column are left untouched, so running\n"
                  "the same deletion twice is safe.").grid(row=3, column=0, columnspan=2,
                                                           sticky="w", pady=(12, 0))

    def _build_rename(self, box):
        box.columnconfigure(1, weight=1)
        hint(box, "Change a column's caption. The data underneath is left exactly "
                  "as it is.").grid(row=0, column=0, columnspan=2,
                                    sticky="w", pady=(0, 12))

        ttk.Label(box, text="Register", style="Card.TLabel")\
            .grid(row=1, column=0, sticky="w", pady=4)
        self.r_reg = tk.StringVar()
        self.rc_reg = ttk.Combobox(box, textvariable=self.r_reg,
                                   values=sorted(self.registers), width=44)
        self.rc_reg.grid(row=1, column=1, sticky="ew", pady=4)
        self.rc_reg.bind("<<ComboboxSelected>>", lambda e: self._ren_register_changed())
        self.r_reg.trace_add("write", lambda *a: self._refresh())

        ttk.Label(box, text="Column", style="Card.TLabel")\
            .grid(row=2, column=0, sticky="w", pady=4)
        self.r_col = tk.StringVar()
        self.rc_col = ttk.Combobox(box, textvariable=self.r_col, width=44)
        self.rc_col.grid(row=2, column=1, sticky="ew", pady=4)
        self.r_col.trace_add("write", lambda *a: self._ren_column_changed())

        ttk.Label(box, text="New name", style="Card.TLabel")\
            .grid(row=3, column=0, sticky="w", pady=4)
        self.r_new = tk.StringVar()
        ttk.Entry(box, textvariable=self.r_new, width=44)\
            .grid(row=3, column=1, sticky="ew", pady=4)
        self.r_new.trace_add("write", lambda *a: self._refresh())

        hint(box, "Only the header caption changes: column order, widths and every\n"
                  "data row stay untouched.").grid(row=4, column=0, columnspan=2,
                                                   sticky="w", pady=(12, 0))

    # ------------------------------------------------------------- helpers

    def _mode(self):
        """Which tab is in front: 'add', 'rename' or 'delete'."""
        try:
            return ("add", "rename", "delete")[self.tabs.index(self.tabs.select())]
        except Exception:
            return "add"

    def _deleting(self):
        return self._mode() == "delete"

    def _columns(self, register):
        return self.registers.get(register.strip(), {}).get("columns", [])

    def _register_changed(self):
        cols = self._columns(self.v_reg.get())
        self.c_anchor.configure(values=cols)
        if self.v_anchor.get() not in cols:
            self.v_anchor.set("")
        self._refresh()

    def _del_register_changed(self):
        cols = [c for c in self._columns(self.d_reg.get()) if c]
        self.dc_col.configure(values=cols)
        if self.d_col.get() not in cols:
            self.d_col.set("")
        self._refresh()

    def _ren_register_changed(self):
        cols = [c for c in self._columns(self.r_reg.get()) if c]
        self.rc_col.configure(values=cols)
        if self.r_col.get() not in cols:
            self.r_col.set("")
        self._refresh()

    def _ren_column_changed(self):
        # start from the current name so the user only edits what differs
        if self.r_col.get() and not self.r_new.get().strip():
            self.r_new.set(self.r_col.get())
        self._refresh()

    def _tab_changed(self):
        label = {"add": "Add modification", "rename": "Rename column",
                 "delete": "Delete column"}[self._mode()]
        self.ok.configure(text=label)
        self.title(label)
        self._refresh()

    def _refresh(self):
        mode = self._mode()
        if mode == "delete":
            self._refresh_delete()
        elif mode == "rename":
            self._refresh_rename()
        else:
            self._refresh_add()

    def _refresh_rename(self):
        cols = self._columns(self.r_reg.get())
        target = self.r_col.get().strip()
        new = self.r_new.get().strip()
        problem = ""
        if not self.r_reg.get().strip():
            problem = "Choose a register."
        elif target and cols and target not in cols:
            problem = "No column with that name was found in this register."
        elif target and cols.count(target) > 1:
            problem = "That name appears more than once here; the rename will be skipped."
        elif new and new in cols and new != target:
            problem = "A column with the new name already exists here."
        elif new and new == target:
            problem = "The new name is the same as the current one."
        self.err.configure(text=problem)

        if target in cols and new and new != target:
            i = cols.index(target)
            after = list(cols)
            after[i] = new
            lo = max(0, i - 1)
            before_row = "   ".join(c or "\u00b7" for c in cols[lo:i + 2])
            after_row = "   ".join(("[ %s ]" % c) if c == new else (c or "\u00b7")
                                    for c in after[lo:i + 2])
            self.preview.configure(
                text=("Before:\n" + before_row + "\n\nAfter:\n" + after_row
                      + "\n\nPosition %d of %d; no data rows change."
                      % (i + 1, len(cols))),
                fg=TEXT)
        else:
            self.preview.configure(
                text="Pick a register and a column, then type the new name.", fg=MUTED)

    def _refresh_add(self):
        cols = self._columns(self.v_reg.get())
        anchor = self.v_anchor.get().strip()
        new = self.v_new.get().strip() or "<new column>"
        problem = ""
        if not self.v_reg.get().strip():
            problem = "Choose a register."
        elif anchor and cols and anchor not in cols:
            problem = "No column with that name was found in this register."
        elif new != "<new column>" and new in cols:
            problem = "That column already exists here; files having it will be skipped."
        self.err.configure(text=problem)

        if anchor in cols:
            i = cols.index(anchor)
            order = list(cols)
            order.insert(i + 1 if self.v_pos.get() == "after" else i, new)
            j = order.index(new)
            shown = order[max(0, j - 2):j + 3]
            text = "   ".join(("[ %s ]" % c) if c == new else (c or "\u00b7") for c in shown)
            self.preview.configure(text="Resulting column order:\n" + text, fg=TEXT)
        else:
            self.preview.configure(
                text="Pick a register and an anchor column to see the resulting order.",
                fg=MUTED)

    def _refresh_delete(self):
        cols = self._columns(self.d_reg.get())
        target = self.d_col.get().strip()
        problem = ""
        if not self.d_reg.get().strip():
            problem = "Choose a register."
        elif target and cols and target not in cols:
            problem = "No column with that name was found in this register."
        elif target and cols.count(target) > 1:
            problem = "That name appears more than once here; the deletion will be skipped."
        self.err.configure(text=problem)

        if target in cols and cols.count(target) == 1:
            i = cols.index(target)
            shown = cols[max(0, i - 2):i + 3]
            text = "   ".join(("[ %s ]" % c) if c == target else (c or "\u00b7")
                              for c in shown)
            remaining = [c for c in cols if c != target]
            after = remaining[max(0, i - 2):i + 2]
            self.preview.configure(
                text=("Removing:\n" + text + "\n\nLeaves:\n"
                      + "   ".join(c or "\u00b7" for c in after)
                      + "\n\n%d columns become %d." % (len(cols), len(remaining))),
                fg=TEXT)
        else:
            self.preview.configure(
                text="Pick a register and the column to remove.", fg=MUTED)

    def _submit(self):
        if self._mode() == "rename":
            reg = self.r_reg.get().strip()
            col = self.r_col.get().strip()
            new = self.r_new.get().strip()
            if not (reg and col and new):
                self.err.configure(text="Register, column and new name are all required.")
                return
            if new == col:
                self.err.configure(text="The new name is the same as the current one.")
                return
            self.on_add(Rule(reg, col, new_col=new, action="rename"))
            self.destroy()
            return
        if self._deleting():
            reg, col = self.d_reg.get().strip(), self.d_col.get().strip()
            if not (reg and col):
                self.err.configure(text="Register and column are both required.")
                return
            self.on_add(Rule(reg, col, action="delete"))
        else:
            reg, anchor = self.v_reg.get().strip(), self.v_anchor.get().strip()
            new = self.v_new.get().strip()
            if not (reg and anchor and new):
                self.err.configure(
                    text="Register, anchor column and new column name are required.")
                return
            self.on_add(Rule(reg, anchor, self.v_pos.get(), new, self.v_val.get()))
        self.destroy()


class OneCSettingsDialog(ModalDialog):
    """Configure the 1C:Enterprise client used for exact previews."""

    def __init__(self, master, current, on_save):
        ModalDialog.__init__(self, master, "1C:Enterprise renderer")
        self.on_save = on_save

        body = card(self)
        body.pack(fill="both", expand=True, padx=14, pady=14)
        inner = ttk.Frame(body, style="Card.TFrame", padding=16)
        inner.pack(fill="both", expand=True)
        inner.columnconfigure(1, weight=1)

        ttk.Label(inner, text="Exact preview renderer",
                  style="H2.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        hint(inner,
             "Exact previews are produced by 1C:Enterprise itself: the document is handed\n"
             "to the bundled MxlToHtml data processor, which reads it into a\n"
             "SpreadsheetDocument and writes HTML.\n\n"
             "Point this at the thin client (1cv8c.exe). 1cv8.exe must sit beside it; it is\n"
             "used once to create a small service infobase.",
             row=1, column=0, columnspan=3, sticky="w", pady=(4, 14))

        ttk.Label(inner, text="Client", style="Card.TLabel")\
            .grid(row=2, column=0, sticky="w")
        self.var = tk.StringVar(value=current or "")
        ttk.Entry(inner, textvariable=self.var, width=62)\
            .grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(inner, text="Browse...", command=self._browse).grid(row=2, column=2)

        self.status = ttk.Label(inner, text="", style="CardMuted.TLabel")
        self.status.grid(row=3, column=0, columnspan=3, sticky="w", pady=(10, 0))

        btns = ttk.Frame(inner, style="Card.TFrame")
        btns.grid(row=4, column=0, columnspan=3, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="Auto-detect", command=self._detect).pack(side="left")
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(btns, text="Save", style="Accent.TButton",
                   command=self._save).pack(side="right")

        self._detect(quiet=True)
        self.centre()

    def _browse(self):
        picked = filedialog.askopenfilename(
            title="Select 1cv8c.exe",
            filetypes=[("1C:Enterprise client", "1cv8c.exe 1cv8.exe"), ("All files", "*.*")])
        if picked:
            self.var.set(picked)
            self._check()

    def _detect(self, quiet=False):
        if self.var.get().strip() and quiet:
            self._check()
            return
        found = find_onec_client()
        if found:
            self.var.set(found)
            self._check()
        elif not quiet:
            self.status.configure(text="No 1C installation found under Program Files.",
                                  style="CardWarn.TLabel")

    def _check(self):
        path = self.var.get().strip()
        if not path or not os.path.isfile(path):
            self.status.configure(text="Not found on disk.", style="CardWarn.TLabel")
            return
        try:
            designer_exe(path)
        except OneCError as exc:
            self.status.configure(text=str(exc).splitlines()[0], style="CardWarn.TLabel")
            return
        epf, dt = onec_assets()
        missing = [os.path.basename(x) for x in (epf, dt) if not os.path.isfile(x)]
        if missing:
            self.status.configure(text="Missing renderer asset(s): " + ", ".join(missing),
                                  style="CardWarn.TLabel")
        else:
            self.status.configure(text="Ready.", style="CardOk.TLabel")

    def _save(self):
        self.on_save(self.var.get().strip())
        self.destroy()


# ==========================================================================
#  Main window
# ==========================================================================

PAGES = (("folder", "Working folder"),
         ("registers", "Registers"),
         ("rules", "Modifications"),
         ("files", "Files"))


class App(ttk.Frame):

    def __init__(self, master, state=None):
        ttk.Frame.__init__(self, master, padding=0)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.folder = tk.StringVar()
        self.registers = {}
        self.per_file = {}
        self.rules = []
        self.opt_backup = tk.BooleanVar(value=True)
        self.opt_skip = tk.BooleanVar(value=True)
        self.opt_verify = tk.BooleanVar(value=True)
        self.opt_all_files = tk.BooleanVar(value=True)
        self.selected_files = set()      # only meaningful when opt_all_files is off
        self.eligible_files = set()      # files a modification would actually change
        self.msgq = queue.Queue()
        self.busy = False
        self.log_open = False
        self.settings = load_settings()
        self.onec_client = self.settings.get("onec_client") or ""
        global ONEC_BACKGROUND, ONEC_PERSISTENT
        ONEC_BACKGROUND = bool(self.settings.get("onec_background", True))
        ONEC_PERSISTENT = bool(self.settings.get("onec_persistent", True))

        self._build_menu()
        self._build_sidebar()
        self._build_content()
        self._build_bottom()
        self.show_page("folder")
        self._pump()
        widget, detail = html_widget_status()
        self.say("Embedded HTML view: %s" % (detail if widget else "unavailable (%s)" % detail))
        self.say("Running from: %s%s"
                 % (app_dir(), "  (frozen .exe)" if getattr(sys, "frozen", False)
                    else "  (python %s)" % sys.version.split()[0]))
        self.say("1C renderer: %s"
                 % ("resident session (MxlToHtmlService.epf found)"
                    if os.path.isfile(service_epf())
                    else "one launch per render (MxlToHtmlService.epf not built yet)"))
        self.say()
        if state:
            self._restore(state)

    # ----------------------------------------------------- theme switching

    def _capture(self):
        return {"folder": self.folder.get(), "rules": list(self.rules),
                "page": next((k for k, v in self.nav.items() if v["active"]), "folder"),
                "backup": self.opt_backup.get(), "skip": self.opt_skip.get(),
                "verify": self.opt_verify.get(), "log": self.log_open,
                "all_files": self.opt_all_files.get(),
                "selected": set(self.selected_files)}

    def _restore(self, state):
        self.opt_backup.set(state.get("backup", True))
        self.opt_skip.set(state.get("skip", True))
        self.opt_verify.set(state.get("verify", True))
        self.opt_all_files.set(state.get("all_files", True))
        folder = state.get("folder", "")
        if folder:
            self.folder.set(folder)
            if os.path.isdir(folder):
                self.scan(quiet=True)
        # rules first: eligibility depends on them, and the selection is pruned to it
        for rule in state.get("rules", []):
            self.rules.append(rule)
            values, tags = self._rule_row(rule)
            self.rule_tree.insert("", "end", values=values, tags=tags)
        self.selected_files = set(state.get("selected", ()))
        self.refresh_file_list()
        self._refresh_badges()
        if state.get("log"):
            self.toggle_log()
        self.show_page(state.get("page", "folder"))

    def switch_theme(self, name):
        """Rebuild the window in the other palette, keeping folder and modifications."""
        if name == THEME_NAME:
            return
        state = self._capture()
        self.settings["theme"] = name
        save_settings(self.settings)
        root = self.master
        apply_theme(name)
        setup_style(root)
        self.destroy()
        App(root, state)

    # ------------------------------------------------------------- chrome

    def _build_menu(self):
        bar = tk.Menu(self.master)
        tools = tk.Menu(bar, tearoff=0)
        tools.add_command(label="1C:Enterprise renderer...", command=self.configure_onec)
        appearance = tk.Menu(tools, tearoff=0)
        self.theme_var = tk.StringVar(value=THEME_NAME)
        for value, label in (("dark", "Dark"), ("light", "Light")):
            appearance.add_radiobutton(label=label, value=value, variable=self.theme_var,
                                       command=lambda v=value: self.switch_theme(v))
        tools.add_cascade(label="Appearance", menu=appearance)
        self.opt_onec_background = tk.BooleanVar(
            value=bool(self.settings.get("onec_background", True)))
        tools.add_checkbutton(label="Run 1C minimised (keeps the preview in front)",
                              variable=self.opt_onec_background,
                              command=self._toggle_onec_background)
        self.opt_onec_persistent = tk.BooleanVar(
            value=bool(self.settings.get("onec_persistent", True)))
        tools.add_checkbutton(label="Keep 1C running between previews",
                              variable=self.opt_onec_persistent,
                              command=self._toggle_onec_persistent)
        tools.add_command(label="1C renderer diagnostics...",
                          command=self.onec_diagnostics)
        tools.add_separator()
        tools.add_command(label="Preview selected printform",
                          command=self.preview_selected_file)
        bar.add_cascade(label="Tools", menu=tools)
        self.master.configure(menu=bar)

    def _build_sidebar(self):
        side = tk.Frame(self, bg=SIDEBAR, width=212)
        side.grid(row=0, column=0, rowspan=2, sticky="nsw")
        side.grid_propagate(False)

        tk.Label(side, text="MXL Column Editor", bg=SIDEBAR, fg=TEXT,
                 font=(_FAMILY, 12, "bold"), anchor="w")\
            .pack(fill="x", padx=18, pady=(20, 0))
        tk.Label(side, text="1C printform templates", bg=SIDEBAR, fg=MUTED,
                 font=(_FAMILY, 9), anchor="w").pack(fill="x", padx=18, pady=(0, 18))

        self.nav = {}
        for key, label in PAGES:
            row = tk.Frame(side, bg=SIDEBAR, cursor="hand2")
            row.pack(fill="x")
            marker = tk.Frame(row, bg=SIDEBAR, width=3)
            marker.pack(side="left", fill="y")
            text = tk.Label(row, text=label, bg=SIDEBAR, fg=TEXT, anchor="w",
                            font=(_FAMILY, 10), padx=15, pady=9)
            text.pack(side="left", fill="x", expand=True)
            badge = tk.Label(row, text="", bg=SIDEBAR, fg=MUTED,
                             font=(_FAMILY, 9), padx=12)
            badge.pack(side="right")
            for widget in (row, text, badge):
                widget.bind("<Button-1>", lambda e, k=key: self.show_page(k))
                widget.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
                widget.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))
            self.nav[key] = {"row": row, "marker": marker, "text": text,
                             "badge": badge, "active": False}

        tk.Frame(side, bg=BORDER, height=1).pack(fill="x", padx=16, pady=16)
        self.side_stats = tk.Label(side, text="No folder selected", bg=SIDEBAR, fg=MUTED,
                                   font=(_FAMILY, 9), anchor="w", justify="left")
        self.side_stats.pack(fill="x", padx=18)

        foot = tk.Frame(side, bg=SIDEBAR)
        foot.pack(side="bottom", fill="x", pady=16, padx=18)
        self.onec_label = tk.Label(foot, text="", bg=SIDEBAR, fg=MUTED,
                                   font=(_FAMILY, 9), anchor="w", justify="left",
                                   cursor="hand2")
        self.onec_label.pack(fill="x")
        self.onec_label.bind("<Button-1>", lambda e: self.configure_onec())
        self._refresh_onec_label()

    def _nav_hover(self, key, entering):
        item = self.nav[key]
        if item["active"]:
            return
        colour = HOVER if entering else SIDEBAR
        item["row"].configure(bg=colour)
        item["text"].configure(bg=colour)
        item["badge"].configure(bg=colour)

    def show_page(self, key):
        for name, item in self.nav.items():
            active = name == key
            item["active"] = active
            bg = CARD if active else SIDEBAR
            item["row"].configure(bg=bg)
            item["text"].configure(bg=bg, fg=ACCENT if active else TEXT,
                                   font=(_FAMILY, 10, "bold" if active else "normal"))
            item["badge"].configure(bg=bg)
            item["marker"].configure(bg=ACCENT if active else bg)
        for name, frame in self.pages.items():
            frame.grid_remove()
        self.pages[key].grid(row=0, column=0, sticky="nsew")

    # -------------------------------------------------------------- pages

    def _build_content(self):
        holder = ttk.Frame(self, padding=(18, 18, 18, 8))
        holder.grid(row=0, column=1, sticky="nsew")
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.pages = {}
        self._page_folder(holder)
        self._page_registers(holder)
        self._page_rules(holder)
        self._page_files(holder)

    def _page(self, holder, title, subtitle):
        frame = ttk.Frame(holder)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)
        ttk.Label(frame, text=title, style="H1.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(frame, text=subtitle, style="Muted.TLabel")\
            .grid(row=1, column=0, sticky="w", pady=(2, 14))
        return frame

    def _page_folder(self, holder):
        frame = self._page(holder, "Working folder",
                           "Choose the folder holding the .mxl templates you want to change.")
        box = card(frame)
        box.grid(row=2, column=0, sticky="nsew")
        inner = ttk.Frame(box, style="Card.TFrame", padding=18)
        inner.pack(fill="both", expand=True)
        inner.columnconfigure(0, weight=1)

        ttk.Label(inner, text="Folder", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        row = ttk.Frame(inner, style="Card.TFrame")
        row.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        row.columnconfigure(0, weight=1)
        ttk.Entry(row, textvariable=self.folder).grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="Browse...", command=self.choose_folder)\
            .grid(row=0, column=1, padx=(8, 0))
        ttk.Button(row, text="Rescan", command=self.scan).grid(row=0, column=2, padx=(8, 0))

        self.folder_summary = ttk.Label(inner, text="Nothing scanned yet.",
                                        style="CardMuted.TLabel", justify="left")
        self.folder_summary.grid(row=2, column=0, sticky="w", pady=(16, 0))

        hint(inner,
             "Subfolders are not searched. Files are only rewritten when you press\n"
             "Apply changes, and a timestamped backup is taken first.",
             row=3, column=0, sticky="w", pady=(16, 0))
        self.pages["folder"] = frame

    def _page_registers(self, holder):
        frame = self._page(holder, "Registers",
                           "Every register found across the folder, and the columns it uses.")
        body = ttk.Frame(frame)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        left = card(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        li = ttk.Frame(left, style="Card.TFrame", padding=12)
        li.pack(fill="both", expand=True)
        li.columnconfigure(0, weight=1)
        li.rowconfigure(1, weight=1)
        ttk.Label(li, text="Registers", style="H2.TLabel").grid(row=0, column=0, sticky="w",
                                                                pady=(0, 8))
        self.reg_tree = ttk.Treeview(li, columns=("files", "cols"), selectmode="browse")
        self.reg_tree.heading("#0", text="Name")
        self.reg_tree.heading("files", text="Files")
        self.reg_tree.heading("cols", text="Columns")
        self.reg_tree.column("#0", width=250)
        self.reg_tree.column("files", width=60, anchor="e", stretch=False)
        self.reg_tree.column("cols", width=70, anchor="e", stretch=False)
        self.reg_tree.grid(row=1, column=0, sticky="nsew")
        rs = ttk.Scrollbar(li, orient="vertical", command=self.reg_tree.yview)
        rs.grid(row=1, column=1, sticky="ns")
        self.reg_tree.configure(yscrollcommand=rs.set)
        self.reg_tree.bind("<<TreeviewSelect>>", self.on_register_pick)
        self.reg_tree.bind("<Double-Button-1>", lambda e: self.add_rule_dialog())

        right = card(body)
        right.grid(row=0, column=1, sticky="nsew")
        ri = ttk.Frame(right, style="Card.TFrame", padding=12)
        ri.pack(fill="both", expand=True)
        ri.columnconfigure(0, weight=1)
        ri.rowconfigure(2, weight=1)
        ttk.Label(ri, text="Columns", style="H2.TLabel").grid(row=0, column=0, sticky="w")
        self.col_hint = hint(ri, "Select a register on the left.")
        self.col_hint.grid(row=1, column=0, sticky="w", pady=(2, 8))
        self.col_list = tk.Listbox(ri, exportselection=False, bd=0, highlightthickness=0,
                                   background=CARD, foreground=TEXT,
                                   selectbackground=ACCENT_LT, selectforeground=TEXT,
                                   activestyle="none", font=(_FAMILY, 10))
        self.col_list.grid(row=2, column=0, sticky="nsew")
        cs = ttk.Scrollbar(ri, orient="vertical", command=self.col_list.yview)
        cs.grid(row=2, column=1, sticky="ns")
        self.col_list.configure(yscrollcommand=cs.set)
        self.col_list.bind("<Double-Button-1>", self.on_column_activate)
        self.col_list.bind("<Return>", self.on_column_activate)
        ttk.Button(ri, text="Add modification here...",
                   command=self.add_rule_dialog).grid(row=3, column=0, sticky="w",
                                                      pady=(10, 0))
        self.pages["registers"] = frame

    def _page_files(self, holder):
        frame = self._page(holder, "Files",
                           "Every .mxl in the folder, and what your modifications would do to it.")
        box = card(frame)
        box.grid(row=2, column=0, sticky="nsew")
        inner = ttk.Frame(box, style="Card.TFrame", padding=12)
        inner.pack(fill="both", expand=True)
        inner.columnconfigure(0, weight=1)
        inner.rowconfigure(1, weight=1)

        top = ttk.Frame(inner, style="Card.TFrame")
        top.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Checkbutton(top, text="Apply to all files", variable=self.opt_all_files,
                        command=self._on_all_files_toggle).pack(side="left")
        self.btn_all = ttk.Button(top, text="Select all", command=lambda: self._bulk_select(True))
        self.btn_none = ttk.Button(top, text="Select none", command=lambda: self._bulk_select(False))
        self.sel_hint = hint(top, "")
        self.sel_hint.pack(side="left", padx=10)

        self.file_tree = ttk.Treeview(
            inner, columns=("sel", "size", "regs", "width", "status"),
            show="tree headings", selectmode="browse")
        self.file_tree.heading("#0", text="File")
        self.file_tree.heading("sel", text="Use")
        self.file_tree.heading("size", text="Size")
        self.file_tree.heading("regs", text="Registers")
        self.file_tree.heading("width", text="Widest")
        self.file_tree.heading("status", text="Effect of your modifications")
        self.file_tree.column("#0", width=270)
        self.file_tree.column("sel", width=44, anchor="center", stretch=False)
        self.file_tree.column("size", width=75, anchor="e", stretch=False)
        self.file_tree.column("regs", width=70, anchor="e", stretch=False)
        self.file_tree.column("width", width=65, anchor="e", stretch=False)
        self.file_tree.column("status", width=280)
        self.file_tree.grid(row=1, column=0, sticky="nsew")
        fs = ttk.Scrollbar(inner, orient="vertical", command=self.file_tree.yview)
        fs.grid(row=1, column=1, sticky="ns")
        self.file_tree.configure(yscrollcommand=fs.set)
        self.file_tree.tag_configure("add", foreground=OK_FG)
        self.file_tree.tag_configure("warn", foreground=WARN_FG)
        self.file_tree.tag_configure("idle", foreground=MUTED)
        self.file_tree.tag_configure("bad", foreground=ERR_FG)
        self.file_tree.bind("<Double-Button-1>", lambda e: self.preview_selected_file())
        self.file_tree.bind("<Button-1>", self._file_click)
        self.file_tree.bind("<Button-3>", self._file_menu_popup)     # Windows / Linux
        self.file_tree.bind("<Button-2>", self._file_menu_popup)     # macOS
        self.file_tree.bind("<space>", self._file_space)

        self.file_menu = tk.Menu(self, tearoff=0)
        self.file_menu.add_command(
            label="Include in the run",
            command=lambda: self._set_included(self._highlighted(), True))
        self.file_menu.add_command(
            label="Exclude from the run",
            command=lambda: self._set_included(self._highlighted(), False))
        self.file_menu.add_command(
            label="Toggle",
            command=lambda: self._toggle_included(self._highlighted()))
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Include every eligible file",
                                   command=lambda: self._bulk_select(True))
        self.file_menu.add_command(label="Include none",
                                   command=lambda: self._bulk_select(False))
        self.file_menu.add_separator()
        self.file_menu.add_command(label="Preview printform",
                                   command=self.preview_selected_file)
        self.file_menu.add_command(label="Open render in browser",
                                   command=self.exact_preview_browser)

        bar = ttk.Frame(inner, style="Card.TFrame")
        bar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(bar, text="Preview printform", style="Accent.TButton",
                   command=self.preview_selected_file).pack(side="left")
        hint(bar, "Double-click a file for the built-in grid. The exact preview renders "
                  "through 1C and opens in your browser.").pack(side="left", padx=6)
        self.pages["files"] = frame

    def _page_rules(self, holder):
        frame = self._page(holder, "Modifications",
                           "Each entry adds, renames or removes one column in one register.")
        body = ttk.Frame(frame)
        body.grid(row=2, column=0, sticky="nsew")
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        box = card(body)
        box.grid(row=0, column=0, sticky="nsew")
        inner = ttk.Frame(box, style="Card.TFrame", padding=12)
        inner.pack(fill="both", expand=True)
        inner.columnconfigure(0, weight=1)
        inner.rowconfigure(1, weight=1)

        head = ttk.Frame(inner, style="Card.TFrame")
        head.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Button(head, text="+  Add modification", style="Accent.TButton",
                   command=self.add_rule_dialog).pack(side="left")
        ttk.Button(head, text="Remove", command=self.remove_rule).pack(side="left", padx=8)
        ttk.Button(head, text="Clear all", command=self.clear_rules).pack(side="left")

        self.rule_tree = ttk.Treeview(
            inner, columns=("act", "reg", "pos", "anchor", "new", "val"),
            show="headings", selectmode="browse")
        for key, label, width in (("act", "Action", 70), ("reg", "Register", 185),
                                  ("pos", "Insert", 62),
                                  ("anchor", "Anchor / column", 170),
                                  ("new", "New column", 215), ("val", "Value", 70)):
            self.rule_tree.heading(key, text=label)
            self.rule_tree.column(key, width=width)
        self.rule_tree.grid(row=1, column=0, sticky="nsew")
        self.rule_tree.tag_configure("del", foreground=WARN_FG)
        self.rule_tree.tag_configure("ren", foreground=ACCENT)
        ms = ttk.Scrollbar(inner, orient="vertical", command=self.rule_tree.yview)
        ms.grid(row=1, column=1, sticky="ns")
        self.rule_tree.configure(yscrollcommand=ms.set)

        self.rules_empty = hint(
            inner, "No modifications yet.  Press  +  Add modification  to add, "
                   "rename or delete a column.")
        self.rules_empty.grid(row=2, column=0, sticky="w", pady=(10, 0))

        opts = ttk.Frame(inner, style="Card.TFrame")
        opts.grid(row=3, column=0, columnspan=2, sticky="w", pady=(14, 0))
        ttk.Checkbutton(opts, text="Back up originals",
                        variable=self.opt_backup).pack(side="left")
        ttk.Checkbutton(opts, text="Skip if column already exists",
                        variable=self.opt_skip).pack(side="left", padx=16)
        ttk.Checkbutton(opts, text="Verify after writing",
                        variable=self.opt_verify).pack(side="left")
        self.pages["rules"] = frame

    # ------------------------------------------------------- bottom chrome

    def _build_bottom(self):
        wrap = tk.Frame(self, bg=BG)
        wrap.grid(row=1, column=1, sticky="ew", padx=18, pady=(0, 14))
        wrap.columnconfigure(0, weight=1)

        self.log_box = card(wrap)
        self.log_box.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.log_box.grid_remove()
        li = ttk.Frame(self.log_box, style="Card.TFrame", padding=10)
        li.pack(fill="both", expand=True)
        li.columnconfigure(0, weight=1)
        self.log = tk.Text(li, height=12, wrap="none", bd=0, highlightthickness=0,
                           background=LOGBG, foreground=TEXT,
                           font=("TkFixedFont", 9), padx=8, pady=6)
        self.log.grid(row=0, column=0, sticky="nsew")
        ls = ttk.Scrollbar(li, orient="vertical", command=self.log.yview)
        ls.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=ls.set, state="disabled")

        bar = tk.Frame(wrap, bg=BG)
        bar.grid(row=1, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)

        self.log_toggle = tk.Label(bar, text="▸  Log", bg=BG, fg=MUTED,
                                   font=(_FAMILY, 9), cursor="hand2")
        self.log_toggle.grid(row=0, column=0, sticky="w")
        self.log_toggle.bind("<Button-1>", lambda e: self.toggle_log())

        mid = tk.Frame(bar, bg=BG)
        mid.grid(row=0, column=1, sticky="ew", padx=16)
        mid.columnconfigure(0, weight=1)
        self.status = tk.Label(mid, text="Choose a working folder to begin.", bg=BG,
                               fg=MUTED, font=(_FAMILY, 9), anchor="w")
        self.status.grid(row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(mid, style="Thin.Horizontal.TProgressbar",
                                        mode="determinate")
        self.progress.grid(row=1, column=0, sticky="ew", pady=(4, 0))
        self.progress.grid_remove()

        self.btn_apply = ttk.Button(bar, text="Apply changes", style="Accent.TButton",
                                    command=self.do_apply)
        self.btn_apply.grid(row=0, column=2)

    def toggle_log(self):
        self.log_open = not self.log_open
        if self.log_open:
            self.log_box.grid()
            self.log_toggle.configure(text="▾  Log")
        else:
            self.log_box.grid_remove()
            self.log_toggle.configure(text="▸  Log")

    # ------------------------------------------------------------ plumbing

    def say(self, line=""):
        self.msgq.put(("log", line))

    def set_status(self, text):
        self.msgq.put(("status", text))

    def _pump(self):
        try:
            while True:
                message = self.msgq.get_nowait()
                kind = message[0]
                if kind == "log":
                    self.log.configure(state="normal")
                    self.log.insert("end", message[1] + "\n")
                    self.log.see("end")
                    self.log.configure(state="disabled")
                elif kind == "status":
                    self.status.configure(text=message[1])
                elif kind == "progress":
                    done, total = message[1], message[2]
                    if total:
                        self.progress.grid()
                        self.progress.configure(maximum=total, value=done)
                    else:
                        self.progress.grid_remove()
                elif kind == "busy":
                    self._lock(message[1])
                elif kind == "rescan":
                    self.scan(quiet=True)
                    if getattr(self, "_verdict", None):
                        # queue it so it lands after scan()'s own status message
                        self.set_status(self._verdict)
        except queue.Empty:
            pass
        self.after(80, self._pump)

    def _lock(self, on):
        self.busy = on
        state = "disabled" if on else "normal"
        self.btn_apply.configure(state=state)

    def _refresh_onec_label(self):
        client = self.onec_client or find_onec_client()
        if client and os.path.isfile(client):
            self.onec_label.configure(text="1C renderer: ready\n" + os.path.basename(client),
                                      fg=OK_FG)
        else:
            self.onec_label.configure(text="1C renderer: not set\nclick to configure",
                                      fg=MUTED)

    def _refresh_badges(self):
        files = len(self.per_file)
        self.nav["registers"]["badge"].configure(text=str(len(self.registers)) if self.registers else "")
        if getattr(self, "opt_all_files", None) is not None and not self.opt_all_files.get():
            self.nav["files"]["badge"].configure(
                text="%d/%d" % (len(self.selected_files), len(self.eligible_files)))
        else:
            self.nav["files"]["badge"].configure(text=str(files) if files else "")
        self.nav["rules"]["badge"].configure(text=str(len(self.rules)) if self.rules else "")
        if not self.per_file:
            self.side_stats.configure(text="No folder selected")
        else:
            self.side_stats.configure(
                text="%d file%s\n%d register%s\n%d modification%s"
                     % (files, "" if files == 1 else "s",
                        len(self.registers), "" if len(self.registers) == 1 else "s",
                        len(self.rules), "" if len(self.rules) == 1 else "s"))
        self.rules_empty.grid() if not self.rules else self.rules_empty.grid_remove()

    # ------------------------------------------------------------- actions

    def _toggle_onec_background(self):
        global ONEC_BACKGROUND
        ONEC_BACKGROUND = bool(self.opt_onec_background.get())
        self.settings["onec_background"] = ONEC_BACKGROUND
        save_settings(self.settings)
        self.say("1C will start %s."
                 % ("minimised and unfocused" if ONEC_BACKGROUND else "normally"))

    def onec_diagnostics(self):
        """Walk the whole renderer path step by step and report where it stops."""
        threading.Thread(target=self._diagnostics_worker, daemon=True).start()

    def _diagnostic_sample(self):
        """Any .mxl we can render as an end-to-end check, wherever it lives."""
        selected = self.file_tree.selection() if hasattr(self, "file_tree") else None
        folder = self.folder.get().strip()
        if selected:
            candidate = os.path.join(folder, self.file_tree.item(selected[0], "text"))
            if os.path.isfile(candidate):
                return candidate
        for name in sorted(self.per_file):
            candidate = os.path.join(folder, name)
            if os.path.isfile(candidate):
                return candidate
        places = [folder,
                  self.settings.get("last_folder", ""),
                  app_dir()]
        for place in places:
            if not place or not os.path.isdir(place):
                continue
            found = sorted(glob.glob(os.path.join(place, "*.mxl")))
            if found:
                return found[0]
        return None

    def _diagnostics_worker(self):
        if not self.log_open:
            self.after(0, self.toggle_log)
        say = self.say
        say("")
        say("=" * 72)
        say("1C RENDERER DIAGNOSTICS")
        say("=" * 72)

        client = self.onec_client or find_onec_client()
        say("client exe          : %s" % (client or "NOT FOUND"))
        if client:
            say("   exists           : %s" % os.path.isfile(client))
            try:
                say("   designer beside  : %s" % designer_exe(client))
            except OneCError as exc:
                say("   designer beside  : %s" % exc)
        one_shot, service = onec_assets()[0], service_epf()
        say("one-shot processor  : %s (%s)"
            % (one_shot, "present" if os.path.isfile(one_shot) else "MISSING"))
        say("resident processor  : %s (%s)"
            % (service, "present" if os.path.isfile(service) else "MISSING"))
        say("persistent enabled  : %s" % ONEC_PERSISTENT)
        say("run minimised       : %s" % ONEC_BACKGROUND)
        say("ready timeout       : %ds" % ONEC_READY_TIMEOUT)

        if not client or not os.path.isfile(client):
            say("Cannot go further without a client. Tools > 1C:Enterprise renderer.")
            return
        if not os.path.isfile(service):
            say("No resident processor, so previews use one launch each. "
                "Build MxlToHtmlService.epf to enable the resident session.")
            return

        infobase = default_infobase(client)
        say("infobase            : %s" % infobase)
        try:
            ensure_infobase(client, infobase, say)
            say("   infobase ready   : yes")
        except OneCError as exc:
            say("   infobase FAILED  : %s" % exc)
            return

        say("")
        say("Starting a throwaway resident session...")
        _WORKER_REFUSED.discard((client, infobase))
        worker = OneCWorker(client, infobase, say)
        try:
            started = worker.start()
        except OneCError as exc:
            say("   start raised     : %s" % exc)
            return
        if not started:
            say("RESULT: the resident processor never reported ready.")
            say("        The details above come from 1C itself.")
            return

        say("   session folder   : %s" % worker.folder)
        say("   contents         : %s" % ", ".join(sorted(os.listdir(worker.folder))))
        trace = _text_of(os.path.join(worker.folder, "startup.log"))
        if trace:
            say("   processor trace  :")
            for line in trace.splitlines():
                say("      " + line)

        target = self._diagnostic_sample()
        if target is None:
            say("")
            say("The resident renderer started correctly, but no .mxl file was")
            say("found to render as a final check. Choose a folder on the Working")
            say("folder page and run this again for an end-to-end test.")
            worker.stop()
            say("Throwaway session stopped.")
            say("=" * 72)
            return

        out = os.path.join(tempfile.gettempdir(), "mxl_diagnostic.html")
        say("")
        say("Rendering %s ..." % target)
        try:
            worker.render([("diagnostic", target, out)])
            size = os.path.getsize(out) if os.path.isfile(out) else 0
            say("RESULT: rendered %d bytes to %s" % (size, out))
            say("        The resident renderer works.")
        except OneCError as exc:
            say("RESULT: render FAILED - %s" % exc)
            say("        session folder still holds: %s"
                % ", ".join(sorted(os.listdir(worker.folder))))
        finally:
            worker.stop()
            say("Throwaway session stopped.")
        say("=" * 72)

    def _toggle_onec_persistent(self):
        global ONEC_PERSISTENT
        ONEC_PERSISTENT = bool(self.opt_onec_persistent.get())
        self.settings["onec_persistent"] = ONEC_PERSISTENT
        save_settings(self.settings)
        if not ONEC_PERSISTENT:
            shutdown_workers()
        self.say("1C will %s between previews."
                 % ("stay running" if ONEC_PERSISTENT else "start and stop"))

    def configure_onec(self):
        def apply(path):
            self.onec_client = path
            self.settings["onec_client"] = path
            self._refresh_onec_label()
            if save_settings(self.settings):
                self.say("1C client set to: %s" % (path or "(none)"))
            else:
                self.say("1C client set for this session (settings file not writable).")
        OneCSettingsDialog(self, self.onec_client or find_onec_client() or "", apply)

    def choose_folder(self):
        picked = filedialog.askdirectory(title="Select the folder containing .mxl files")
        if picked:
            self.folder.set(picked)
            self.scan()

    def scan(self, quiet=False):
        folder = self.folder.get().strip()
        if not os.path.isdir(folder):
            if not quiet:
                messagebox.showerror(APP_TITLE, "Please choose a valid folder first.")
            return
        files = glob.glob(os.path.join(folder, "*.mxl"))
        if not files:
            if not quiet:
                messagebox.showwarning(APP_TITLE, "No .mxl files found in that folder.")
            return
        self.set_status("Scanning %d file(s)..." % len(files))
        self.say("Scanning %d file(s) in %s" % (len(files), folder))
        self.per_file, self.registers, errors = scan_folder(folder)
        for name, err in errors:
            self.say("  ! %s -- %s" % (name, err))

        self.reg_tree.delete(*self.reg_tree.get_children())
        for name in sorted(self.registers):
            info = self.registers[name]
            self.reg_tree.insert("", "end", text=name,
                                 values=(len(info["files"]), len(info["columns"])))
        self.refresh_file_list()
        self._refresh_badges()

        summary = "%d file(s) readable, %d distinct register(s)." % (
            len(self.per_file), len(self.registers))
        if errors:
            summary += "  %d file(s) could not be parsed." % len(errors)
        self.settings["last_folder"] = folder
        save_settings(self.settings)
        self.folder_summary.configure(text=summary)
        self.set_status(summary)
        self.say(summary)
        self.say()
        if not quiet and self.registers:
            self.show_page("registers")

    def refresh_file_list(self):
        """Redraw the Files page, including which files a run would actually touch."""
        if not hasattr(self, "file_tree"):
            return
        highlighted = set(self._highlighted())
        self.file_tree.delete(*self.file_tree.get_children())
        self.eligible_files = set()
        items = {}
        folder = self.folder.get().strip()
        for path in sorted(glob.glob(os.path.join(folder, "*.mxl"))):
            name = os.path.basename(path)
            tabs = self.per_file.get(name)
            try:
                size = "%.1f KB" % (os.path.getsize(path) / 1024.0)
            except OSError:
                size = "?"
            if tabs is None:
                items[name] = self.file_tree.insert(
                    "", "end", text=name,
                    values=("-", size, "-", "-", "could not be parsed"), tags=("bad",))
                continue
            widest = max([len(t.columns) for t in tabs] or [0])
            text, tag = self._effect_on(tabs)
            eligible = tag == "add"
            if eligible:
                self.eligible_files.add(name)
            items[name] = self.file_tree.insert(
                "", "end", text=name,
                values=(self._tick(name, eligible), size, len(tabs), widest, text),
                tags=(tag,))
        # a file that stopped being eligible cannot stay included
        self.selected_files &= self.eligible_files
        keep = [items[n] for n in highlighted if n in items]
        if keep:
            self.file_tree.selection_set(keep)
        self._sync_selection_controls()

    def _tick(self, name, eligible):
        if not eligible:
            return "\u2013"                      # en dash: nothing to do here
        if self.opt_all_files.get():
            return "\u2713"
        return "\u2713" if name in self.selected_files else "\u2610"

    def _sync_selection_controls(self):
        if not hasattr(self, "sel_hint"):
            return
        total = len(self.eligible_files)
        if self.opt_all_files.get():
            self.btn_all.pack_forget()
            self.btn_none.pack_forget()
            self.sel_hint.configure(
                text="%d file(s) would be changed. Untick to pick files yourself."
                     % total)
        else:
            self.btn_all.pack(side="left", padx=(12, 6))
            self.btn_none.pack(side="left")
            self.sel_hint.configure(
                text="%d of %d eligible file(s) included.  Ctrl/Shift-click to pick "
                     "several, then right-click or press Space."
                     % (len(self.selected_files), total))
        try:
            self.file_tree.configure(
                selectmode="extended" if not self.opt_all_files.get() else "browse")
        except tk.TclError:
            pass
        self._refresh_badges()

    def _highlighted(self):
        """File names currently highlighted in the tree (Ctrl/Shift multi-select)."""
        return [self.file_tree.item(i, "text") for i in self.file_tree.selection()]

    def _apply_ticks(self, names=None):
        """Repaint the Use cells in place, so the highlight survives."""
        for item in self.file_tree.get_children():
            name = self.file_tree.item(item, "text")
            if names is not None and name not in names:
                continue
            if name in self.per_file:
                self.file_tree.set(item, "sel",
                                   self._tick(name, name in self.eligible_files))
        self._sync_selection_controls()

    def _set_included(self, names, included):
        if self.opt_all_files.get():
            return
        touched = [n for n in names if n in self.eligible_files]
        if not touched:
            return
        for name in touched:
            if included:
                self.selected_files.add(name)
            else:
                self.selected_files.discard(name)
        self._apply_ticks(set(touched))

    def _toggle_included(self, names):
        if self.opt_all_files.get():
            return
        touched = [n for n in names if n in self.eligible_files]
        for name in touched:
            if name in self.selected_files:
                self.selected_files.discard(name)
            else:
                self.selected_files.add(name)
        if touched:
            self._apply_ticks(set(touched))

    def _file_space(self, _event=None):
        self._toggle_included(self._highlighted())
        return "break"

    def _file_menu_popup(self, event):
        item = self.file_tree.identify_row(event.y)
        if item and item not in self.file_tree.selection():
            self.file_tree.selection_set(item)     # right-click outside the selection
        if item:
            self.file_tree.focus(item)
        manual = not self.opt_all_files.get()
        for index in (0, 1, 2, 4, 5):
            self.file_menu.entryconfigure(
                index, state=("normal" if manual else "disabled"))
        try:
            self.file_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.file_menu.grab_release()
        return "break"

    def _on_all_files_toggle(self):
        if not self.opt_all_files.get():
            self.selected_files = set()      # start from nothing; pick files deliberately
        self.refresh_file_list()

    def _bulk_select(self, everything):
        if self.opt_all_files.get():
            return
        self.selected_files = set(self.eligible_files) if everything else set()
        self._apply_ticks()

    def _file_click(self, event):
        """Toggle a file when its Use cell is clicked (manual mode only)."""
        if self.opt_all_files.get():
            return
        try:
            if self.file_tree.identify_region(event.x, event.y) != "cell":
                return
            if self.file_tree.identify_column(event.x) != "#1":
                return
            item = self.file_tree.identify_row(event.y)
        except tk.TclError:
            return
        if not item:
            return
        name = self.file_tree.item(item, "text")
        if name not in self.eligible_files:
            return
        self._toggle_included([name])

    def files_to_process(self):
        """Paths a run should visit, honouring the Apply-to-all switch."""
        folder = self.folder.get().strip()
        every = sorted(glob.glob(os.path.join(folder, "*.mxl")))
        if self.opt_all_files.get():
            return every
        return [p for p in every if os.path.basename(p) in self.selected_files]

    def _effect_on(self, tabs):
        if not self.rules:
            return ("-", "idle")
        added = removed = renamed = already = missing = 0
        for rule in self.rules:
            for t in [x for x in tabs if x.name == rule.register]:
                if rule.is_delete:
                    if rule.anchor in t.columns:
                        removed += 1
                    else:
                        missing += 1
                elif rule.is_rename:
                    if rule.anchor in t.columns and rule.new_col not in t.columns:
                        renamed += 1
                    elif rule.new_col in t.columns:
                        already += 1
                    else:
                        missing += 1
                elif rule.new_col in t.columns:
                    already += 1
                elif rule.anchor in t.columns:
                    added += 1
                else:
                    missing += 1
        bits, tag = [], "idle"
        if added:
            bits.append("%d column(s) would be added" % added)
            tag = "add"
        if removed:
            bits.append("%d column(s) would be removed" % removed)
            tag = "add"
        if renamed:
            bits.append("%d column(s) would be renamed" % renamed)
            tag = "add"
        if already:
            bits.append("%d already present" % already)
        if missing:
            bits.append("%d column not found" % missing)
            if not (added or removed or renamed):
                tag = "warn"
        return (", ".join(bits) if bits else "register not present", tag)

    def on_register_pick(self, _event=None):
        sel = self.reg_tree.selection()
        if not sel:
            return
        name = self.reg_tree.item(sel[0], "text")
        cols = self.registers.get(name, {}).get("columns", [])
        self.col_list.delete(0, "end")
        for c in cols:
            self.col_list.insert("end", c)
        self.col_hint.configure(
            text=("%d column(s) in %s.  Double-click one to build a modification "
                  "around it." % (len(cols), name)) if cols else
                 "This register has no named columns, so it cannot be used as an anchor.")

    def _selected_register(self):
        sel = self.reg_tree.selection()
        return self.reg_tree.item(sel[0], "text") if sel else None

    def on_column_activate(self, _event=None):
        """Double-click a column: straight to Add modification, anchored on it."""
        selection = self.col_list.curselection()
        if not selection:
            return
        column = self.col_list.get(selection[0])
        register = self._selected_register()
        if not register or not column:
            return
        self.add_rule_dialog(register=register, anchor=column)

    def add_rule_dialog(self, register=None, anchor=None):
        if not self.registers:
            messagebox.showinfo(APP_TITLE, "Scan a folder first so the register names "
                                           "can be offered.")
            return
        RuleDialog(self, self.registers, self._accept_rule,
                   register or self._selected_register(), anchor)

    @staticmethod
    def _rule_row(rule):
        """One Rule as (treeview values, tags)."""
        if rule.is_delete:
            return ("Delete", rule.register, "", rule.anchor, "", ""), ("del",)
        if rule.is_rename:
            return ("Rename", rule.register, "", rule.anchor, rule.new_col, ""), ("ren",)
        return ("Add", rule.register, rule.position, rule.anchor,
                rule.new_col, rule.value), ()

    def _accept_rule(self, rule):
        self.rules.append(rule)
        values, tags = self._rule_row(rule)
        self.rule_tree.insert("", "end", values=values, tags=tags)
        self.refresh_file_list()
        self._refresh_badges()
        self.say("+ " + rule.describe())
        self.set_status("%d modification(s) queued." % len(self.rules))
        self.show_page("rules")

    def remove_rule(self):
        sel = self.rule_tree.selection()
        if not sel:
            return
        index = self.rule_tree.index(sel[0])
        self.rule_tree.delete(sel[0])
        del self.rules[index]
        self.refresh_file_list()
        self._refresh_badges()

    def clear_rules(self):
        self.rule_tree.delete(*self.rule_tree.get_children())
        self.rules = []
        self.refresh_file_list()
        self._refresh_badges()

    def preview_selected_file(self, path=None):
        """Open the preview window; the 1C render starts straight away."""
        path = self._resolve_preview_path(path)
        if path is None:
            return
        window = PreviewWindow(self, path, self.rules, self.opt_skip.get(),
                               on_exact=self.exact_preview_browser)
        if html_widget_class() is None:
            self.say("Embedded HTML view unavailable (%s); the 1C preview tab "
                     "explains how to enable it." % html_widget_status()[1])
        else:
            window.show_rendered()
        return window

    def _resolve_preview_path(self, path):
        if path is not None:
            if not os.path.isfile(path):
                messagebox.showerror(APP_TITLE, "File no longer exists.")
                return None
            return path
        sel = self.file_tree.selection() if hasattr(self, "file_tree") else None
        if not sel:
            messagebox.showinfo(APP_TITLE, "Select a file on the Files page first.")
            return None
        candidate = os.path.join(self.folder.get().strip(),
                                 self.file_tree.item(sel[0], "text"))
        if not os.path.isfile(candidate):
            messagebox.showerror(APP_TITLE, "File no longer exists.")
            return None
        return candidate

    def exact_preview_browser(self, path=None, use_rules=True):
        """Always render through 1C and hand the HTML to the default browser."""
        path = self._resolve_preview_path(path)
        if path is None:
            return
        client = self.onec_client or find_onec_client()
        if not client or not os.path.isfile(client):
            if messagebox.askyesno(
                    APP_TITLE,
                    "The 1C:Enterprise client is not configured, so an exact preview "
                    "cannot be produced.\n\nConfigure it now?"):
                self.configure_onec()
            return
        self.onec_client = client
        threading.Thread(target=self._exact_worker,
                         args=(path, client, use_rules), daemon=True).start()

    def render_html_file(self, path, use_rules=True, log=None):
        """
        Render one document to HTML through 1C and return the file path.

        Blocking; call it off the UI thread. When `use_rules` is set the queued
        modifications are applied to a temporary copy first, so the file on disk
        is never touched.
        """
        client = self.onec_client or find_onec_client()
        if not client or not os.path.isfile(client):
            raise OneCError("The 1C:Enterprise client is not configured.\n"
                            "Set it under Tools > 1C:Enterprise renderer.")
        self.onec_client = client
        name = os.path.basename(path)
        work = tempfile.mkdtemp(prefix="mxl-exact-")
        try:
            source = path
            if self.rules and use_rules:
                staged = os.path.join(work, name)
                applied = write_modified_copy(path, self.rules,
                                              self.opt_skip.get(), staged)
                source = staged
                if log:
                    log("   %d modification(s) applied to a temporary copy; "
                        "the file on disk is untouched." % len(applied))
            out = os.path.join(tempfile.gettempdir(),
                               "mxl_preview_" + os.path.splitext(name)[0]
                               + ("" if use_rules else "_original") + ".html")
            render_mxl_to_html(source, out, client, log=log)
            return out
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def _exact_worker(self, path, client, use_rules=True):
        name = os.path.basename(path)
        self.set_status("Rendering %s through 1C..." % name)
        self.say("Exact preview: %s" % name)
        try:
            keep = self.render_html_file(path, use_rules, log=self.say)
            webbrowser.open("file:///" + keep.replace("\\", "/"))
            self.say("   opened in your browser: %s" % keep)
            self.set_status("Exact preview of %s opened in your browser." % name)
        except OneCError as exc:
            self.say("   1C renderer failed: %s" % exc)
            self.set_status("1C renderer failed - see the log.")
            self.after(0, lambda: messagebox.showerror(
                APP_TITLE, "The 1C renderer could not produce a preview.\n\n%s" % exc))
        except Exception as exc:
            self.say("   unexpected error: %s" % exc)
            self.set_status("Exact preview failed - see the log.")

    # ------------------------------------------------------------ the run

    def do_apply(self):
        if not self.rules:
            messagebox.showerror(APP_TITLE, "Add at least one modification first.")
            return
        targets = self.files_to_process()
        if not targets:
            messagebox.showerror(APP_TITLE, "No files are selected on the Files page.")
            return
        scope = ("all %d file(s) in the folder" % len(targets) if self.opt_all_files.get()
                 else "the %d file(s) you selected" % len(targets))
        if not messagebox.askyesno(
                APP_TITLE,
                "Apply %d modification(s) to %s?\n\nThis rewrites the files in place."
                % (len(self.rules), scope)):
            return
        self._start_job(write=True)

    def _start_job(self, write):
        folder = self.folder.get().strip()
        if not os.path.isdir(folder):
            messagebox.showerror(APP_TITLE, "Please choose a valid folder first.")
            return
        if not self.rules:
            messagebox.showerror(APP_TITLE, "Add at least one modification first.")
            return
        if not self.files_to_process():
            messagebox.showerror(APP_TITLE, "No files are selected on the Files page.")
            return
        if self.busy:
            return
        if not self.log_open:
            self.toggle_log()
        self._lock(True)
        threading.Thread(target=self._worker, args=(folder, write), daemon=True).start()

    def _worker(self, folder, write):
        try:
            self._work(folder, write)
        except Exception:
            self.say("UNEXPECTED ERROR:")
            for line in traceback.format_exc().splitlines():
                self.say("   " + line)
            self.set_status("Failed - see the log.")
        finally:
            self.msgq.put(("progress", 0, 0))
            self.msgq.put(("busy", False))

    def _work(self, folder, write):
        self._verdict = None
        files = self.files_to_process()
        skip = self.opt_skip.get()
        mode = "APPLYING CHANGES" if write else "PREVIEW -- nothing will be written"
        scope = "every file in the folder" if self.opt_all_files.get() else "selected files only"
        self.say("=" * 72)
        self.say("%s   (%d file(s), %s, %d modification(s))"
                 % (mode, len(files), scope, len(self.rules)))
        for rule in self.rules:
            self.say("   " + rule.describe())
        self.say("=" * 72)

        backup_dir = None
        if write and self.opt_backup.get():
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = os.path.join(folder, "_mxl_backup_" + stamp)
            os.makedirs(backup_dir, exist_ok=True)
            self.set_status("Backing up %d file(s)..." % len(files))
            for path in files:
                shutil.copy2(path, backup_dir)
            self.say("Backup of %d file(s): %s" % (len(files), backup_dir))
            self.say()

        changed = skipped = failed = 0
        counts = {}
        for index, path in enumerate(files, 1):
            name = os.path.basename(path)
            self.msgq.put(("progress", index, len(files)))
            self.set_status("%s %d of %d: %s"
                            % ("Writing" if write else "Checking", index, len(files), name))
            try:
                report = process_file(path, self.rules, skip_existing=skip, write=write)
            except Exception as exc:
                failed += 1
                self.say("  ERROR  %-40s %s" % (name, exc))
                continue
            lines, touched = [], False
            for rule, outcomes in report:
                for status, detail in outcomes:
                    key = (rule.register, status if status != "ok" else "modified")
                    counts[key] = counts.get(key, 0) + 1
                    if status == "ok":
                        touched = True
                    lines.append("       %-24s %-8s %s" % (rule.register[:24], status, detail))
            if touched:
                changed += 1
                self.say("  %-42s %s" % (name, "written" if write else "would change"))
            else:
                skipped += 1
            for line in lines:
                self.say(line)

        self.say()
        self.say("-" * 72)
        self.say("Files %s: %d   unchanged: %d   errors: %d"
                 % ("written" if write else "that would change", changed, skipped, failed))
        for (reg, status), n in sorted(counts.items()):
            self.say("   %-30s %-28s %d table(s)" % (reg[:30], status, n))

        verdict = "%s: %d file(s) %s, %d unchanged" % (
            "Applied" if write else "Preview", changed,
            "written" if write else "would change", skipped)
        if failed:
            verdict += ", %d error(s)" % failed

        if write and self.opt_verify.get() and backup_dir:
            self.say()
            self.say("Verifying...")
            self.set_status("Verifying %d file(s)..." % len(files))
            bad = 0
            for index, path in enumerate(files, 1):
                self.msgq.put(("progress", index, len(files)))
                name = os.path.basename(path)
                problems = verify_file(path, os.path.join(backup_dir, name), self.rules)
                if problems:
                    bad += 1
                    self.say("  FAIL  %s" % name)
                    for problem in problems:
                        self.say("          " + problem)
            self.say("Verification: %s (%d file(s) with problems)"
                     % ("PASSED" if bad == 0 else "FAILED", bad))
            verdict += ".  Verification " + ("passed" if bad == 0 else "FAILED")
            if bad:
                self.say("Originals are intact in: %s" % backup_dir)
        elif write and self.opt_verify.get():
            self.say("Verification needs backups enabled; skipped.")

        self._verdict = verdict
        self.set_status(verdict)
        self.say()
        if write:
            self.msgq.put(("rescan",))


def main():
    apply_theme(load_settings().get("theme", "dark"))
    root = tk.Tk()
    root.protocol("WM_DELETE_WINDOW", lambda: (shutdown_workers(), root.destroy()))
    root.title(APP_TITLE)
    root.geometry("1120x780")
    root.minsize(960, 640)
    setup_style(root)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
