"""Heuristic VBA lint for the vba/ sources (no Excel available here).

With Option Explicit, any undeclared identifier is a fatal compile error on
the user's machine — this lint catches typos before they ship. Also checks
block balance (Sub/Function, If, For, Do, While, With, Select).

Not a real compiler: keep the whitelist below in sync when adding modules.
"""
import re
import sys
from pathlib import Path

VBA_KEYWORDS = {
    "if", "then", "else", "elseif", "end", "sub", "function", "exit", "dim",
    "const", "public", "private", "as", "for", "to", "step", "next", "do",
    "while", "until", "loop", "select", "case", "with", "set", "new", "not",
    "and", "or", "xor", "mod", "is", "in", "each", "goto", "on", "error",
    "resume", "true", "false", "nothing", "byval", "byref", "optional",
    "call", "let", "redim", "erase", "preserve", "wend", "like", "attribute",
    "option", "explicit", "static", "me", "vbcrlf", "vblf", "vbcr", "vbtab",
    "empty", "null", "paramarray", "boolean", "byte", "integer", "long",
    "single", "double", "currency", "string", "variant", "object", "date",
    "stop", "debug", "unload", "load",
}

BUILTINS = {
    # VBA runtime
    "msgbox", "inputbox", "format", "val", "cstr", "clng", "cint", "cdbl",
    "cdate", "cbool", "cvar", "ccur", "len", "left", "right", "mid", "instr",
    "instrrev", "replace", "trim", "ltrim", "rtrim", "ucase", "lcase",
    "split", "join", "array", "ubound", "lbound", "isnumeric", "isdate",
    "isarray", "isempty", "isnull", "ismissing", "isobject", "chr", "chrw",
    "asc", "ascw", "space", "strconv", "strcomp", "string", "abs", "int",
    "fix", "sgn", "sqr", "rnd", "randomize", "timer", "now", "time",
    "dateadd", "datediff", "datepart", "dateserial", "timeserial", "year",
    "month", "day", "hour", "minute", "second", "weekday", "environ",
    "createobject", "getobject", "typename", "vartype", "err", "iif",
    "filelen", "dir", "kill", "filecopy", "shell", "round",
    # VBA enums / constants
    "vbok", "vbcancel", "vbyes", "vbno", "vbokcancel", "vbyesno",
    "vbyesnocancel", "vbexclamation", "vbinformation", "vbcritical",
    "vbquestion", "vbboolean", "vbnarrow", "vbblack", "vbred", "vbwhite",
    # Excel objects / members commonly used unqualified
    "application", "worksheetfunction", "thisworkbook", "activesheet",
    "activewindow", "activecell", "sheets", "worksheets", "workbooks",
    "range", "cells", "rows", "columns", "selection", "union", "intersect",
    "rgb",
    # Excel enums
    "xlnone", "xltypepdf", "xlqualitystandard", "xlpapera4", "xlportrait",
    "xllandscape", "xlcenter", "xlright", "xlleft", "xlhairline", "xlthin",
    "xlmedium", "xledgeleft", "xledgeright", "xledgetop", "xledgebottom",
    "xlinsidevertical", "xlinsidehorizontal", "xlcontinuous", "xlvalue",
    "xlpastevalues", "xldown", "xlup", "xltoleft", "xltoright",
    "vbnullstring",
}

# Public symbols defined across this project's modules
PROJECT_PUBLIC = {
    # Module1 consts/vars
    "kcolor", "mcolor", "pcolor", "tcolor", "dcolor",
    "maxs", "maxslot", "state_col", "state_ver",
    "dd", "kk", "ss", "knumber", "fail", "choseisagyo", "nowsh",
    # Module1 procs
    "ensuredims", "btnsettings", "btnmenu", "btnimport", "btnquickinput",
    "btnsurveyprint", "sheetcopy", "komareset", "komacheck", "runchosei",
    "stateisconsistent", "repaintfromstate", "makeschedulecore",
    "makenotifysheet", "exportsheetpdf", "cellcolor", "myprev",
    "module1",
    # Module2 procs
    "importwishes", "makesurveysheets", "quickinput", "module2",
    # Forms
    "userform1", "userform2", "userform3", "userform4",
}

# Form controls (referenced unqualified inside their own form modules)
FORM_CONTROLS = {
    "combobox1", "combobox2", "combobox3", "combobox4", "combobox5",
    "combobox6", "commandbutton1", "commandbutton2", "commandbutton3",
    "commandbutton4", "commandbutton5", "commandbutton6", "checkbox1",
    "checkbox2", "textbox1", "label3", "optionbutton1", "optionbutton2",
    "optionbutton3", "optionbutton4", "optionbutton5", "optionbutton6",
}

IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def strip_strings_comments(line):
    out = []
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 1
                else:
                    in_str = False
            i += 1
            continue
        if ch == '"':
            in_str = True
            out.append(" ")
            i += 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    return "".join(out)


def logical_lines(text):
    """Join _ continuations; yield (first_lineno, joined)."""
    lines = text.split("\n")
    buf, start = "", 0
    for n, raw in enumerate(lines, 1):
        s = strip_strings_comments(raw.rstrip("\r"))
        if not buf:
            start = n
        if s.rstrip().endswith("_"):
            buf += s.rstrip()[:-1] + " "
            continue
        yield start, buf + s
        buf = ""
    if buf:
        yield start, buf


def lint_module(path):
    text = path.read_text(encoding="utf-8")
    errors = []
    declared = set()
    labels = set()
    block_stack = []

    is_form = path.suffix == ".frm"
    name = path.stem.lower()

    decl_re = re.compile(
        r"^\s*(?:public\s+|private\s+|static\s+)?(?:dim|const|redim)\s+(.*)$", re.I)
    proc_re = re.compile(
        r"^\s*(?:public\s+|private\s+|friend\s+)?(?:static\s+)?(sub|function|property\s+(?:get|let|set))\s+([A-Za-z_][A-Za-z0-9_]*)\s*(\(.*)?$", re.I)
    pubvar_re = re.compile(
        r"^\s*(?:public|private|global)\s+(?!sub|function|const|property)(.*)$", re.I)
    const_re = re.compile(
        r"^\s*(?:public\s+|private\s+)?const\s+(.*)$", re.I)
    label_re = re.compile(r"^\s*([A-Za-z_][A-Za-z0-9_]*):\s*$")
    forvar_re = re.compile(r"^\s*for\s+(?:each\s+)?([A-Za-z_][A-Za-z0-9_]*)", re.I)

    def add_decl_list(body):
        # "a As Long, b(1 To 5) As String, c"
        for part in split_top_commas(body):
            m = IDENT.match(part.strip())
            if m:
                declared.add(m.group(0).lower())

    def split_top_commas(s):
        depth = 0
        cur = ""
        for ch in s:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            if ch == "," and depth == 0:
                yield cur
                cur = ""
            else:
                cur += ch
        yield cur

    # pass 1: declarations
    for lineno, line in logical_lines(text):
        l = line.strip()
        if not l or l.lower().startswith(("attribute ", "option ")):
            continue
        m = label_re.match(l)
        if m:
            labels.add(m.group(1).lower())
            continue
        m = proc_re.match(l)
        if m:
            declared.add(m.group(2).lower())
            if m.group(3):
                params = m.group(3)
                params = params[params.find("(") + 1:]
                if params.endswith(")"):
                    params = params[:-1]
                # strip trailing "As Type" of function
                for part in split_top_commas(params):
                    part = re.sub(r"^\s*(byval|byref|optional|paramarray)\s+", "", part.strip(), flags=re.I)
                    part = re.sub(r"^\s*(byval|byref|optional|paramarray)\s+", "", part, flags=re.I)
                    mm = IDENT.match(part)
                    if mm:
                        declared.add(mm.group(0).lower())
            continue
        m = decl_re.match(l)
        if m:
            add_decl_list(m.group(1))
            continue
        m = const_re.match(l)
        if m:
            add_decl_list(m.group(1))
            continue
        m = pubvar_re.match(l)
        if m and re.match(r"^\s*(public|private|global)\s", l, re.I):
            # module-level var like "Public dd As Long, kk As Long"
            add_decl_list(m.group(1))
            continue

    # pass 2: usage + block balance
    known = (VBA_KEYWORDS | BUILTINS | PROJECT_PUBLIC | declared | labels |
             {name})
    if is_form:
        known |= FORM_CONTROLS

    for lineno, line in logical_lines(text):
        l = line.strip()
        low = l.lower()
        if not l or low.startswith(("attribute ", "option ")):
            continue

        # block tracking
        if re.match(r"^(public |private |friend )?(static )?(sub|function)\b", low):
            block_stack.append(("proc", lineno))
        elif re.match(r"^end (sub|function)\b", low):
            if not block_stack or block_stack[-1][0] != "proc":
                errors.append((lineno, "unbalanced End Sub/Function"))
            else:
                block_stack.pop()
        elif re.match(r"^if\b.*\bthen\s*$", low):
            block_stack.append(("if", lineno))
        elif re.match(r"^end if\b", low):
            if not block_stack or block_stack[-1][0] != "if":
                errors.append((lineno, "unbalanced End If"))
            else:
                block_stack.pop()
        elif re.match(r"^(for)\b", low):
            block_stack.append(("for", lineno))
        elif re.match(r"^next\b", low):
            if not block_stack or block_stack[-1][0] != "for":
                errors.append((lineno, "unbalanced Next"))
            else:
                block_stack.pop()
        elif re.match(r"^do\b", low):
            block_stack.append(("do", lineno))
        elif re.match(r"^loop\b", low):
            if not block_stack or block_stack[-1][0] != "do":
                errors.append((lineno, "unbalanced Loop"))
            else:
                block_stack.pop()
        elif re.match(r"^with\b", low):
            block_stack.append(("with", lineno))
        elif re.match(r"^end with\b", low):
            if not block_stack or block_stack[-1][0] != "with":
                errors.append((lineno, "unbalanced End With"))
            else:
                block_stack.pop()
        elif re.match(r"^select case\b", low):
            block_stack.append(("select", lineno))
        elif re.match(r"^end select\b", low):
            if not block_stack or block_stack[-1][0] != "select":
                errors.append((lineno, "unbalanced End Select"))
            else:
                block_stack.pop()

        # identifier check: skip declarations (already harvested)
        if re.match(r"^\s*(public|private|static|dim|const|redim)\b", low):
            continue
        # remove member accesses (.xxx) and named args (x:=)
        scrubbed = re.sub(r"\.\s*[A-Za-z_][A-Za-z0-9_]*", " ", l)
        scrubbed = re.sub(r"[A-Za-z_][A-Za-z0-9_]*\s*:=", " ", scrubbed)
        # remove hex literals
        scrubbed = re.sub(r"&H[0-9A-Fa-f]+", " ", scrubbed)
        for m in IDENT.finditer(scrubbed):
            ident = m.group(0).lower()
            if ident in known:
                continue
            if ident.startswith("vb") or ident.startswith("xl"):
                continue
            errors.append((lineno, f"undeclared identifier: {m.group(0)}"))

    for kind, lineno in block_stack:
        errors.append((lineno, f"unclosed block: {kind}"))

    return errors


def main():
    vba_dir = Path(__file__).resolve().parent.parent / "vba"
    total = 0
    for f in sorted(vba_dir.glob("*.*")):
        if f.suffix not in (".bas", ".cls", ".frm"):
            continue
        errs = lint_module(f)
        # dedupe
        seen = set()
        uniq = [e for e in errs if not (e in seen or seen.add(e))]
        if uniq:
            print(f"--- {f.name}")
            for lineno, msg in uniq:
                print(f"  L{lineno}: {msg}")
            total += len(uniq)
    if total:
        print(f"\n{total} issue(s)")
        sys.exit(1)
    print("lint OK")


if __name__ == "__main__":
    main()
