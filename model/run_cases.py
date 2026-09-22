"""
Run every MCASE scenario through the workbook in LibreOffice (headless, via UNO) so the
self-referencing case-capture cells on the Sensitivity tab are populated, then save the
recalculated workbook (formulas + cached values) back to disk.

This replaces the Excel "data table over MCASE" step: Excel users can instead set MCASE to
each case number in turn (iterative calculation enabled) and finish on the base case.

Each case is recalculated until the named outputs OUT_IRR / OUT_MOIC stop moving, because the
model is circular by design (average-balance interest, cash sweep, interest on cash).

Usage:
    python3 model/run_cases.py model/Project_Daylight_LBO_v1.xlsx [--cases 1-24] [--default 3]
"""

import argparse
import os
import subprocess
import sys
import tempfile
import time
from pathlib import Path

import uno  # LibreOffice Python-UNO bridge
from com.sun.star.beans import PropertyValue

sys.path.insert(0, "/mnt/skills/public/xlsx/scripts")
try:
    from office.soffice import get_soffice_env  # sandbox-safe LibreOffice environment
except Exception:  # pragma: no cover - outside the sandbox
    def get_soffice_env():
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "svp"
        return env


def _prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def start_office(port=2083):
    profile = tempfile.mkdtemp(prefix="lo_uno_profile_")
    cmd = [
        "soffice", "--headless", "--invisible", "--nologo", "--norestore", "--nodefault",
        f"-env:UserInstallation={Path(profile).as_uri()}",
        f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext",
    ]
    proc = subprocess.Popen(cmd, env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    last = None
    for _ in range(120):
        try:
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            return proc, ctx
        except Exception as exc:  # office still starting
            last = exc
            time.sleep(0.5)
    proc.kill()
    raise RuntimeError(f"Could not connect to LibreOffice: {last}")


def _named_value(doc, name):
    cells = doc.NamedRanges.getByName(name).getReferredCells()
    return cells.getCellByPosition(0, 0).getValue()


def converge(doc, probes=("OUT_IRR", "OUT_MOIC"), max_passes=40, tol=1e-10):
    prev, stable = None, 0
    for i in range(max_passes):
        doc.calculateAll()
        cur = [_named_value(doc, p) for p in probes]
        if prev is not None and max(abs(a - b) for a, b in zip(cur, prev)) < tol:
            stable += 1
            if stable >= 2:
                return i + 1
        else:
            stable = 0
        prev = cur
    return max_passes


def _read(doc, ref):
    sheet, addr = ref.split("!")
    return doc.Sheets.getByName(sheet).getCellRangeByName(addr).getValue()


def run(path, cases, default_case, readback=None, port=2083, verbose=True):
    path = str(Path(path).resolve())
    proc, ctx = start_office(port)
    results = {}
    desktop = None
    try:
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            Path(path).as_uri(), "_blank", 0,
            (_prop("Hidden", True), _prop("FilterName", "Calc MS Excel 2007 XML")))
        doc.IsIterationEnabled = True
        doc.IterationCount = 1000
        doc.IterationEpsilon = 1e-9
        mcase = doc.NamedRanges.getByName("MCASE").getReferredCells().getCellByPosition(0, 0)
        converge(doc)
        for n in cases:
            mcase.setValue(n)
            passes = converge(doc)
            if readback:
                results[n] = {k: _read(doc, ref) for k, ref in readback.items()}
            if verbose:
                print(f"case {n:>2}: {passes:>2} passes  IRR={_named_value(doc, 'OUT_IRR'):.4f}  "
                      f"MoIC={_named_value(doc, 'OUT_MOIC'):.3f}  checks={_named_value(doc, 'CHECKS'):.4f}")
        mcase.setValue(default_case)
        converge(doc)
        # second sweep so every capture cell holds its fully converged snapshot
        for n in cases:
            mcase.setValue(n)
            converge(doc)
        mcase.setValue(default_case)
        converge(doc)
        doc.storeToURL(Path(path).as_uri(),
                       (_prop("FilterName", "Calc MS Excel 2007 XML"), _prop("Overwrite", True)))
        doc.close(True)
    finally:
        try:
            if desktop is not None:
                desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()
    return results


def _parse_cases(spec):
    out = []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-")
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--cases", default="1-24")
    ap.add_argument("--default", type=int, default=3)
    args = ap.parse_args()
    run(args.xlsx, _parse_cases(args.cases), args.default)
    print(f"Ran cases {args.cases}; saved {args.xlsx} with MCASE={args.default}")
