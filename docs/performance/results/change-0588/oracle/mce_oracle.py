"""Differential oracle for change 0588.

For every .xlsx/.docx/.pptx fixture under a test-data tree, run the MCE codec
over every XML part with the BEFORE binary and the AFTER binary and compare a
namespace-resolving canonical form of the result: the resolved (namespace URI,
local name) of every element and attribute with its normalized value, every
text run, CDATA, comment, PI and declaration, and the processing Report; or, on
refusal, the error's Debug identity and Display message.

Namespace declarations themselves are deliberately excluded from the canonical
form: they are exactly what this change alters. Any name that stops resolving
shows up as `!UNKNOWN-PREFIX:<prefix>` and fails the comparison.

usage: mce_oracle.py <before-bin> <after-bin> <test-data-dir> <report.tsv>
"""

import concurrent.futures as cf
import os
import subprocess
import sys
import zipfile

BEFORE, AFTER, ROOT, REPORT = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
MODE = sys.argv[5] if len(sys.argv) > 5 else "canon"
SUFFIXES = (".xlsx", ".docx", ".pptx")
PART_SUFFIXES = (".xml", ".rels")


def canon(binary, data):
    proc = subprocess.run(
        [binary, MODE, "-"], input=data, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, timeout=600,
    )
    if proc.returncode != 0:
        return b"PROBE-FAILED\t%d\t%s" % (proc.returncode, proc.stderr[-400:])
    return proc.stdout


def check(fixture):
    rows, parts, refusals = [], 0, 0
    try:
        zf = zipfile.ZipFile(fixture)
    except Exception as error:  # noqa: BLE001 - fixture corpora include bad zips
        return [("SKIP-ZIP", fixture, "", str(error)[:120])], 0, 0
    with zf:
        for info in zf.infolist():
            if not info.filename.lower().endswith(PART_SUFFIXES):
                continue
            try:
                data = zf.read(info.filename)
            except Exception as error:  # noqa: BLE001
                rows.append(("SKIP-PART", fixture, info.filename, str(error)[:120]))
                continue
            parts += 1
            before, after = canon(BEFORE, data), canon(AFTER, data)
            if before.startswith(b"ERR\t"):
                refusals += 1
            if before != after:
                head_b = before.decode("utf-8", "replace").splitlines()
                head_a = after.decode("utf-8", "replace").splitlines()
                first = next(
                    (i for i, (x, y) in enumerate(zip(head_b, head_a)) if x != y),
                    min(len(head_b), len(head_a)),
                )
                rows.append((
                    "MISMATCH", fixture, info.filename,
                    "line %d | before=%r | after=%r"
                    % (first,
                       head_b[first] if first < len(head_b) else "<eof>",
                       head_a[first] if first < len(head_a) else "<eof>"),
                ))
    return rows, parts, refusals


def main():
    fixtures = []
    for base, _dirs, files in os.walk(ROOT):
        for name in files:
            if name.lower().endswith(SUFFIXES):
                fixtures.append(os.path.join(base, name))
    fixtures.sort()
    total_parts = total_refusals = 0
    findings = []
    with cf.ThreadPoolExecutor(max_workers=8) as pool:
        for rows, parts, refusals in pool.map(check, fixtures):
            findings.extend(rows)
            total_parts += parts
            total_refusals += refusals
    with open(REPORT, "w") as out:
        out.write("kind\tfixture\tpart\tdetail\n")
        for row in findings:
            out.write("\t".join(row) + "\n")
        out.write(
            "#summary\tfixtures=%d\tparts=%d\trefusals=%d\tmismatches=%d\tskipped=%d\n"
            % (len(fixtures), total_parts, total_refusals,
               sum(1 for r in findings if r[0] == "MISMATCH"),
               sum(1 for r in findings if r[0].startswith("SKIP"))))
    print("fixtures=%d parts=%d refusals=%d mismatches=%d skipped=%d"
          % (len(fixtures), total_parts, total_refusals,
             sum(1 for r in findings if r[0] == "MISMATCH"),
             sum(1 for r in findings if r[0].startswith("SKIP"))))


main()
