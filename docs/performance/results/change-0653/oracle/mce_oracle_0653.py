"""Corpus differential for change 0653.

For every .xlsx/.docx/.pptx fixture under a test-data tree, run the MCE codec
over every XML part with the BEFORE binary and the AFTER binary and compare:

  * `canon`: a namespace-resolving projection of the processed output -- the
    resolved (namespace URI, local name) of every element and attribute with
    its normalized value, every text run, CDATA, comment and declaration, with
    namespace declarations themselves excluded, plus the processing Report; or,
    on refusal, the error's Debug identity and Display message.  This is the
    property the namespace-emission rewrite must preserve exactly.
  * `raw`: the exact output length and the borrow-versus-own decision.  Byte
    identity is NOT expected here -- the rewrite is byte-visible by design --
    so the assertions are: the refusal identity is the same, the borrow
    decision is the same, and the new output is never longer than the old.

usage: mce_oracle_0653.py <before-bin> <after-bin> <test-data-dir> <report.tsv>
"""

import concurrent.futures as cf
import glob
import os
import subprocess
import sys
import zipfile

BEFORE, AFTER, ROOT, REPORT = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
SUFFIXES = (".xlsx", ".docx", ".pptx")
PART_SUFFIXES = (".xml", ".rels")


def run(binary, mode, data):
    proc = subprocess.run(
        [binary, mode, "-"], input=data, stdout=subprocess.PIPE,
        stderr=subprocess.PIPE, timeout=900,
    )
    if proc.returncode != 0:
        return b"PROBE-FAILED\t%d\t%s" % (proc.returncode, proc.stderr[-400:])
    return proc.stdout


def check(fixture):
    rows, parts, refusals, grew, shrank, before_bytes, after_bytes = [], 0, 0, 0, 0, 0, 0
    try:
        zf = zipfile.ZipFile(fixture)
    except Exception as error:  # noqa: BLE001 - fixture corpora include bad zips
        return [("SKIP-ZIP", fixture, "", str(error)[:120])], 0, 0, 0, 0, 0, 0
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
            cb, ca = run(BEFORE, "canon", data), run(AFTER, "canon", data)
            if cb != ca:
                lb = cb.decode("utf-8", "replace").splitlines()
                la = ca.decode("utf-8", "replace").splitlines()
                first = next((i for i, (x, y) in enumerate(zip(lb, la)) if x != y),
                             min(len(lb), len(la)))
                rows.append(("CANON-MISMATCH", fixture, info.filename,
                             "line %d | before=%r | after=%r" % (
                                 first,
                                 lb[first] if first < len(lb) else "<eof>",
                                 la[first] if first < len(la) else "<eof>")))
            rb = run(BEFORE, "raw", data).decode("utf-8", "replace").strip()
            ra = run(AFTER, "raw", data).decode("utf-8", "replace").strip()
            fb, fa = rb.split("\t"), ra.split("\t")
            if fb[0] == "ERR" or fa[0] == "ERR":
                refusals += 1
                if rb != ra:
                    rows.append(("REFUSAL-MISMATCH", fixture, info.filename,
                                 "before=%r after=%r" % (rb, ra)))
                continue
            # OK <len> <borrowed> <hash> <report>
            if fb[2] != fa[2]:
                rows.append(("BORROW-MISMATCH", fixture, info.filename,
                             "before=%s after=%s" % (fb[2], fa[2])))
            if fb[4] != fa[4]:
                rows.append(("REPORT-MISMATCH", fixture, info.filename,
                             "before=%s after=%s" % (fb[4], fa[4])))
            nb, na = int(fb[1]), int(fa[1])
            before_bytes += nb
            after_bytes += na
            if na > nb:
                grew += 1
                rows.append(("GREW", fixture, info.filename, "before=%d after=%d" % (nb, na)))
            elif na < nb:
                shrank += 1
    return rows, parts, refusals, grew, shrank, before_bytes, after_bytes


fixtures = sorted(
    path for suffix in SUFFIXES
    for path in glob.glob(os.path.join(ROOT, "**", "*" + suffix), recursive=True)
)
all_rows, parts, refusals, grew, shrank, bb, ab = [], 0, 0, 0, 0, 0, 0
# Threads, not processes: every unit of work is a blocking subprocess call, so
# the GIL is released for all of it and the pool needs no fork.
with cf.ThreadPoolExecutor(max_workers=int(os.environ.get("JOBS", "12"))) as pool:
    for rows, count, refused, g, s, b, a in pool.map(check, fixtures):
        all_rows.extend(rows)
        parts += count
        refusals += refused
        grew += g
        shrank += s
        bb += b
        ab += a

mismatches = [row for row in all_rows if row[0].endswith("MISMATCH") or row[0] == "GREW"]
with open(REPORT, "w") as handle:
    handle.write("fixtures=%d parts=%d refusals=%d mismatches=%d grew=%d shrank=%d "
                 "before_output_bytes=%d after_output_bytes=%d\n"
                 % (len(fixtures), parts, refusals, len(mismatches), grew, shrank, bb, ab))
    for row in all_rows:
        handle.write("\t".join(str(field) for field in row) + "\n")
print("fixtures=%d parts=%d refusals=%d mismatches=%d grew=%d shrank=%d"
      % (len(fixtures), parts, refusals, len(mismatches), grew, shrank))
print("output bytes: before=%d after=%d delta=%+.2f%%" % (bb, ab, 100.0 * (ab - bb) / bb))
