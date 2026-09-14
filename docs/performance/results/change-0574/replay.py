#!/usr/bin/env python3
"""Recompute every number change 0574 cites, from this directory alone.

    python3 -B docs/performance/results/change-0574/replay.py

Sources, all retained here:

  counters/<mode>-<op>.json     xls_source_attribution, 20 warmups + 100 samples
  perf/<mode>-<op>-s{100,1100}.csv   perf stat isolation pairs
  callgrind/ann-<stem>-s{small,large}.txt   callgrind_annotate isolation pairs
  summary.json                  the folded document, including the 117 per-fixture
                                corpus rows whose raw captures are not retained
  model/globals-closure.json    the static consumed-closure model

Every table printed here must match the change record. Deterministic.
"""
import glob, json, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
SUMMARY = json.load(open(os.path.join(HERE, "summary.json")))

SST_MARKS = ("from_utf16", "SstCursor", "shared_string", "SharedString")
FMT_MARKS = ("xf_ext", "border_fill", "number_format", "differential_format", "style_ext")
CFB_VALIDATE = ("collect_exact", "collect_sector_chain", "validate_stream_allocations",
                "validate_physical_sector_layout", "claim_chain")
CALLGRIND = [("ConditionalFormattingSamples.xls", "flagship", 20, 220, 200),
             ("WithCustomViews.xls", "cv", 20, 120, 100),
             ("54016.xls", "54016", 10, 60, 50)]


def p(vals, q):
    v = sorted(vals)
    k = (len(v) - 1) * q / 100.0
    f = int(k); c = min(f + 1, len(v) - 1)
    return v[f] + (v[c] - v[f]) * (k - f)


def table_counters():
    print("== deterministic counters, ConditionalFormattingSamples.xls ==")
    print("%-24s %6s %9s %5s %9s %9s %9s %9s" % (
        "mode/operation", "reads", "bytes", "ver", "p50 ns", "read ns", "ver ns", "rest ns"))
    for path in sorted(glob.glob(os.path.join(HERE, "counters", "*.json"))):
        d = json.load(open(path))
        recs = d["records"]; m = recs[0]["metrics"]
        for r in recs:
            assert r["metrics"]["read_calls"] == m["read_calls"], path
            assert r["metrics"]["read_bytes"] == m["read_bytes"], path
            assert r["metrics"]["version_calls"] == m["version_calls"], path
        el = p([r["elapsed_ns"] for r in recs], 50)
        rd = p([r["metrics"]["read_ns"] for r in recs], 50)
        vs = p([r["metrics"]["version_ns"] for r in recs], 50)
        print("%-24s %6d %9d %5d %9.0f %9.0f %9.0f %9.0f" % (
            f"{d['mode']}/{d['operation']}", m["read_calls"], m["read_bytes"],
            m["version_calls"], el, rd, vs, el - rd - vs))


def table_perf():
    def load(path):
        d = {}
        for line in open(path):
            if line.startswith("#") or not line.strip():
                continue
            f = line.split(",")
            try: d[f[2]] = float(f[0])
            except (IndexError, ValueError): pass
        return d
    print("\n== perf stat, isolated by differencing 1100- and 100-sample children ==")
    print("%-24s %12s %12s %10s %12s" % ("mode/operation", "instr/op", "branches/op",
                                         "bmiss/op", "task-clock us"))
    for mode in ("owned-readat", "file-source"):
        for op in ("open", "one-cell"):
            a = load(os.path.join(HERE, "perf", f"{mode}-{op}-s100.csv"))
            b = load(os.path.join(HERE, "perf", f"{mode}-{op}-s1100.csv"))
            print("%-24s %12.0f %12.0f %10.1f %12.2f" % (
                f"{mode}/{op}", (b["instructions"] - a["instructions"]) / 1000.0,
                (b["branches"] - a["branches"]) / 1000.0,
                (b["branch-misses"] - a["branch-misses"]) / 1000.0,
                b["task-clock"] - a["task-clock"]))


def ann(path):
    out = {}
    for line in open(path):
        m = re.match(r"^\s*([\d,]+)(?:\s*\([^)]*\))?\s+(\S.*?)\s*$", line)
        if not m: continue
        try: n = int(m.group(1).replace(",", ""))
        except ValueError: continue
        name = m.group(2)
        if name.startswith("Ir") or "PROGRAM TOTALS" in name: continue
        out[name] = out.get(name, 0) + n
    return out


def table_callgrind():
    print("\n== callgrind instruction attribution per source-backed open ==")
    cols = [c[0] for c in CALLGRIND]
    data = {}
    for label, stem, small, large, opens in CALLGRIND:
        a = ann(os.path.join(HERE, "callgrind", f"ann-{stem}-s{large}.txt"))
        b = ann(os.path.join(HERE, "callgrind", f"ann-{stem}-s{small}.txt"))
        rows = sorted(((v - b.get(k, 0)) / float(opens), k)
                      for k, v in a.items() if v - b.get(k, 0) > 0)
        rows.sort(reverse=True)
        tot = sum(r[0] for r in rows)
        def share(marks, rows=rows, tot=tot):
            return sum(v for v, k in rows if any(s in k for s in marks)) / tot
        data[label] = {
            "Ir per open": tot,
            "shared-string scan": share(SST_MARKS),
            "memset + memcpy": share(("memset", "memcpy")),
            "allocator": share(("/malloc/", "__rdl_alloc", "__rdl_dealloc", "finish_grow")),
            "formatting parse": share(FMT_MARKS),
            "CFB whole-container validation": share(CFB_VALIDATE),
            "FAT chain prefix walk": share(("next_chain_sector",)),
            "top": rows[:6],
        }
    print("%-32s %14s %14s %14s" % ("category", *[c[:14] for c in cols]))
    for key in ("shared-string scan", "memset + memcpy", "allocator", "formatting parse",
                "CFB whole-container validation", "FAT chain prefix walk"):
        print("%-32s %13.2f%% %13.2f%% %13.2f%%" % (key, *[100 * data[c][key] for c in cols]))
    print("%-32s %14.0f %14.0f %14.0f" % ("Ir per open", *[data[c]["Ir per open"] for c in cols]))
    for c in cols:
        print(f"\n  top symbols, {c}:")
        for v, k in data[c]["top"]:
            print("    %10.0f %6.2f%%  %s" % (v, 100 * v / data[c]["Ir per open"],
                                              k.split(" [")[0][:78]))


def solve(A, y):
    n = len(A[0])
    M = [[sum(A[k][i] * A[k][j] for k in range(len(A))) for j in range(n)]
         + [sum(A[k][i] * y[k] for k in range(len(A)))] for i in range(n)]
    for i in range(n):
        q = max(range(i, n), key=lambda r: abs(M[r][i]))
        M[i], M[q] = M[q], M[i]
        for r in range(n):
            if r != i and M[i][i]:
                f = M[r][i] / M[i][i]
                for c in range(i, n + 1):
                    M[r][c] -= f * M[i][c]
    return [M[i][n] / M[i][i] for i in range(n)]


def table_corpus():
    rows = SUMMARY["corpus"]["rows"]
    A = [[1.0, r["skipped_payload_bytes"], r["sst_bytes"],
          r["closure_bytes"] - r["sst_bytes"]] for r in rows]
    y = [r["p50_ns"] for r in rows]
    a, b, c, d = solve(A, y)
    pred = [a + b * r["skipped_payload_bytes"] + c * r["sst_bytes"]
            + d * (r["closure_bytes"] - r["sst_bytes"]) for r in rows]
    ss = sum((q - w) ** 2 for q, w in zip(y, pred))
    ym = sum(y) / len(y)
    sm = sum((q - ym) ** 2 for q in y)
    T = sum(y)
    print("\n== corpus regression over %d XLS fixtures that open ==" % len(rows))
    print("p50_open_ns = %.0f + %.4f*skipped + %.4f*sst + %.4f*other_closure   R^2 = %.4f"
          % (a, b, c, d, 1 - ss / sm))
    print("  opaque (skipped) globals bytes : %8.1f ns/KiB" % (b * 1024))
    print("  SST bytes                      : %8.1f ns/KiB" % (c * 1024))
    print("  other consumed closure bytes   : %8.1f ns/KiB" % (d * 1024))
    S = sum(r["sst_bytes"] for r in rows)
    K = sum(r["skipped_payload_bytes"] for r in rows)
    O = sum(r["closure_bytes"] - r["sst_bytes"] for r in rows)
    print("  aggregate p50 open time        : %10.0f ns" % T)
    for name, val in (("SST bytes", c * S), ("opaque over-read", b * K),
                      ("other consumed closure", d * O), ("fixed per open", a * len(rows))):
        print("    %-28s %10.0f ns  %5.1f%%" % (name, val, 100 * val / T))


def table_closure():
    m = json.load(open(os.path.join(HERE, "model", "globals-closure.json")))["summary"]
    print("\n== static consumed-closure model over %d XLS fixtures ==" % m["files"])
    print("  globals bytes read today       : %10d" % m["total_globals_bytes_read_today"])
    print("  four-byte record headers       : %10d" % m["total_header_bytes"])
    print("  consumed record payloads       : %10d" % m["total_consumed_payload_bytes"])
    print("  touched closure                : %10d  (%.2f%%)" % (
        m["total_closure_bytes"], 100 * m["closure_share"]))
    print("  payloads never interpreted     : %10d  (%.2f%%)" % (
        m["total_skipped_payload_bytes"], 100 * (1 - m["closure_share"])))
    print("  modelled schedule, 512 B gap   : %10d reads, %d bytes" % (
        m["total_reads_gap512"], m["total_bytes_gap512"]))


def table_ministream():
    d = json.load(open(os.path.join(HERE, "ministream-probe.json")))
    print("\n== mini-stream materialization probe ==")
    print("%-22s %8s %8s %10s %12s %10s" % ("fixture", "reads", "payload", "globals",
                                            "ministream", "payload<ms"))
    for r in d["rows"]:
        print("%-22s %8d %8d %10d %12d %10s" % (
            r["fixture"], r["read_calls"], r["payload_bytes"], r["globals_end"],
            r["ministream_size"], r["payload_lt_ministream"]))
    print("  ->", d["answer"])


def table_inclusive():
    """Inclusive cost of the whole shared-string scan, and of the CFB constructor."""
    print("\n== inclusive Ir per open, same isolation pairs ==")
    print("%-34s %12s %12s %10s" % ("fixture", "scan_sst", "share", "CFB ctor"))
    for label, stem, small, large, opens in CALLGRIND:
        a = ann(os.path.join(HERE, "callgrind", f"incl-{stem}-s{large}.txt"))
        b = ann(os.path.join(HERE, "callgrind", f"incl-{stem}-s{small}.txt"))
        self_a = ann(os.path.join(HERE, "callgrind", f"ann-{stem}-s{large}.txt"))
        self_b = ann(os.path.join(HERE, "callgrind", f"ann-{stem}-s{small}.txt"))
        total = sum((v - self_b.get(k, 0)) / float(opens)
                    for k, v in self_a.items() if v - self_b.get(k, 0) > 0)
        def incl(mark):
            return sum(a[k] - b.get(k, 0) for k in a if mark in k) / float(opens)
        sst = incl("scan_shared_string_records")
        cfb = incl("OleFile<R>::open_with_limits")
        print("%-34s %12.0f %11.2f%% %10s" % (
            label, sst, 100 * sst / total,
            ("%.0f" % cfb) if cfb else "below threshold"))


def table_distribution():
    """How concentrated each opportunity is across the corpus."""
    def q(v, r):
        v = sorted(v)
        k = (len(v) - 1) * r / 100.0
        f = int(k); c = min(f + 1, len(v) - 1)
        return v[f] + (v[c] - v[f]) * (k - f)
    m = json.load(open(os.path.join(HERE, "model", "globals-closure.json")))
    ok = [r for r in m["files"] if "globals_end" in r]
    sh = [r["closure_bytes"] / r["globals_end"] for r in ok]
    print("\n== concentration ==")
    print("  closure share of globals: p10 %.1f%%  p50 %.1f%%  p90 %.1f%%"
          % (100 * q(sh, 10), 100 * q(sh, 50), 100 * q(sh, 90)))
    print("  fixtures consuming more than half their globals: %d of %d"
          % (sum(1 for x in sh if x > 0.5), len(ok)))
    big = sorted(ok, key=lambda r: -r["skipped_payload_bytes"])
    tot = sum(r["skipped_payload_bytes"] for r in ok)
    run = 0
    for i, r in enumerate(big):
        run += r["skipped_payload_bytes"]
        if run >= 0.9 * tot:
            print("  90%% of all skippable bytes live in %d of %d fixtures" % (i + 1, len(ok)))
            break
    rows = SUMMARY["corpus"]["rows"]
    ss = [r["sst_bytes"] / max(r["globals_end"], 1) for r in rows]
    print("  SST share of globals, opened fixtures: p50 %.1f%%  p90 %.1f%%  max %.1f%%"
          % (100 * q(ss, 50), 100 * q(ss, 90), 100 * max(ss)))
    print("  fixtures whose SST exceeds a quarter of their globals: %d of %d"
          % (sum(1 for x in ss if x > 0.25), len(rows)))
    comp = json.load(open(os.path.join(HERE, "model", "globals-composition.json")))
    for r in comp["files"]:
        if r.get("path", "").endswith("ole/xls/ConditionalFormattingSamples.xls"):
            print("  flagship continuation bytes by continued record: %s"
                  % r["continuation_bytes_by_continued_record"])
            mso = dict(r["top_record_bytes"]).get("MsoDrawingGroup", 0)
            cont = dict(r["continuation_bytes_by_continued_record"]).get("MsoDrawingGroup", 0)
            print("  flagship MsoDrawingGroup plus its continuation: %d of %d = %.1f%%"
                  % (mso + cont, r["globals_end"], 100 * (mso + cont) / r["globals_end"]))


def table_callers():
    """Per-open call counts and caller attribution, flagship fixture."""
    import re as _re
    def edge(path, callee, caller):
        seen, block = None, None
        for line in open(path):
            m = _re.match(r"^\s*[\d,]+\s*\([^)]*\)\s+\*\s+(.*)$", line)
            if m:
                block = m.group(1)
                continue
            c = _re.match(r"^\s*([\d,]+)\s*\([^)]*\)\s+<\s+(.*?)\s+\((\d+(?:,\d+)*)x\)", line)
            if c and caller in c.group(2):
                pend = (int(c.group(1).replace(",", "")), int(c.group(3).replace(",", "")))
                nxt = None
                continue
        # second pass: the callee block follows its caller lines
        lines = open(path).read().splitlines()
        for i, line in enumerate(lines):
            m = _re.match(r"^\s*([\d,]+)\s*\([^)]*\)\s+\*\s+(.*)$", line)
            if not m or callee not in m.group(2):
                continue
            for j in range(i - 1, max(i - 30, -1), -1):
                c = _re.match(r"^\s*([\d,]+)\s*\([^)]*\)\s+<\s+(.*?)\s+\((\d+(?:,\d+)*)x\)", lines[j])
                if c and caller in c.group(2):
                    return int(c.group(1).replace(",", "")), int(c.group(3).replace(",", ""))
                if _re.match(r"^\s*[\d,]+\s*\([^)]*\)\s+\*\s+", lines[j]):
                    break
        return 0, 0
    print("\n== per-open call counts, ConditionalFormattingSamples.xls ==")
    for callee, caller, label in (
        ("next_chain_sector", "read_stream_range", "FAT chain steps from read_stream_range"),
        ("__memset_avx2_unaligned_erms", "GlobalsBuffer::ensure", "memset calls from GlobalsBuffer::ensure"),
        ("String>::from_utf16", "SstCursor::read_characters", "from_utf16 calls (= shared strings decoded)"),
    ):
        ir_a, n_a = edge(os.path.join(HERE, "callgrind", "tree-flagship-s220.txt"), callee, caller)
        ir_b, n_b = edge(os.path.join(HERE, "callgrind", "tree-flagship-s20.txt"), callee, caller)
        print("  %-46s %10.1f calls, %10.0f Ir" % (label, (n_a - n_b) / 200.0, (ir_a - ir_b) / 200.0))


if __name__ == "__main__":
    table_counters(); table_perf(); table_callgrind(); table_inclusive(); table_callers()
    table_corpus(); table_closure(); table_distribution(); table_ministream()
