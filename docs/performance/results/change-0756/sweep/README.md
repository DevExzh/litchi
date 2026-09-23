# Sweep 0756: before/after across the whole wave (009d515bef to ddf788eb80)

This is evidence for integration record 0756. **It is descriptive only, not a registered claim.** Records 0742–0755 each carry their own rigorous measurements, and those remain the authority for each change. This sweep only describes how a fixed list of cases moved between the wave's start and its merged tip.

Results are in `summary.md` (table) and `summary.json` (the same data plus per-process values). This file records how they were produced and what they do not show.

## Arms and binaries

| Arm | Commit | Working directory (cwd of every process) | Binary (invoked by absolute path) | SHA-256 | Bytes |
|---|---|---|---|---|---:|
| A = base | `009d515befc1b56c07653258b5250139507a33f5` | `/home/zhuhe/code/litchi-worktrees/base-009d515bef` | `/home/zhuhe/code/litchi-worktrees/targets/base-009d515bef/release/litchi-perf-baseline` | `fb535ebb4abb4154c4ac445b90a60399edf189d52cae76fea2947852a406cb0b` | 62,962,456 |
| B = final | `ddf788eb80e5370de042dbd9a9525b3a8783d9f0` (records 0742–0755 merged) | `/home/zhuhe/code/litchi-worktrees/wave-ddf788eb80` (removed after the run) | `/home/zhuhe/code/litchi-worktrees/targets/wave-ddf788eb80/release/litchi-perf-baseline` (removed after the run) | `eabe691dc58b7d148d4d29ec06f78ef69c6efea7f40307b0422a25c7578d42ba` | 63,768,568 |

The two arms have checkout paths of equal length (49 bytes) and binary paths of equal length (86 bytes), so `argv[0]` has the same length in both. Each report records its binary's SHA-256, the git revision and a clean worktree, and all 56 reports agree with the table above.

Both binaries were built with the same command, run from their own checkout. Each checkout held a copy of the gitignored `Cargo.lock` that is byte-identical to `/home/zhuhe/code/litchi/Cargo.lock`:

```sh
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/targets/<base-009d515bef|wave-ddf788eb80> CARGO_BUILD_JOBS=12 \
  cargo build --release --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
```

The base binary already existed (built 2026-09-22 18:56). The final binary was built for this sweep (2026-09-23 08:23, 1 min 35 s) after these commands:

```sh
git -C /home/zhuhe/code/litchi worktree add --detach /home/zhuhe/code/litchi-worktrees/wave-ddf788eb80 ddf788eb80
cp /home/zhuhe/code/litchi/Cargo.lock /home/zhuhe/code/litchi-worktrees/wave-ddf788eb80/Cargo.lock
```

No `RUSTFLAGS` were set, and `.cargo/config.toml` holds only a lint alias. The reports record `rustflags: null` and the Rust system allocator.

## Environment

- **CPU:** AMD EPYC 9R45 (AWS VM), 32 cores, 1 thread per core, 1 socket, 128 MiB L3 in 4 instances. Memory is 123 GiB (MemTotal 129,447,068 kB).
- **Kernel:** `Linux 7.0.0-1012-aws #12-Ubuntu SMP PREEMPT Tue Aug 11 15:33:41 UTC 2026`. The kernel command line has no `isolcpus` or `nohz_full`, so core 20 is not isolated. The VM does not expose the cpufreq governor.
- **Toolchain:** `rustc -V` gives `rustc 1.95.0 (59807616e 2026-04-14)` in both checkouts, pinned by `rust-toolchain.toml`. Cargo is `cargo 1.95.0 (f2d3ce0bd 2026-03-21)`.
- **perf:** version 7.0.14, with `kernel.perf_event_paranoid = 1`.
- **Temporary files:** every command ran with `TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/sweep-0756/tmp`, and no command in this sweep used `/tmp`. The harness processes left nothing behind in `TMPDIR`.
- **Load:** other agents' work shared the host. The 1-minute load average sampled before each of the 56 processes ranged from 0.65 to 4.96 (see `raw/runlog.txt`). `mpstat` showed core 20 idle just before the run.
- **Run window:** 2026-09-23 08:28:54 to about 08:45:08 EDT.

## Case groups

| Group | Shape flags | Cases | Rows |
|---|---|---|---:|
| S | `--semantic-shape medium,large` | docx_semantic_open, docx_semantic_full_text, docx_semantic_noop_edit_save, docx_semantic_one_edit_save, docx_semantic_one_percent_edit_save, pptx_semantic_open, pptx_semantic_full_text, pptx_semantic_noop_edit_save, pptx_semantic_one_edit_save, pptx_semantic_one_percent_edit_save, doc_semantic_open, doc_semantic_full_text, doc_semantic_one_edit_save, xls_semantic_open, xls_semantic_full_cell_scan, xls_semantic_one_edit_save, ppt_semantic_open, ppt_semantic_full_text, ppt_semantic_one_edit_save, docx_streaming_create, pptx_streaming_create | 42 |
| X | `--xlsx-shape medium,dense-wide` | xlsx_open_owned, xlsx_first_cell, xlsx_full_cell_scan, xlsx_one_cell_commit_save, xlsx_one_percent_commit_save, xlsx_streaming_create | 13 |
| W | `--writer-shape tiny,large,payload-heavy` | doc_fresh_write_to, ppt_fresh_write_to, xls_fresh_write_to | 9 |
| M | none (harness defaults) | pptx_cross_copy_plain_lifecycle, pptx_cross_copy_media_rich_lifecycle, pptx_source_backed_cross_copy_media_rich_lifecycle, xlsx_source_backed_cell_values_one_edit_save, xlsx_eager_cell_values_one_edit_save, docx_source_backed_one_edit_save, pptx_source_backed_one_edit_save, docx_ordinary_save_lifecycle, xlsx_ordinary_save_lifecycle, pptx_ordinary_save_lifecycle, xls_visibility_eager_edit_save, xls_comments_eager_edit_save, xls_numeric_eager_rk_mulrk_edit_save, ole_common_one_edit_save | 23 |

Some rows come from corpora that the shape flags do not control:

- **Legacy DOC, XLS and PPT semantic cases (group S).** `--semantic-shape` does not apply to them. The harness builds their corpora with each arm's fresh writer over the default `--writer-shape` list minus `payload-heavy`, which gives `doc-tiny`, `doc-large` and so on. The descriptors, including the archive SHA-256, are nonetheless identical across the arms.
- **`xlsx_streaming_create` (group X).** It brings its own tiny, medium and large corpora.
- **Group M.** `ole_common_one_edit_save` runs over the 4 default corpus shapes times 2 payload kinds. The `xlsx_*_cell_values_*` cases run over the default cell-CRUD shapes `medium` and `dense-sparse`.

**No cases were dropped.** Every case ran on both binaries, and all 56 processes exited 0 with empty stdout and stderr. Within each group, all 8 reports contain the same (case, corpus) set.

## Protocol and exact commands

Each group ran 8 processes back to back in the order **A B B A A B B A** (positions 1–8, where A is base and B is final), giving 4 processes per arm. The groups ran in the order W, X, S, M. Every process had this form:

```sh
cd <arm working directory> && taskset -c 20 <absolute binary path> --samples 9 --warmup 2 <group shape flags> --case <group case list> --json /home/zhuhe/code/litchi-worktrees/scratch/sweep-0756/raw/<G>-seq<pos>-<arm>.json
```

`raw/runlog.txt` holds the literal command line of every process, with its cwd, return code, start time, wall time and load averages before and after. `run.sh` below generated those command lines.

**Instruction counts.** After a group's timed rounds, one extra process per arm (base first) ran with the same flags, wrapped as follows:

```sh
perf stat -e instructions,instructions:u -x, -o raw/<G>-perfstat-<arm>.csv taskset -c 20 <binary> --samples 9 --warmup 2 <flags> --case <list> --json raw/<G>-perfstat-<arm>.json
```

These processes are not used in the timing table. `instructions:u` was added next to the requested `instructions` so that kernel work can be separated out.

**Supplementary run M1.** Case `pptx_source_backed_cross_copy_media_rich_lifecycle` ran alone, with default flags and the same A B B A A B B A protocol (`raw/M1-*`). It was added because that row settled into different modes in different processes within group M.

**Probe.** Before the sweep, one 1-sample probe per arm per group checked that every case runs on both binaries and that the corpus descriptors match. The probe reports were not kept.

## Statistic

The same calculation is applied to every (case, corpus) row:

1. Each process reports the p50 of its 9 samples (`elapsed_ns.p50`).
2. **Base p50** and **final p50** are the medians of the four per-process p50s of each arm, that is, the mean of the middle two.
3. **Ratio** is final p50 divided by base p50. A value below 1 means final is faster.
4. **Paired range** is the minimum and maximum of final p50 divided by base p50 over the adjacent position pairs (1,2), (3,4), (5,6) and (7,8). Each of those pairs holds one process per arm.

The geometric means in `summary.md` are unweighted means over rows, and are descriptive only.

## Checks that hold for every row

- **Same inputs.** The corpus descriptors are identical in all 8 processes of every row: name, shape, entry counts, sizes, archive SHA-256, target entry and its SHA-256. Both arms therefore timed the same input bytes.
- **Same harness behaviour.** The two harnesses print the same `--help` and case list. Their sources differ only in `tools/perf-baseline/src/xls_numeric.rs`, where record 0748 moved an evidence-schema label to v2 and changed test assertions. No timed region changed.
- **Same outputs, with one exception.** 22 of the 87 rows report an output SHA-256. For 21 of them it is identical between the arms and stable within each arm. The exception is `pptx_cross_copy_media_rich_lifecycle`, where final writes 33,599,745 bytes and base writes 33,599,873, with a different SHA-256. This is by design: record 0742 (commit `daaa0ff721`) copies images' source-compressed bytes instead of recompressing them. That row's 0.29 ratio therefore compares different output bytes.

## Caveats

- **Descriptive only.** This sweep is not a registered claim. The per-record measurements in 0742–0755 are the authority for each change.
- **Small samples.** Each process takes 9 samples after 2 warmups, and each arm has 4 processes. The p50 of 9 samples is the 5th order statistic.
- **One core.** All processes ran under `taskset -c 20`, and the reports show `logical_cpus_available = 1`. Any parallel path, such as ADR 0031's parallel deflate, therefore ran at width 1. Nothing here describes multi-core behaviour.
- **Loaded host.** Other agents' builds and measurements shared the host, with load averages from 0.65 to 4.96. Core 20 is not isolated, and the host is a VM.
- **Code layout.** The paths were kept at equal lengths, but the binaries differ in size and layout. Shifts of a few percent, such as the 1.01–1.07 ratios on several open, scan and full-text rows, cannot be told apart from layout effects by this sweep.
- **Timer granularity.** Rows flagged `sub-10us per sample` are close to timer granularity, so their ratios carry no weight.
- **Shared heap.** The cases of a group run in one process and share one heap, so a case's time can depend on the cases before it. The row `pptx_source_backed_cross_copy_media_rich_lifecycle` shows this: each process settles into one page-fault mode (0, 4,577, 8,203 or 12,268 minor faults per sample) and keeps it for all 9 samples. Every mode seen in both arms gives the same time in both. The M-group ratio of 1.27, with a paired range of 0.99–1.63, therefore reflects which mode each process landed in. Run alone in M1, both arms show the same two modes at the same times. See the supplementary table in `summary.md`.
- **Instruction counts cover whole processes.** They include corpus generation, verification and other untimed work, and the corpus generators use library code that the wave changed. The counts therefore cannot be attributed to the timed regions alone.
- **Warm cache only.** Everything ran with a warm page cache, and no filesystem cases were selected.

## Files

- `summary.md`: the table sorted by format and then case, the geometric means, the supplementary M1 table and the instruction counts.
- `summary.json`: the same data, plus per-position p50s, paired ratios, spread within each arm, output SHA-256 sets and flags.
- `raw/<G>-seq<pos>-<arm>.json.gz`: the 32 timed reports of groups S, X, W and M, plus the 8 reports of M1.
- `raw/<G>-perfstat-<arm>.csv` and `raw/<G>-perfstat-<arm>.json.gz`: the `perf stat` output and harness report of each instruction-count process.
- `raw/runlog.txt`: for each of the 56 processes, the command line, cwd, return code, start time, wall time and load averages.

## Cleanup

After the run, these were removed:

- the final worktree, `/home/zhuhe/code/litchi-worktrees/wave-ddf788eb80` (9.9 GiB), with `git -C /home/zhuhe/code/litchi worktree remove --force`;
- its target directory, `/home/zhuhe/code/litchi-worktrees/targets/wave-ddf788eb80` (1,013 MiB);
- the probe reports, the final build log, the empty per-process stdout and stderr logs, and `tmp/`.

The base checkout and base target directory existed before this sweep and were left untouched.

## Scripts

To reproduce `summary.md` and `summary.json` from `raw/`, run `python3 summarize.py`. The script reads either `.json` or `.json.gz`.

<details><summary><code>run.sh</code>: the runner, invoked as <code>run.sh W 1 2 3 4 5 6 7 8</code>, <code>run.sh X 1 2 3 4 5 6 7 8</code>, <code>run.sh S 1 2 3 4</code>, <code>run.sh S 5 6 7 8</code>, <code>run.sh M 1</code>, <code>run.sh M 2 3 4 5 6 7 8</code>, then <code>run.sh &lt;G&gt; perf</code> for W, X, S and M, then <code>run.sh M1 1 2 3 4 5 6 7 8</code></summary>

```bash
#!/usr/bin/env bash
# Usage: run.sh GROUP POS... | run.sh GROUP perf
# POS is a 1-based position in the sequence A B B A A B B A (A = base, B = final).
set -u
export TMPDIR=/home/zhuhe/code/litchi-worktrees/scratch/sweep-0756/tmp
OUT=/home/zhuhe/code/litchi-worktrees/scratch/sweep-0756
G=$1; shift
SEQ=(none base final final base base final final base)

S_CASES=docx_semantic_open,docx_semantic_full_text,docx_semantic_noop_edit_save,docx_semantic_one_edit_save,docx_semantic_one_percent_edit_save,pptx_semantic_open,pptx_semantic_full_text,pptx_semantic_noop_edit_save,pptx_semantic_one_edit_save,pptx_semantic_one_percent_edit_save,doc_semantic_open,doc_semantic_full_text,doc_semantic_one_edit_save,xls_semantic_open,xls_semantic_full_cell_scan,xls_semantic_one_edit_save,ppt_semantic_open,ppt_semantic_full_text,ppt_semantic_one_edit_save,docx_streaming_create,pptx_streaming_create
X_CASES=xlsx_open_owned,xlsx_first_cell,xlsx_full_cell_scan,xlsx_one_cell_commit_save,xlsx_one_percent_commit_save,xlsx_streaming_create
W_CASES=doc_fresh_write_to,ppt_fresh_write_to,xls_fresh_write_to
M_CASES=pptx_cross_copy_plain_lifecycle,pptx_cross_copy_media_rich_lifecycle,pptx_source_backed_cross_copy_media_rich_lifecycle,xlsx_source_backed_cell_values_one_edit_save,xlsx_eager_cell_values_one_edit_save,docx_source_backed_one_edit_save,pptx_source_backed_one_edit_save,docx_ordinary_save_lifecycle,xlsx_ordinary_save_lifecycle,pptx_ordinary_save_lifecycle,xls_visibility_eager_edit_save,xls_comments_eager_edit_save,xls_numeric_eager_rk_mulrk_edit_save,ole_common_one_edit_save

case $G in
  S) FLAGS=(--semantic-shape medium,large); CASES=$S_CASES ;;
  X) FLAGS=(--xlsx-shape medium,dense-wide); CASES=$X_CASES ;;
  W) FLAGS=(--writer-shape tiny,large,payload-heavy); CASES=$W_CASES ;;
  M) FLAGS=(); CASES=$M_CASES ;;
  M1) FLAGS=(); CASES=pptx_source_backed_cross_copy_media_rich_lifecycle ;;
  *) echo "unknown group $G"; exit 2 ;;
esac

arm_paths() {
  if [ "$1" = base ]; then
    WD=/home/zhuhe/code/litchi-worktrees/base-009d515bef
    BIN=/home/zhuhe/code/litchi-worktrees/targets/base-009d515bef/release/litchi-perf-baseline
  else
    WD=/home/zhuhe/code/litchi-worktrees/wave-ddf788eb80
    BIN=/home/zhuhe/code/litchi-worktrees/targets/wave-ddf788eb80/release/litchi-perf-baseline
  fi
}

run_one() { # tag arm prefix...
  local tag=$1 arm=$2; shift 2
  arm_paths "$arm"
  local json=$OUT/raw/$G-$tag-$arm.json
  local cmd=("$@" taskset -c 20 "$BIN" --samples 9 --warmup 2 "${FLAGS[@]}" --case "$CASES" --json "$json")
  local l0 l1 t0 t1 rc
  l0=$(cut -d' ' -f1-3 /proc/loadavg); t0=$(date +%s.%N)
  (cd "$WD" && "${cmd[@]}") > "$OUT/raw/$G-$tag-$arm.log" 2>&1
  rc=$?
  t1=$(date +%s.%N); l1=$(cut -d' ' -f1-3 /proc/loadavg)
  printf '%s %s %s rc=%s start=%s wall_s=%.1f load_before=[%s] load_after=[%s] cwd=%s cmd=%s\n' \
    "$G" "$tag" "$arm" "$rc" "$(date -d @"${t0%.*}" +%H:%M:%S)" "$(echo "$t1 - $t0" | bc)" "$l0" "$l1" "$WD" "${cmd[*]}" >> "$OUT/raw/runlog.txt"
  echo "$G $tag $arm rc=$rc wall=$(echo "$t1 - $t0" | bc)"
}

if [ "${1:-}" = perf ]; then
  for arm in base final; do
    run_one perfstat "$arm" perf stat -e instructions,instructions:u -x, -o "$OUT/raw/$G-perfstat-$arm.csv"
  done
else
  for pos in "$@"; do
    run_one "seq$pos" "${SEQ[$pos]}"
  done
fi
```

</details>

<details><summary><code>summarize.py</code></summary>

```python
#!/usr/bin/env python3
"""Summarize the sweep-0756 raw reports into summary.json and summary.md."""
import json, statistics, os, glob, re, gzip

OUT = "/home/zhuhe/code/litchi-worktrees/scratch/sweep-0756"
RAW = os.path.join(OUT, "raw")
SEQ = ["base", "final", "final", "base", "base", "final", "final", "base"]  # positions 1..8
PAIRS = [(1, 2), (3, 4), (5, 6), (7, 8)]
GROUPS = {
    "S": "--semantic-shape medium,large",
    "X": "--xlsx-shape medium,dense-wide",
    "W": "--writer-shape tiny,large,payload-heavy",
    "M": "(default shape flags)",
}
FORMAT_ORDER = ["DOCX", "PPTX", "XLSX", "DOC", "PPT", "XLS", "OLE2 common"]


def fmt_of(case):
    for pre, f in (("docx_", "DOCX"), ("pptx_", "PPTX"), ("xlsx_", "XLSX"), ("doc_", "DOC"),
                   ("ppt_", "PPT"), ("xls_", "XLS"), ("ole_common_", "OLE2 common")):
        if case.startswith(pre):
            return f
    raise ValueError(case)


def load(path):
    if os.path.exists(path):
        with open(path) as fh:
            return json.load(fh)
    with gzip.open(path + ".gz", "rt") as fh:
        return json.load(fh)


ANNOTATIONS = {
    ("pptx_cross_copy_media_rich_lifecycle", "pptx_cross_copy_media_rich_lifecycle"):
        "final writes different bytes (33,599,745 vs 33,599,873; record 0742 transfers source-compressed image bytes)",
    ("pptx_source_backed_cross_copy_media_rich_lifecycle", "pptx_source_backed_cross_copy_media_rich_lifecycle"):
        "per-process page-fault modes; equal per mode and no arm difference when run alone (supplementary M1 below)",
}


def g4(x):
    return float(f"{x:.4g}")


rows = []
checks = []
instr = {}
for g in GROUPS:
    reports = {}
    for pos in range(1, 9):
        arm = SEQ[pos - 1]
        reports[pos] = load(os.path.join(RAW, f"{g}-seq{pos}-{arm}.json"))
    # identity and configuration checks
    for pos, rep in reports.items():
        arm = SEQ[pos - 1]
        cfg = rep["configuration"]
        assert cfg["samples_per_case"] == 9 and cfg["warmup_iterations_per_case"] == 2, (g, pos)
        exp_rev = {"base": "009d515bef", "final": "ddf788eb80"}[arm]
        assert rep["environment"]["git_revision"].startswith(exp_rev), (g, pos)
        assert rep["environment"]["git_worktree_dirty"] is False, (g, pos)
    base_sha = {reports[p]["binary_identity"]["binary_sha256"] for p in range(1, 9) if SEQ[p - 1] == "base"}
    final_sha = {reports[p]["binary_identity"]["binary_sha256"] for p in range(1, 9) if SEQ[p - 1] == "final"}
    assert len(base_sha) == 1 and len(final_sha) == 1
    keys = [(r["case"], r["corpus"]["name"]) for r in reports[1]["results"]]
    for pos in range(2, 9):
        k2 = [(r["case"], r["corpus"]["name"]) for r in reports[pos]["results"]]
        assert set(k2) == set(keys), (g, pos)
    idx = {pos: {(r["case"], r["corpus"]["name"]): r for r in reports[pos]["results"]} for pos in reports}
    for key in keys:
        case, corpus = key
        per = {pos: idx[pos][key] for pos in range(1, 9)}
        corp = {json.dumps(per[p]["corpus"], sort_keys=True) for p in per}
        outs = {arm: {per[p].get("output_sha256") for p in per if SEQ[p - 1] == arm} for arm in ("base", "final")}
        p50 = {pos: per[pos]["elapsed_ns"]["p50"] / 1e6 for pos in per}
        base = [p50[p] for p in range(1, 9) if SEQ[p - 1] == "base"]
        final = [p50[p] for p in range(1, 9) if SEQ[p - 1] == "final"]
        mb, mf = statistics.median(base), statistics.median(final)
        paired = []
        for a, b in PAIRS:
            fb = p50[a] if SEQ[a - 1] == "final" else p50[b]
            bb = p50[a] if SEQ[a - 1] == "base" else p50[b]
            paired.append(fb / bb)
        flags = []
        if key in ANNOTATIONS:
            flags.append(ANNOTATIONS[key])
        if min(paired) < 1.0 < max(paired):
            flags.append("paired range straddles 1.0")
        if max(mb, mf) < 0.01:
            flags.append("sub-10us per sample")
        if len(corp) != 1:
            flags.append("CORPUS IDENTITY DIFFERS")
        if outs["base"] != outs["final"] and (None not in outs["base"]) and key not in ANNOTATIONS:
            flags.append("output bytes differ between arms")
        rows.append({
            "group": g,
            "format": fmt_of(case),
            "case": case,
            "corpus": corpus,
            "order": len(rows),
            "base_p50_ms_by_position": {str(p): p50[p] for p in range(1, 9) if SEQ[p - 1] == "base"},
            "final_p50_ms_by_position": {str(p): p50[p] for p in range(1, 9) if SEQ[p - 1] == "final"},
            "base_median_p50_ms": mb,
            "final_median_p50_ms": mf,
            "ratio_final_over_base": mf / mb,
            "paired_ratios": paired,
            "paired_ratio_min": min(paired),
            "paired_ratio_max": max(paired),
            "base_p50_spread_max_over_min": max(base) / min(base),
            "final_p50_spread_max_over_min": max(final) / min(final),
            "corpus_identical_across_all_8_processes": len(corp) == 1,
            "output_sha256": {k: sorted(x for x in v if x) for k, v in outs.items()},
            "flags": flags,
        })
    # perf stat
    instr[g] = {}
    for arm in ("base", "final"):
        vals = {}
        with open(os.path.join(RAW, f"{g}-perfstat-{arm}.csv")) as fh:
            for line in fh:
                if not line.strip() or line.startswith("#"):
                    continue
                parts = line.strip().split(",")
                vals[parts[2]] = int(parts[0])
        instr[g][arm] = vals

# supplementary isolated run of one case (group M1), same protocol
supp = []
for pos in range(1, 9):
    arm = SEQ[pos - 1]
    rep = load(os.path.join(RAW, f"M1-seq{pos}-{arm}.json"))
    assert rep["configuration"]["samples_per_case"] == 9
    (r,) = rep["results"]
    src = r["source"]["pptx_source_backed_cross_copy_lifecycle"]
    supp.append({
        "position": pos,
        "arm": arm,
        "p50_ms": r["elapsed_ns"]["p50"] / 1e6,
        "publication_p50_ms": statistics.median(src["publication_ns"]) / 1e6,
        "plan_p50_ms": statistics.median(src["plan_ns"]) / 1e6,
        "minor_faults_per_sample_median": statistics.median(r["operation_metrics"]["process"]["minor_faults"]["values"]),
    })
m_rows = {}
for pos in range(1, 9):
    arm = SEQ[pos - 1]
    rep = load(os.path.join(RAW, f"M-seq{pos}-{arm}.json"))
    for r in rep["results"]:
        if r["case"] == "pptx_source_backed_cross_copy_media_rich_lifecycle":
            src = r["source"]["pptx_source_backed_cross_copy_lifecycle"]
            m_rows[pos] = {
                "position": pos, "arm": arm, "p50_ms": r["elapsed_ns"]["p50"] / 1e6,
                "publication_p50_ms": statistics.median(src["publication_ns"]) / 1e6,
                "plan_p50_ms": statistics.median(src["plan_ns"]) / 1e6,
                "minor_faults_per_sample_median": statistics.median(r["operation_metrics"]["process"]["minor_faults"]["values"]),
            }

rows.sort(key=lambda r: (FORMAT_ORDER.index(r["format"]), r["case"], r["order"]))

# geometric means per group and per format (descriptive only)
def geomean(xs):
    import math
    return math.exp(sum(math.log(x) for x in xs) / len(xs))

by_group = {g: geomean([r["ratio_final_over_base"] for r in rows if r["group"] == g]) for g in GROUPS}
by_format = {f: geomean([r["ratio_final_over_base"] for r in rows if r["format"] == f])
             for f in FORMAT_ORDER if any(r["format"] == f for r in rows)}
counts = {f: sum(1 for r in rows if r["format"] == f) for f in by_format}

summary = {
    "schema": "sweep-0756-descriptive-v1",
    "claim": "descriptive only; not a registered claim",
    "base_commit": "009d515befc1b56c07653258b5250139507a33f5",
    "final_commit": "ddf788eb80e5370de042dbd9a9525b3a8783d9f0",
    "base_binary_sha256": base_sha.pop(),
    "final_binary_sha256": final_sha.pop(),
    "protocol": {
        "sequence": "A B B A A B B A (A = base, B = final), per group",
        "processes_per_arm_per_group": 4,
        "samples_per_process": 9,
        "warmup_per_process": 2,
        "cpu_affinity": "taskset -c 20",
        "statistic": "per (case, corpus): median of the four per-process p50s per arm; ratio = final median / base median; paired ratios = final p50 / base p50 within positions (1,2), (3,4), (5,6), (7,8)",
    },
    "groups": GROUPS,
    "rows": rows,
    "geomean_ratio_by_group": by_group,
    "geomean_ratio_by_format": by_format,
    "rows_by_format": counts,
    "instructions_whole_process": instr,
    "dropped_cases": [],
    "supplementary_isolated_pptx_source_backed_cross_copy_media_rich_lifecycle": {
        "command_group": "M1: --case pptx_source_backed_cross_copy_media_rich_lifecycle alone, default flags, same protocol",
        "processes": supp,
        "in_group_M_processes": [m_rows[p] for p in range(1, 9)],
        "reading": "Each process settles into one page-fault mode for this case and keeps it for all 9 samples. Every mode seen in both arms gives the same time in both. The 12,268-fault mode (19.4-19.9 ms) appeared only in two final processes in group M, where the case runs directly after pptx_cross_copy_media_rich_lifecycle, whose allocation pattern record 0742 changed; run alone (M1), both arms show the same two modes at the same times.",
    },
}
with open(os.path.join(OUT, "summary.json"), "w") as fh:
    json.dump(summary, fh, indent=1)


def ms(x):
    return f"{x:.4g}"


lines = []
lines.append("# Sweep 0756: wave-wide before/after (009d515bef -> ddf788eb80)")
lines.append("")
lines.append("Descriptive only, not a registered claim. Four processes per arm per group in the order "
             "A B B A A B B A (A = base 009d515bef, B = final ddf788eb80), each `taskset -c 20 ... --samples 9 --warmup 2`. "
             "Base/final p50 = median of the four per-process p50s; ratio = final / base (below 1 is faster); "
             "paired range = min-max of final/base over the four adjacent (A,B) process pairs. See README.md for caveats.")
lines.append("")
lines.append("| Format | Case | Corpus | Grp | Base p50 (ms) | Final p50 (ms) | Ratio | Paired range | Notes |")
lines.append("|---|---|---|---|---:|---:|---:|---|---|")
for r in rows:
    notes = "; ".join(r["flags"])
    lines.append(f"| {r['format']} | `{r['case']}` | `{r['corpus']}` | {r['group']} | {ms(r['base_median_p50_ms'])} | "
                 f"{ms(r['final_median_p50_ms'])} | {r['ratio_final_over_base']:.3f} | "
                 f"{r['paired_ratio_min']:.3f}-{r['paired_ratio_max']:.3f} | {notes} |")
lines.append("")
lines.append("## Unweighted geometric mean of row ratios (descriptive)")
lines.append("")
lines.append("| Format | Rows | Geomean ratio |")
lines.append("|---|---:|---:|")
for f, v in by_format.items():
    lines.append(f"| {f} | {counts[f]} | {v:.3f} |")
lines.append("")
lines.append("| Group | Flags | Rows | Geomean ratio |")
lines.append("|---|---|---:|---:|")
for g, v in by_group.items():
    lines.append(f"| {g} | `{GROUPS[g]}` | {sum(1 for r in rows if r['group'] == g)} | {v:.3f} |")
lines.append("")
lines.append("## Supplementary: `pptx_source_backed_cross_copy_media_rich_lifecycle` by process")
lines.append("")
lines.append("Each process settles into one page-fault mode for this case and keeps it for all nine samples. "
             "Every mode seen in both arms gives the same time in both. The 12,268-fault mode (19.4-19.9 ms) appeared "
             "only in two final processes in group M, where this case runs directly after "
             "`pptx_cross_copy_media_rich_lifecycle`, whose allocation pattern record 0742 changed. M1 re-runs the case "
             "alone (default flags, same A B B A A B B A protocol): both arms show the same two modes at the same times, "
             "so the M-group ratio of 1.27 reflects which mode each process landed in rather than a slower code path.")
lines.append("")
lines.append("| Run | Pos | Arm | p50 (ms) | Publication p50 (ms) | Plan p50 (ms) | Minor faults / sample |")
lines.append("|---|---:|---|---:|---:|---:|---:|")
for label, series in (("M (in group)", [m_rows[p] for p in range(1, 9)]), ("M1 (alone)", supp)):
    for e in series:
        lines.append(f"| {label} | {e['position']} | {e['arm']} | {e['p50_ms']:.2f} | {e['publication_p50_ms']:.2f} | "
                     f"{e['plan_p50_ms']:.2f} | {e['minor_faults_per_sample_median']:,.0f} |")
lines.append("")
lines.append("## Whole-process instruction counts (`perf stat -e instructions,instructions:u -x,`)")
lines.append("")
lines.append("One extra process per arm per group, run after the timed rounds with the same flags (its JSON is kept "
             "as `raw/<G>-perfstat-<arm>.json.gz` and is not used in the timing table). Counts cover the whole "
             "process, including corpus generation and verification outside the timed regions.")
lines.append("")
lines.append("| Group | Base instructions | Final instructions | Ratio | Base instructions:u | Final instructions:u | Ratio (:u) |")
lines.append("|---|---:|---:|---:|---:|---:|---:|")
for g in GROUPS:
    b, f = instr[g]["base"], instr[g]["final"]
    lines.append(f"| {g} | {b['instructions']:,} | {f['instructions']:,} | {f['instructions']/b['instructions']:.3f} | "
                 f"{b['instructions:u']:,} | {f['instructions:u']:,} | {f['instructions:u']/b['instructions:u']:.3f} |")
lines.append("")
with open(os.path.join(OUT, "summary.md"), "w") as fh:
    fh.write("\n".join(lines))
print("\n".join(lines))
```

</details>
