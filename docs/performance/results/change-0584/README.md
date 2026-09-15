# Evidence: change 0584, a fresh OLE2 profile at `e927e139c`

Change record: [`0584-ole2-profile-at-head.md`](../../0584-ole2-profile-at-head.md).

Disposition: retained. `performance_claim: none`. This packet is an attribution
profile. It carries per-symbol instruction counts, caller attribution, call
counts, chain-link counts and two independent byte-layout models — no timing, no
before/after comparison and no optimization claim.

## Contents

| Path | What it is |
| --- | --- |
| `PROFILE.txt` | The consolidated XLS profile every table in the record is read from: per-operation rankings, caller attribution, call counts, chain links and the noise measurement. Its header carries the binary hash and the three controls against change 0579's retained evidence. |
| `ranking.txt` | The same rankings without the commentary. |
| `folded.json` | The folded per-symbol XLS data, machine-readable. |
| `folded-docppt.json` | The same for the DOC and PPT legs. |
| `callgrind-xls/incl-head-<stem>-<op>-s<N>.txt` | `callgrind_annotate --inclusive=yes` for each XLS isolation pair. Every inclusive percentage in the record is read from these. |
| `callgrind-xls/` | 91 `callgrind_annotate` outputs: `{self,incl,tree}-head-<stem>-<op>-s<N>.txt` for each XLS isolation pair, 30 of them `--tree=caller`. |
| `callgrind-docppt/` | 108 equivalents for the six DOC and PPT fixtures. |
| `analysis/self-ir-per-op.txt` | Per-symbol self Ir per operation for the DOC/PPT legs. |
| `analysis/memcpy-memset-callers.txt` | Caller attribution for the bulk byte-movement symbols, the basis of the zero-fill candidate's per-site shares. |
| `analysis/sst_walk.py` | A byte-layout model of the shared-string chain walk, written independently of the Rust. It parses the FAT, the `Workbook` stream, the `SST` record with its `Continue` segments and every `LabelSst` index in stream order, converts each string's source offset to a sector ordinal and sums the walk. Runs no repository code and needs no build. |
| `analysis/sst-corpus-full.txt` | That model's output over every `.xls` fixture under `test-data/`. |
| `analysis/sheet_cursor_model.py` | The same treatment for the per-sheet cursor term, derived from `BoundSheet8` starts. Imports `cfb_workbook` from `sst_walk.py`, so run it from this directory. |
| `analysis/sheet-cursor-model.txt` | That model's output for the three fixtures the record names. |
| `analysis/xls-labelsst-census.txt` | An independent BIFF record census of the two primary fixtures, used to check the model's resolve counts. |
| `harness/` | The throwaway DOC/PPT driver source and its manifest patch. No DOC or PPT profiling harness existed before this change. |
| `scripts/` | Every capture and analysis driver used. |
| `capture.log` | Raw capture transcript. |

## Provenance

Base revision `e927e139c0ed9883b8315af5ae5a035f2cb869a9`. The profiled binary was
built from a **detached git worktree of that revision outside the repository
working copy**, with its own `CARGO_TARGET_DIR`, because three files were being
edited in the working copy while this profile ran. `xls_source_attribution`
sha256 `99ad002bc83c2fe06fd1f0c1888e90f07334fca80d2764a5a4d0a9bf64e3bbd9`.

Host: AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws. rustc 1.95.0,
valgrind 3.26.0, `--release --locked`. Every child ran under `setarch -R` with
ASLR off, `taskset`-pinned, with `RAYON_NUM_THREADS=OMP_NUM_THREADS=1`.

## Method

Callgrind isolation pairs, per **symbol**. A large-sample and a small-sample
child are profiled and differenced symbol by symbol, then divided by the sample
delta. One-time process cost therefore cancels exactly — the harness's SHA-256
staging is 53% of the raw profile and vanishes entirely in the delta. This
extends change 0564's isolation method, which differenced program totals only.

## Reproducing an inclusive figure by hand

Every inclusive percentage in the record is a difference of two retained
annotations divided by the sample delta, then divided by the operation's
whole-operation cost. For example, `SourceBackedWorkbook::from_read_at` on the
flagship open, from `callgrind-xls/incl-head-flagship-open-s{20,220}.txt`:

```
(510,888,421 - 50,717,764) / (220 - 20) = 2,300,853 Ir per open
2,300,853 / 2,379,097                   = 96.71%
```

and `GlobalsBuffer::ensure` from the same pair:

```
(304,695,338 - 30,314,931) / 200 = 1,371,902 Ir per open
1,371,902 / 2,379,097            = 57.66%
```

Both reproduce the record's table exactly. The same arithmetic recovers every
other cell.

## Controls

| check | this profile | change 0579 retained |
| --- | ---: | ---: |
| flagship `open` whole-operation Ir | 2,379,097 | 2,380,011 (0.04%) |
| flagship `open` FAT chain links | 2,099 | 2,099 (exact) |
| flagship `open` reads / bytes | 53 / 565,201 | 53 / 565,201 (exact) |

## Noise

Two byte-identical runs on different CPUs differ by **635 Ir in 678,704,129**
(0.0001%). Maximum per-symbol drift is 0.47%, confined entirely to glibc malloc
internals; every `litchi_*` symbol drifts below 0.05%. The DOC and PPT legs carry
a third sample point as a linearity control, 0.036–0.418%.

## What this packet does not establish

Instruction counts rank **work, not latency**. Change 0579 is the standing proof
of the gap in this repository: it removed a dependent-load pointer chase worth
1.24% of instructions on `54016.xls` and 6.19% of cycles. Callgrind also counts
`rep movsb`/`rep stosb` one instruction per byte, so every `memcpy`/`memset`
share here is an upper bound in instruction terms rather than a cycle cost —
change 0574's caveat, unchanged.

No `perf stat`, cycle, cache-miss, branch-miss, wall-clock, allocation,
peak-RSS, cold-cache, concurrency or cross-platform measurement was taken for
this record. No candidate named in it is authorized by it; each needs its own
paired measurement before it is believed.

Two files were deliberately **excluded** from this packet after an audit. A
`--tree=caller` annotation named `tree-cfs-open.txt` and a shared `build.log` were
produced by a **different** concurrent investigation — an OOXML probe over
`ConditionalFormattingSamples.xlsx`, not the XLS fixture of the same stem — which
wrote into the same scratch directory. They are not evidence for this record and
are not retained here. The 199 annotation files, `PROFILE.txt`, `ranking.txt`,
`folded.json` and `capture.log` were each checked for the same contamination and
are clean: zero references to the other binary, and 140 to this one.

The raw `callgrind.out` files are **not** retained. They are regenerable from
`scripts/capture.sh`, `scripts/capture_docppt.sh` and the recorded binary hash,
and the annotations they produce are retained in full.
