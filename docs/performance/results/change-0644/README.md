# Evidence packet — change 0644 (OLE2 snapshot source-identity fence design)

Record: [`docs/performance/0644-ole2-snapshot-fence-design.md`](../../0644-ole2-snapshot-fence-design.md).
Outcome: **frozen design, no production code.** `performance_claim: none`.

## Provenance

| field | value |
| --- | --- |
| base commit | `c7326f68065edf6f2198ca3cb39c38c48cf00ed9` (branch `feat/office-format-completeness`) |
| branch | `perf/0644-ole2-snapshot-fence-design` |
| worktree | `/home/zhuhe/code/litchi-worktrees/0644` |
| before checkout | `/home/zhuhe/code/litchi-worktrees/before-c7326f680` (read-only, shared) |
| crates changed | **none** — `git diff --name-only c7326f680 -- crates/` is empty, and `diff -rq` against the before checkout's `crates/` is empty |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), `cargo build --release` |
| CPU pin | `taskset -c 19` for every measured process |
| quiescence | **not established** — eight agents were building concurrently on the same 32-core host; the A/A leg in `perf/perfstat-legB.tsv` is the in-window floor |
| probe, built against the before checkout | sha256 `6cc68b61ba5a48d106f61aafaee949748b8904e759eb5e640505a66c2d20e46d` |
| probe, built against this worktree | sha256 `18b49f5032f9104418cd34eee00818b402668be24a9189c1d0089da1c4a82ddc` |
| `probe/main.rs` | sha256 `35251d10232406bc145377b88dc814143c19c03485bd3bfca08ae9e21a900e57` |

The two binaries differ (different dependency paths compiled in) but produce
byte-identical output; the cross-check is the last section of `gates.txt`.

## Contents

| path | what it is |
| --- | --- |
| `probe/main.rs` | the scratch driver. Six modes: `trace` (every `ReadAt` call of one operation, classified as a complete-scan chunk or a parse read), `sweep` (a held mutation after every read ordinal, with an optional forced witness offset), `transient` (one flip-then-revert witness), `census` (per-fixture admission, reads and bytes), `filever` (the `FileSource` metadata witness), `profile` (a loop for `perf stat` isolation pairs) |
| `probe/Cargo.toml.tmpl` | the manifest; `@ROOT@` is replaced per leg with the checkout to measure |
| `trace/trace-documentProperties-doc-open.txt` | the 30-read inventory of a generic `SourceSnapshot::open`; 6.00 complete scans |
| `trace/trace-picture-doc-open.txt` | the same on a 1.45 MB artifact: 46 reads, 12 scan chunks, 6.00 complete scans |
| `trace/trace-documentProperties-doc-open-read.txt` | `open` plus `paragraph(0)`: 59 reads and **14.00** complete scans — the open's 6 plus `resolve`'s 8, four composed reopens and the five paragraph-resolution reads |
| `trace/trace-45543-ppt-open.txt` | `litchi_ppt::text_edit::SourceSnapshot::open`: 8 reads, 2.00 complete scans |
| `sweeps/sweep-documentProperties-doc-open.tsv` | change-under-read sweep over all 30 ordinals, witness byte outside every parsed range; the clean control and the trailing control are both in the file |
| `sweeps/sweep-documentProperties-doc-open-read.tsv` | the same over `open` plus `paragraph(0)`, 59 ordinals — the fence that receives Option C's relocated refusals |
| `sweeps/sweep-documentProperties-doc-open-index-region.tsv` | witness byte inside the directory sector (offset 8,800); identical to the payload sweep on this fixture |
| `sweeps/sweep-documentProperties-doc-open-fat-region.tsv` | witness byte inside the FAT sector (offset 520), so a flip corrupts the chain; this is the sweep that resolves which parse reports which error wrapper (witness W-D2) |
| `sweeps/sweep-45543-ppt-open.tsv` | the PPT sweep, 8 ordinals; triggers 4–7 are refused by the internal read-twice-compare and by nothing else |
| `sweeps/sweep-45543-ppt-open-index-region.tsv` | witness byte inside the PPT allocation-table region (offset 381,028, in FAT sector 743; the directory chain starts at byte 384,000); triggers 1–5 are refused, and for 2–5 the composed reopen is the only detector (witness W-D1) |
| `witness/filesource-version-witness.txt` | witness W-A: an in-place `pwrite` through a second descriptor with the modification time restored leaves `SourceVersion` and `len()` unchanged |
| `witness/transient-doc-flip5-revert10.txt` | witness W-T1: flip after R1, revert after R2 — refused today, the artifact ends byte-identical |
| `witness/transient-doc-flip5-revert9.txt` | witness W-T2: flip after R1, revert one read earlier — accepted today; the control that bounds W-T1's claim |
| `witness/transient-docread-flip31-revert36.txt` | witness W-T3: the same shape on the call the adopted design actually reduces — `open` plus `paragraph(0)`, flip after r31 (the first `ensure_current`'s planning scan), revert after r36 (its confirming scan) — refused today at 36 reads, artifact byte-identical |
| `witness/transient-docread-flip31-revert35.txt` | witness W-T4: revert one read earlier, before the confirming scan — accepted today in 59 reads; the control that bounds W-T3 |
| `counts/census-doc.tsv` | all 57 `.doc` under `test-data/`: bytes, read calls, read bytes, complete-artifact reads, `len`/`version` counts, and the outcome |
| `counts/doc-fixtures.txt`, `counts/ppt-fixtures.txt` | the 8 admitted `.doc` and the 30 `.ppt`, copied from change 0589's packet so both records measure the same corpus |
| `perf/perfstat-legA.tsv` | native cycles and instructions per operation, `perf stat -r 3` isolation pairs, all 38 fixtures |
| `perf/perfstat-legB.tsv` | the A/A leg in the same window |
| `perf/perfstat-cfb-index.tsv` | one `SharedOleFile::open` per fixture — the index-parse term Options D and E would buy |
| `perf/predicted-savings.txt` | the output of `scripts/predict.py`: the per-scan constant by OLS and by two-point interpolation with residuals, the A/A floor, the predicted after-value per fixture for the three variants B3+C, B2+C and B1+C, the readback saving, and the index-parse term |
| `scripts/perfstat.sh` | one isolation pair: profile N and N+M operations, difference the totals, divide by M |
| `scripts/perfall.sh` | the same across the 38-fixture corpus, driven by `counts/doc-fixtures.txt` and `counts/ppt-fixtures.txt`; takes the probe binary as its argument and resolves everything else relative to the packet |
| `scripts/predict.py` | produces `perf/predicted-savings.txt`; runs from the packet root with no arguments and no absolute paths |
| `gates.txt` | the tail of every gate |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |

## Replay

```sh
ROOT=/home/zhuhe/code/litchi-worktrees/before-c7326f680
mkdir -p /tmp/probe/src && cp probe/main.rs /tmp/probe/src/main.rs
sed "s#@ROOT@#$ROOT#" probe/Cargo.toml.tmpl > /tmp/probe/Cargo.toml
( cd /tmp/probe && CARGO_TARGET_DIR=/var/tmp/0644 cargo build --release )
B=/var/tmp/0644/release/fence_probe_0644
T=/home/zhuhe/code/litchi/test-data

taskset -c 19 $B trace  doc-open      $T/ole/doc/documentProperties.doc
taskset -c 19 $B trace  doc-open-read $T/ole/doc/documentProperties.doc
taskset -c 19 $B trace  ppt-open      $T/poi/test-data/slideshow/45543.ppt
taskset -c 19 $B sweep  doc-open $T/ole/doc/documentProperties.doc          # payload witness
taskset -c 19 $B sweep  doc-open-read $T/ole/doc/documentProperties.doc
taskset -c 19 $B sweep  doc-open $T/ole/doc/documentProperties.doc 520      # FAT-sector witness
taskset -c 19 $B sweep  ppt-open $T/poi/test-data/slideshow/45543.ppt 381028
taskset -c 19 $B transient doc-open      $T/ole/doc/documentProperties.doc  5 10 8703
taskset -c 19 $B transient doc-open      $T/ole/doc/documentProperties.doc  5  9 8703
taskset -c 19 $B transient doc-open-read $T/ole/doc/documentProperties.doc 31 36 8703
taskset -c 19 $B transient doc-open-read $T/ole/doc/documentProperties.doc 31 35 8703
taskset -c 19 $B filever $T/ole/doc/documentProperties.doc
taskset -c 19 $B census  doc-open $T doc
./scripts/perfall.sh $B             > perf/perfstat-legA.tsv   # from the packet root
./scripts/perfall.sh $B cfb-index   > perf/perfstat-cfb-index.tsv
python3 scripts/predict.py          > perf/predicted-savings.txt
```

All three scripts run from the packet root and contain no absolute paths:
`predict.py` takes no arguments, `perfall.sh` takes the probe binary, and
`perfstat.sh` takes `<binary> <mode> <fixture> <N1> <N2>`. `perfstat.sh` pins to
CPU 19, which is this run's assignment; change the `taskset` line to measure
elsewhere.

## What the packet does not contain

* **No after leg.** Nothing was implemented, so there is nothing to compare
  against `perf/perfstat-legA.tsv`. The predicted after-values in
  `perf/predicted-savings.txt` are modelled from the measured per-scan constant.
* **No sweep of the changed code.** The record's costs B-ii (the trailing window)
  and B-iii (the error-precedence change the design relocates rather than
  accepts) are derived from the traced read order and from where
  `finish_overlay_plan_with_owner` places the composed reopen; no modified
  `litchi-cfb` was built or swept. Gate G1 is what would prove them.
* **No publication differential.** Change 0589's 87-artifact differential is
  gate G2 of the record, to be run by an implementing change; rerunning it here
  would compare a tree with itself.
* **No callgrind, allocation, RSS, syscall, cold-cache or paired-timing data.**
  The design's claims are about read counts and refusal identity, which are
  deterministic, plus one native cycles baseline. Paired timing adds nothing to
  a record with one leg.
* **No Windows or non-SHA-NI measurement**, and no fixture above 1.45 MB.
