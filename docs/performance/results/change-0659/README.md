# Evidence packet — change 0659 (B2 + C on the OLE2 snapshot source-identity fence)

Record:
[`docs/performance/0659-cfb-single-scan-identity-entry-point.md`](../../0659-cfb-single-scan-identity-entry-point.md).
Outcome: **retained, implemented.** `performance_claim: none`.

## Provenance

| field | value |
| --- | --- |
| base commit (both legs' source) | `70d7768cc6dada420ede063f72c88dc99ad30383` (`docs(perf): record the owner's decisions for the third wave and accept ADRs 0030 and 0031 (0652)`) |
| branch | `perf/0659-cfb-single-scan-identity-entry-point` |
| after leg | `/home/zhuhe/code/litchi-worktrees/0659` (this branch) |
| before leg | `/home/zhuhe/code/litchi-worktrees/before-70d7768cc` (read-only, shared, detached at the base) |
| crates changed | `litchi-cfb` (two public entry points and one private helper), `litchi-doc` (`body_text/source.rs`). **`litchi-ppt` unchanged.** |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), `cargo build --release` |
| CPU pin | `taskset -c 14` for every measured process |
| quiescence | **not established** — eight agents were building concurrently on the same 32-core host; the A/A legs in `perf/` and `bench/` are the in-window floor |

### Probe binaries (sha256)

Two probes, both **reused verbatim** from earlier packets so that the before leg
can be checked against their retained outputs.

| probe | source | source sha256 | before binary | after binary |
| --- | --- | --- | --- | --- |
| fence probe (trace, sweep, transient, census, filever, profile) | `results/change-0644/probe/main.rs` | `35251d10232406bc145377b88dc814143c19c03485bd3bfca08ae9e21a900e57` | `9d3f9575d6a7e7835f030e0fd43bd88315c295b38866700293632b0c1f15a60c` | `0fafa8677cac149d3cbf8d2c700f9dc687d564891b7c942083dd46d9ddf27c55` |
| differential and timing probe (digest, bench) | `results/change-0589/probe/main.rs` | `87d17da72f7082da49c6f9d0b5ef5343d9aae515b935bfd7697ca149e292c6a7` | `06b10086d0ae7a9c319c738d323b7fe3723b5dfd4d0a3a1fc7029dbf04ef5f1f` | `936b57273b22b8d2f1217d5431840d1b793729c1ce489008b71e191ea259119a` |

Neither probe's source is copied into this packet; both are already in the
repository at the paths above, and the sha256 column is what pins the reuse.
`lto` is off in these manifests, but the release binaries are still not
reproducible byte for byte between legs (different dependency paths are compiled
in), so the digests above are of the binaries actually measured.

### Two measurement windows, and which one the record quotes

Everything was measured twice. A code review after the first window asked for
three behaviour-preserving cleanups (the `SourceSnapshot` construction folded
into one accessor shared with two other planners, a two-variant enum dropped,
and two rustdoc sentences corrected), so the after binaries were rebuilt and
every leg re-run. The cleanups changed nothing observable: all **156**
deterministic after-leg outputs — traces, sweeps, censuses, witnesses — are
byte-identical between the two builds, and so is the 87-artifact differential.
The second window is the one this packet retains and the record quotes; it was
the quieter of the two (A/A p95 0.62% on cycles against 0.78%, and 0.71% on
timing against 2.20%). The first window's medians were −34.73% for the DOC open,
−24.34% for the readback and +0.25% for PPT, against the quoted −35.41%, −24.96%
and −0.12% — agreeing to within 0.7 percentage points, which is a
between-window repeatability check the record does not otherwise have.

### The before leg reproduces its predecessors exactly

| check | result |
| --- | --- |
| all **4** traces against `results/change-0644/trace/` | identical |
| all **6** named sweeps against `results/change-0644/sweeps/` | identical |
| all **4** transient witnesses and the `FileSource` witness against `results/change-0644/witness/` | identical |
| `counts/census-doc-before.tsv` against `results/change-0644/counts/census-doc.tsv` | identical |
| `differential/digest-before.tsv` against `results/change-0589/differential/digest-before.tsv` | identical |

So nothing on this path moved between change 0644's base (`c7326f680`) and this
one, and the before halves of gates G1–G4 are this packet's own measurements
rather than a reuse of older numbers.

## Contents

| path | what it is |
| --- | --- |
| `trace/trace-*-{before,after}.txt` | the read inventory of one operation, every `ReadAt` call classified as a complete-scan chunk or a parse read: DOC `open` on `documentProperties.doc` (30 → 23 reads, 6 → 3 scans) and `picture.doc` (46 → 34, 6 → 3), DOC `open` + `paragraph(0)` (59 → 50, 14 → 9), PPT `open` (8 → 8, 2 → 2) |
| `trace/scans-{before,after}.tsv` | the same two counts for **all 38 fixtures** in three modes — gate G3's corpus half |
| `sweeps/sweep-*-{before,after}.tsv` | change 0644's six named change-under-read sweeps on both legs: DOC `open` and `open`+`paragraph(0)` with the payload witness, DOC `open` with the directory-sector and FAT-sector witnesses, PPT `open` with the payload and allocation-table witnesses |
| `sweeps/manifest.tsv` | sha256 of **every** sweep on both legs, 144 per leg, with a byte-identity verdict: all **92** PPT files identical, all 52 DOC files different (they have fewer ordinals) |
| `sweeps/bands-doc-corpus.txt` | the run-length bands of every `.doc` sweep on both legs — the compact form of 6,453 mutated opens, and where the record's band tables come from |
| `sweeps/gate1-sweep-table.txt` | per sweep pair: the leading `OK` window, the trailing `OK` window, the ordinal and refusal counts, and every outcome-class count that changed. Ends with gate G1's mechanical verdict |
| `counts/census-{doc,ppt,doc-read}-{before,after}.tsv` | all 57 `.doc` and all 30 `.ppt`: bytes, read calls, read bytes, complete-artifact reads, `len`/`version` counts and the outcome |
| `counts/census-*-delta.tsv` | the same as before → after pairs, with the admitted/refused summary lines the record quotes |
| `counts/{doc,ppt,doc-read}-fixtures.txt` | the fixture lists, copied from change 0644's packet so both records measure the same corpus |
| `witness/transient-*-{before,after}.txt` | change 0644's four flip-then-revert witnesses W-T1…W-T4, rerun on both legs |
| `witness/filesource-version-witness-{before,after}.txt` | witness W-A: an in-place `pwrite` with the modification time restored leaves `SourceVersion` and `len()` unchanged |
| `perf/perfstat-leg{A1,B1,B2,A2}.tsv` | `perf stat -r 3 -e cycles,instructions` isolation pairs over the 38-fixture corpus, in the paired order before, after, after, before |
| `perf/perfstat-read-leg{A1,B1,B2,A2}.tsv` | the same for `open` + `paragraph(0)` on the four fixtures whose paragraph read is admitted |
| `perf/summary-{open,read}.txt` | per fixture: before, after, the delta, change 0644's predicted after-value, the residual against it, and the two A/A columns; then the medians and the floor |
| `perf/fit-{open,read}.txt` | the OLS decomposition of the measured saving into a per-byte term and a fixed term, with per-fixture residuals |
| `bench/bench-*-{A1,B1,B2,A2}.txt` | raw per-operation nanosecond samples, 120 per leg |
| `bench/bench-summary.txt` | p50, mean, p95 and p99 per leg, the paired delta in both directions, and the A/A floor |
| `differential/digest-{before,after}.tsv` | gate G2: change 0589's 87-artifact publication differential, rerun unchanged on both legs. **`diff` is empty** |
| `scripts/run-legs.sh` | runs the whole deterministic battery for one leg: 4 traces, 6 named sweeps, 144 corpus sweeps, 3 censuses, 5 witnesses |
| `scripts/cfboffsets.py` | derives a byte inside the first FAT sector and one inside the first directory sector from a CFB header, so the sweep can place its witness in the allocation table or the directory on any fixture |
| `scripts/sweepcmp.py` | run-length-encodes a sweep's outcome classes and compares the two legs |
| `scripts/gate1.py` | gate G1's mechanical half: leading window, trailing window, interior `OK`, and the per-class counts |
| `scripts/censuscmp.py` | the census before/after delta tables |
| `scripts/perfstat.sh`, `scripts/perfall.sh` | one isolation pair, and the corpus sweep of them (change 0644's, repinned to CPU 14) |
| `scripts/analyse.py` | gate G4's arithmetic: deltas, the prediction residual and the A/A floor |
| `scripts/fit.py` | the per-byte and fixed-term decomposition |
| `scripts/bench.sh`, `scripts/benchsum.py` | the paired-timing legs and their summary |
| `gates.txt` | the tail of every gate |
| `decision.json` | the machine-readable decision |
| `log-sections.md` | the four log paragraphs for the coordinator to merge |
| `cleanup.json` | what was deleted and what was kept |

## Replay

```sh
BEFORE=/home/zhuhe/code/litchi-worktrees/before-70d7768cc
AFTER=/home/zhuhe/code/litchi-worktrees/0659
SCRATCH=/var/tmp/0659            # anywhere disk-backed and outside a Cargo target

# the fence probe, one build per leg
for leg in before after; do
  case $leg in before) ROOT=$BEFORE;; after) ROOT=$AFTER;; esac
  mkdir -p $SCRATCH/probe-$leg/src
  cp $AFTER/docs/performance/results/change-0644/probe/main.rs $SCRATCH/probe-$leg/src/main.rs
  sed "s#@ROOT@#$ROOT#" $AFTER/docs/performance/results/change-0644/probe/Cargo.toml.tmpl \
    > $SCRATCH/probe-$leg/Cargo.toml
  ( cd $SCRATCH/probe-$leg && CARGO_TARGET_DIR=$SCRATCH/target-$leg cargo build --release )
done

# the deterministic battery, per leg  (about 3 seconds each)
./scripts/run-legs.sh $SCRATCH/target-before/release/fence_probe_0644 out/before 14
./scripts/run-legs.sh $SCRATCH/target-after/release/fence_probe_0644  out/after  14

python3 scripts/gate1.py     out/before/sweeps out/after/sweeps       # gate G1
python3 scripts/censuscmp.py out/before/counts/census-doc.tsv out/after/counts/census-doc.tsv

# gate G4, in the paired order A1 B1 B2 A2
./scripts/perfall.sh $SCRATCH/target-before/release/fence_probe_0644 > out/perf/perfstat-legA1.tsv
./scripts/perfall.sh $SCRATCH/target-after/release/fence_probe_0644  > out/perf/perfstat-legB1.tsv
./scripts/perfall.sh $SCRATCH/target-after/release/fence_probe_0644  > out/perf/perfstat-legB2.tsv
./scripts/perfall.sh $SCRATCH/target-before/release/fence_probe_0644 > out/perf/perfstat-legA2.tsv
python3 scripts/analyse.py out/perf/perfstat-leg{A1,B1,B2,A2}.tsv
python3 scripts/fit.py     out/perf/perfstat-leg{A1,B1,B2,A2}.tsv doc-open 3
```

`scripts/perfall.sh` and `scripts/perfstat.sh` carry this run's absolute
scratch paths and CPU pin (14); change the `taskset` line and the two `HERE`
assignments to measure elsewhere. Gate G2 and the paired timing use change
0589's probe, built the same way with `sha2 = "0.11"` added to the manifest for
its `digest` subcommand, then `snapfence_probe digest test-data/` and
`snapfence_probe bench <mode> <fixture> <warmups> <samples>`.

## What the packet does not contain

* **No copy of either probe's source.** Both are already retained, in
  `results/change-0644/probe/` and `results/change-0589/probe/`; the sha256
  table above is what pins that they were reused unmodified.
* **Not every corpus sweep's per-ordinal rows.** 144 sweeps per leg is 6,453
  mutated opens; `sweeps/` keeps change 0644's six named sweeps in full on both
  legs, plus a sha256 manifest of all 288 files and the run-length bands of
  every `.doc` sweep. `scripts/run-legs.sh` regenerates the rest in about three
  seconds per leg.
* **No B1 leg.** B1 was rejected on a witness drawn from the landed B2 sweeps;
  no B1 build exists, so nothing here measures what it would have cost or saved.
* **No cold-cache, physical-device, range-source, remote, cross-platform,
  allocation or RSS measurement**, and no callgrind: this change moves bulk
  SHA-256 and metadata syscalls, both of which callgrind misprices (change 0604
  for `rep movsb`, change 0649 for software SHA-256), so everything here is
  `perf stat` cycles or wall-clock nanoseconds.
