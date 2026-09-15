# Evidence packet: change 0610

Size candidate C2′ — lazy OPC part decode behind the fallible package accessors
— by measuring what an ordinary open-then-one-edit-then-save actually reads
against what `load_parts_eager` inflates today, enumerate its migration surface,
and draft the proposed ADR that gate 1 of change 0581 requires.

Record: [`../../0610-opc-lazy-part-decode-design.md`](../../0610-opc-lazy-part-decode-design.md).
Proposed ADR: [`../../../adr/0030-lazy-opc-part-decode.md`](../../../adr/0030-lazy-opc-part-decode.md).
Implements SAVE-5 of [0587](../../0587-remaining-opportunity-survey.md); satisfies
gate 1 of [0581](../../0581-opc-package-retention.md) as far as a draft can.

**Design only. No file under `crates/` was modified. `performance_claim: none`,
and no timing, cycle or instruction figure is measured or claimed — every number
in this packet is a deterministic count of parts or bytes.**

## Provenance

| field | value |
| --- | --- |
| base commit | `818e58bee20a9d95bbbda5bf3e1f60aa4db45b3f` |
| branch | `perf/0610-opc-lazy-part-decode-adr-draft` |
| host | AMD EPYC 9R45, 32 cores, 123 GiB, Linux 7.0.0-1012-aws |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 |
| probe build | `cargo build --release` in a scratch crate with path dependencies on `litchi-opc` and `litchi-xlsx`; profile `lto = true`, `panic = "abort"`, mirroring the workspace |
| pinning | every probe process `taskset -c 30` |
| host state | seven other agents were building and measuring concurrently; run-window load average 26–44 on 32 cores. This does not affect the figures: they are counts, not timings. |

Binary SHA-256:

| binary | sha256 |
| --- | --- |
| `opc-touch-probe` | `e895d92cf46a0bd0db720452eebe91754dd7a5768de7c6ad91b74f2d6b4d2a7b` |

## Contents

| path | what it is |
| --- | --- |
| `probe/src/main.rs` | the sizing probe: `census` sums what `load_parts_eager` inflates per fixture; `touch` measures what an operation reads, by ablation |
| `probe/Cargo.toml.example` | the probe manifest; point the two path dependencies at the checkout being measured and give it its own `CARGO_TARGET_DIR` |
| `probe/run-touch-corpus.sh` | the corpus ablation driver that produced the two `touch-*.txt` files |
| `probe/classify-iter-parts.py` | the `iter_parts()` call-site classifier that produced `iter-parts-classification.txt` |
| `census.txt` | one row per OOXML fixture: archive bytes, admitted parts, inflated bytes, XML parts, XML bytes, largest part. 336 rows, 2 of them open refusals |
| `touch-opc-reblob.txt` | ablation summary for the physical route, one row per fixture: 325 measured, 11 refused |
| `touch-xlsx-hide.txt` | ablation summary for the documented XLSX editor route: 33 measured, 147 refused with the reason |
| `touch-xlsx-hide-parts.txt` | the part each of the 33 measured fixtures reads — `/xl/workbook.xml` on every one |
| `touch-flagship/` | the per-part ablation detail for the flagship fixtures, including the `untouched` verdict for each individual part |
| `iter-parts-classification.txt` | the per-crate, per-shape classification of all 307 `iter_parts()` sites under `crates/*/src`, in two scopes, with the out-of-scope source-backed sites and the 77 scope disagreements listed |
| `blob-site-counts.txt` | the `.blob()` and `iter_parts()` population counts with the exact `ripgrep` commands, and the drift against 0581's and 0587's figures |
| `run-gates.sh` | the gate runner that produced `gates.txt` |
| `gates.txt` | the tail of every gate |
| `decision.json` | the machine-readable decision, accepted evidence, accepted costs and known gaps |
| `log-sections.md` | the four log paragraphs for HOTSPOTS, GOAL\_AUDIT, REPORT and ADR\_COMPLIANCE |
| `cleanup.json` | what this batch retained in the repository and what it removed afterwards |

## Replaying

```sh
# build the probe (scratch crate, path dependencies, own target dir)
cp probe/Cargo.toml.example /scratch/probe/Cargo.toml   # edit the two paths
cp probe/src/main.rs        /scratch/probe/src/main.rs
CARGO_TARGET_DIR=/scratch/target cargo build --release --manifest-path /scratch/probe/Cargo.toml

# what the eager open inflates
taskset -c 30 opc-touch-probe census test-data > census.txt

# what one operation reads
taskset -c 30 opc-touch-probe touch <fixture> opc-reblob      # per-part detail
taskset -c 30 opc-touch-probe touch <fixture> xlsx-hide
probe/run-touch-corpus.sh <probe-binary> <checkout> <out-dir> 30   # whole corpus

# the migration surface
python3 probe/classify-iter-parts.py <checkout> 8
```

Both `touch` operations are `O(parts)` opens, edits and publishes per fixture;
the whole-corpus run over 336 fixtures and 5,077 parts completes in minutes.

## Reading the numbers

`census.txt` columns are
`fixture archive_bytes parts inflated_bytes xml_parts xml_bytes largest_part_bytes`;
a refused fixture reads `<path> <archive_bytes> open-refused - - - - <error>`.

`touch-*.txt` summary rows are
`fixture operation archive_bytes parts inflated_bytes touched touched_bytes`;
a fixture the operation cannot run on reads
`# <path> <operation> baseline-refused <reason>`.

`touched` is the number of parts whose payload the operation reads, measured by
substituting a sentinel payload for one part at a time and comparing the
operation's outcome everywhere except that part's own member. It does **not**
include a read whose only effect is on the ablated part itself, which is why the
physical route's own re-blob read is stated separately in the record rather than
counted here.

## Correction, after the first version of this packet

The first version reported the `iter_parts()` migration surface as "267
production sites, 2 of them on `SourceBackedPackage`, 17 reading payload bytes".
That came from a classifier with two defects, and the corrected figures are
**259 in-scope production sites, 27 on `SourceBackedPackage`, and a
payload-reading group bracketed 16–88 with an independent estimate of about 25**.

The defects were:

- It treated the **first `#[cfg(test)]` line in a file** as the start of an
  inline test module. A `#[cfg(test)] mod name;` *declaration* leaves the rest of
  the file production code, and `crates/litchi-opc/src/package.rs:1238` is
  exactly that case — so the production site at `:1269` was wrongly counted as a
  test.
- It classified on a **fixed 8-line window**. `package.rs:1269` reads
  `part.blob_arc()` 17 lines later and `pkgwriter.rs:140` reads `part.blob()` 14
  lines later; both were classified metadata-only.

The corrected classifier in `probe/` reports two scopes rather than one, skips
only genuinely inline test modules, and separates `SourceBackedPackage` sites,
whose `PartView` has no `blob()` and whose `data()` is already fallible.

**No measured figure in this packet changed.** The census, both ablation runs and
the retention corroboration are produced by the probe, not by the classifier, and
are untouched. The design, the proposed ADR's decision content and the
admission-gate analysis are unaffected; the corrected numbers strengthen the case
rather than weaken it, because 27 of the sites turn out to be on the already-lazy
type.
