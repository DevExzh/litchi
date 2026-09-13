# Large-cell/tag guard design for change-0552

**Status: prospective design only.** This document defines a supplementary
OOXML diagnostic for the transient attribute decoding called out in the
[0552 proof review](proof-review.md). It records no performance result and
does not add an admission gate. It is conditional on the representative
0552 pilot passing its existing native, allocation, correctness, and memory
checks. OLE2/OOXML remains the active priority; ODF work can wait for that
goal.

## Feasibility with the retained binaries

The retained binaries cannot exercise the requested input. The current
retained set contains the main `litchi-perf-baseline` normal/allocator
executables and the planning/cap guard executables under
`/home/zhuhe/litchi-goal-0552-target/retained/baseline/`; there is no retained
candidate pair for this custom case. More significantly, their inputs are
fixed in the binaries:

| Existing entry point | Input it accepts | Why it is insufficient |
| --- | --- | --- |
| `litchi-perf-baseline` XLSX child | fixed synthetic four-sheet cell-value shapes and cases | no worksheet or ZIP fixture argument |
| `xlsx_planning_guard` | `--shape medium\|dense-sparse` and fixed `valid\|late-validator\|late-raw` grids | measures planning only and cannot insert a large attribute |
| `perf_cap_boundary` | fixed numeric grids and `--size 1\|2\|160\|164\|256` | exercises event/source cap boundaries, not tag normalization |

Consequently, a retained main binary cannot be pointed at an 8 MiB worksheet,
and a future retained candidate binary would have the same limitation. The
guard needs a separately built, hash-bound harness and one normal plus one
allocator binary for each source revision. This is a custody/design
requirement, not a request to change the frozen `run.py`, `capture.py`,
`plan.json`, analyzers, or main gates.

## Proposed standalone harness

Place the harness in
`docs/performance/results/change-0552/large-tag-guard/` as its own Cargo
workspace. It should contain a small `Cargo.toml`, its own `Cargo.lock`, a
fixture generator/specification, `src/main.rs`, and a guard-local runner.
Use path dependencies on `litchi-core`, `litchi-xlsx`, and
`soapberry-zip`; reuse the public `allocation_metrics` module from
`tools/perf-baseline` and include its existing
`src/bin/support/counting_allocator.rs` by path for the allocator binary.
That keeps the only unsafe allocator boundary and counter semantics shared
with the existing evidence. The harness must not edit those files.

The runner must accept only an immutable fixture/spec and an explicit binary
path. Before every invocation it should verify the binary SHA-256, the
fixture/spec SHA-256, its own source manifest, the separate Cargo lock, the
production baseline or candidate `source-manifest.json`, the repository
ignored workspace lock, and the host/toolchain identity. Emit one receipt per
child containing the command, environment, hashes, phase samples, stderr,
`/usr/bin/time -v` output, and correctness result. Keep evidence under a
separate `large-tag-guard/runs/{baseline,candidate}/` tree.

The standalone binary should call the same public path as the existing guard:
`SourceBackedEditor::from_read_at`, `edit_sheets`, one existing-cell
`MultiSourceEdit::set`, `commit`, and
`publish_multi_commit_to_stream`. Fixture creation, ZIP parsing needed to
obtain expected hashes, selectors, sink reservation, and all expected-output
oracles belong outside the measured regions. The public API does not expose
whether the candidate selected the compact writer, so route acceptance must
come from the existing private compact-versus-complete differential tests and
source review; this supplement must not infer it from a missing scan symbol.

## Fixture grammar and byte shape

Generate a deterministic five-member stored ZIP, matching the existing
planning guard's compact raw package: `[Content_Types].xml`, `_rels/.rels`,
`xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, and
`xl/worksheets/sheet1.xml`. The worksheet must contain no formatting
whitespace, because the publication validator rejects an authored XML part
with inter-element whitespace. Use a two-row, two-cell sheet so the changed
and unchanged controls have the same small semantic workload.

The accepted worksheet grammar is deliberately within the current strict
validator:

* `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`
  contains `<dimension ref="A1:B2"/>` and `<sheetData>`.
* `<row>` has `r` and, in the row-tag variant, `spans`.
* `<c>` has `r`, an unused namespace declaration `xmlns:q`, and a scalar
  `<v>` child. Namespace declarations are accepted by the validator and are
  ignored by the raw cell-attribute semantics.
* The source has no formulas, inline strings, MCE, `x14ac`, `dyDescent`,
  vendor extension elements, comments, or dependency-bearing worksheet
  elements. XML text occurs only in scalar `<v>` elements.

Use this compact shape (line wrapping here is explanatory; generated bytes
must be contiguous):

```xml
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B2"/><sheetData><row r="1"><c xmlns:q="urn:litchi:large:&#x61;&#x61;..." r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row></sheetData></worksheet>
```

The generator should choose the number of `&#x61;` terms so that the
worksheet is exactly a recorded target such as
`8 * 1024 * 1024 - 128 * 1024` bytes, and assert that it is below
`MAX_SHARED_SOURCE_BYTES` (8 MiB). Keep the ZIP and output below their package
limits as well. `&#x61;` is six source bytes and decodes to one `a`; repeating
it makes `decoded_and_normalized_value` take the owned, normalization-heavy
attribute path while retaining valid UTF-8. Record the exact worksheet,
archive, and output-budget sizes. The source must remain far below the
131,072 shared provisional-event cap and the 1,000,000 raw XML-event cap.

Use two tag families:

1. **Cell tag:** Put the generated entity sequence in the unused `xmlns:q` on
   `A1`, with the namespace declaration before `r`. In the changed case the
   writer must materialize this cell tag; in the control case it must copy the
   same source span while changing `B1`.
2. **Row tag:** Put the sequence in `row spans` on row 1 and use the same
   `A1`/`B1` changed-versus-unchanged actions. The raw parser currently
   accepts and ignores `spans`, while the complete scanner still decodes and
   retains the row tag. This variant covers the neighboring `tag()` path
   without adding unsupported worksheet semantics.

For each family, use a late-refusal copy of the fixture by appending a final
`<c r="B2" future="1"><v>5</v></c>` after the large valid tag. The expected
typed error is exactly `Invalid` with message
`value-only edits refuse attribute 'future' on 'c'`. If the harness also runs
the existing late-raw control, append a boolean cell whose value is
`maybe`; its expected message is exactly `invalid worksheet boolean 'maybe'`.
Both refusals must still occur after the large attribute has been decoded.

## Cases and correctness oracle

The minimum matrix is the following. Every row uses the same generated bytes
for baseline and candidate.

| Fixture/action | Purpose |
| --- | --- |
| cell tag, set `A1` from `1` to `9` | changed large cell tag; exercises the changed-cell materialization |
| cell tag, set `B1` from `2` to `9` | unchanged large cell control; `A1` should be copied from its source span |
| row tag, set `A1` from `1` to `9` | row-tag variant with a changed cell |
| row tag, set `B1` from `2` to `9` | row-tag unchanged control |
| large valid tag followed by `future` | late validator refusal and retry stability |
| large valid tag followed by `t="b"`/`<v>maybe</v>` (optional) | late raw-parser refusal and retry stability |

For every accepted action, retain the commit and bounded output sink until
the phase clock and allocator region close. Check all of the following:

* baseline and candidate return the same changed-cell count, worksheet count,
  semantic values, and styles;
* the complete baseline writer's output and the candidate output are
  byte-for-byte identical, including ZIP member order and worksheet bytes;
* the input archive SHA-256 is unchanged, all untouched members are preserved,
  and the unchanged large `A1` cell (or row opening tag in the row family)
  has the expected source-span bytes;
* the output sink reports the same accepted byte count, write-call behavior,
  and maximum 64 KiB write bound; and
* the candidate output is also checked by the existing semantic readback
  oracle. A normalized entity spelling in a changed tag is therefore judged
  against the baseline bytes, not against the original lexical spelling.

For every refusal, require the same typed variant and exact message on both
revisions and on an immediate retry, unchanged input SHA-256, no commit, and
zero published output. Preserve the full serialized error fields and source
identity in each receipt. A mismatch or an unverified fixture is a failed
guard result; timing data from that row must not be presented as evidence.

## Timing, allocation, and RSS boundaries

Run normal and allocator binaries separately. For each fresh child process,
report these phases:

* `open`: `from_read_at` only;
* `plan`: `edit_sheets` only, which includes strict validation, ordinary
  parsing, and compact-proof tag decoding;
* `commit`: `set` plus `transaction.commit`, retaining the result through the
  clock;
* `publication`: `publish_multi_commit_to_stream` into a pre-reserved,
  retained bounded sink; and
* `workflow`: a direct end-to-end timer from a fresh editor through retained
  publication. Do not derive it by subtracting phase medians.

For allocator runs, surround each phase with
`allocation_metrics::begin()`/`finish()` and retain phase results until
`finish` returns. Record allocation, reallocation, deallocation, failed-call,
allocated/deallocated-byte, live-before/after, and attributable-region-peak
fields. The allocator-instrumented elapsed time is diagnostic only and must
never be used as native latency evidence. Add `/usr/bin/time -v` per child and
retain maximum RSS as a process-level measure.

Use the existing 0552 schedule: two repeats, native warmup 20 and 200
samples, allocator warmup 3 and 20 samples. Run ABBA within each fixture,
action, and binary kind: `baseline r1`, `candidate r1`, `candidate r2`, then
`baseline r2`. Pin each child to CPU 2, use the campaign temporary directory,
and keep build settings (`jobs=2`, `incremental=0`) identical. Report raw
vectors plus p50, mean, p95, p99, and maximum for every phase and both
revisions.

This is a diagnostic guard, so it has no 3% improvement, 1.05x, or other
performance admission threshold and must not be added to `plan.json` or the
main analyzer. Correctness, source custody, output bounds, and measured
counter availability are mandatory. Any adverse timing, allocation, or RSS
row remains in the report for review; it cannot be dropped or averaged away.

## Binary and evidence custody

Freeze a guard-local `inputs.json` before either supplementary build. It must
bind every harness file, the guard lockfile, fixture generator/spec, the
0552 plan and proof review, the exact baseline and candidate production
source-manifest hashes, `workspace-lock.json` and its retained Cargo lock,
the toolchain, host, and intended binary paths.

Restore the exact unchanged revision and build/copy immutable
`large-tag-guard/baseline/{normal,alloc}` binaries. Then build the final
candidate source and copy immutable
`large-tag-guard/candidate/{normal,alloc}` binaries. Never overwrite a
retained binary; receipts must include each binary hash and the source
manifest used to build it. A guard-local runner can reuse the source and
lock checks from `check_attempt.py`, but it must additionally hash the
standalone harness because the main checker inventories production source,
not arbitrary files under `docs/`.

Prefer committing and freezing the standalone harness before any main
capture. If it is prepared after the main artifacts are sealed, run it only
as a separate evidence bundle and add its immutable custody records without
editing the frozen main scripts, captures, gates, or source. Do not let an
unbound untracked harness become part of a main candidate identity. The guard
is complete only when its source, fixture, binary, host, and receipt hashes
can be replayed independently.
