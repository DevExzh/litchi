# Change 0206: ODS settings pass materializes element names only for recorded spans

Date: 2026-08-19

## Decision

**Banked — allocation-count claim only, latency neutral.** The change
eliminates 4,164 allocations per source-open on the measurement corpus
(14,891 → 10,727 allocations per open, -27.96%; 1,974,783 → 1,928,711
allocated bytes, -2.33%), measured deterministically with a counting
allocator in the profiling driver (counts, not timing — exact and
reproducible, identical across repeated runs and rebuilds). Latency
readings accepted in both directions on several phases but every accepted
magnitude sits below the 0205-calibrated layout-noise floor, so no
latency claim is made; no adverse both-directions pattern exceeded any
floor, so nothing blocks.

## Mechanism and invariants

Post-0204 profiling of source-open attributes 4.38% of cycles (3.47% self)
to a single `String::from_utf8` chain: `settings::codec`'s location scan
(`LocateHandler::on_event` and the standalone `locate()` shell)
materialized the qualified name of EVERY Start/Empty element
(`settings/codec.rs` — handler and shell, two arms each), one `Vec`
allocation + copy + full UTF-8 validation + free per element. The locate
pass records spans only for the `office:spreadsheet` host and the
`table:calculation-settings` element (`record()` drops every other
kind's span immediately), and the only `qname` consumer anywhere is
`replace()`'s empty-spreadsheet expansion (`spreadsheet.qname`). All but
1–2 elements' name strings per document were dead work.

The change introduces `qualified_name(element, kind)` in
`settings/codec.rs`, called identically from all four sites (handler and
shell, Start and Empty arms): the name string is materialized only when
the kind is `Spreadsheet` or `Calculation`; other kinds get
`String::new()`.

Invariants:

- Recorded spans keep byte-identical `qname`, so `Location` contents and
  the `replace()` expansion are unchanged.
- The dropped `"ODS XML element name is not valid UTF-8"` error is
  unreachable for every input: the source is a `&str`
  (`NsReader::from_str`), and element names are subslices delimited by
  ASCII markup bytes, which cannot split a UTF-8 multi-byte sequence —
  so removing the conversion for non-recorded kinds removes no fireable
  error, and for recorded kinds it fires at the same position as before.
- Stack discipline, depth/root/event limits, attribute validation, and
  error precedence are untouched; the guard sits exactly where the
  unconditional conversion sat.
- Shell and handler changed identically, preserving the 0200 shell-oracle
  invariant; the existing open_parse equivalence tests (shell vs fused
  over the corpus) remain the decisive gate.

Verification: the full `litchi-ods` suite (369 tests) passes, including
the 0200 open_parse equivalence tests (shell vs fused over the .ods
corpus, both sides changed identically). fmt, clippy (`-D warnings`),
rustdoc (`-D warnings`), and `tools/check_crate_boundaries.py` pass.

## Matched release timing

Two frozen release binaries differ only in the lazy qualified-name
materialization; both carry changes 0192-0196, 0198-0202, and 0204.
Control SHA-256 `5f0dab648cdf8f693ec01c171bfb81b2776e9e5abfe117b5be85c42a5bd89f66`
(the banked 0204 binary), candidate SHA-256
`8e17ab3e9857cb5c7d6b28ea31ef85ec7512e779270f72bd11361c980e1a0eb8`
(tree verified to rebuild bit-exact to the candidate after banking).
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg, drift ceilings
5%/5%/10%/15% (p50/mean/p95/p99). The 0205 floor rule applies: accepted
statistics below the calibrated floor are neutral, not claims.

### ods_file_source_open (the executed phase)

| stat | A1→B1 | A2→B2 | control drift | candidate drift | verdict |
|---|---:|---:|---:|---:|---|
| p50 | 1.22% | 3.91% | 2.43% | -0.37% | accepted, below floor 5.5% → neutral |
| mean | 1.01% | 3.32% | 1.47% | -0.89% | accepted, below floor 5.5% → neutral |
| p95 | 1.10% | 0.98% | -2.57% | -2.45% | accepted, below floor 4.5% → neutral |
| p99 | 4.09% | 1.94% | -9.47% | -7.44% | accepted, below floor 36% → neutral |

### ods_file_eager_open (no changed code)

p50 accepted (2.35%/0.23%, below floor 3.0% → neutral); mean/p95/p99
withheld on disagreeing paired directions. No adverse pattern.

### ods_source_backed_one_edit_save

lifecycle p50/mean/p95 accepted (0.51%-3.08%, below floor 2.2%/2.8% →
neutral), p99 withheld; commit mean/p95/p99 accepted (0.53%-16.80%,
below floor 3.7%/5.8%/17% → neutral), p50 withheld. No adverse pattern.

### ods_source_backed_one_percent_edit_save

lifecycle p50/mean accepted (~1%, below floor 2.4% → neutral); commit p50
accepted (~1%, below floor 3.1% → neutral); commit p99 adverse in both
directions (-2.96%/-12.18%) but within the 13.5% p99 floor → layout
reading, does not block. Remaining statistics withheld on disagreeing
directions.

### ods_source_backed_repeated_edit

total mean/p95/p99 accepted (0.11%-1.35%, below floor 1.8%/2.5%/2% →
neutral); stage withheld on disagreeing directions; commit p99 accepted
(below floor 7.5% → neutral); publication all-four accepted
(0.02%-1.04%, at/below floor 1.1%/2%/1% → neutral). No adverse pattern.

### Allocation evidence (deterministic, driver-instrumented)

A counting global allocator in the standalone profiling driver (not part
of the repository) measured the source-open loop after warmup, 200
iterations, identical corpus in both builds:

| build | allocations/open | allocated bytes/open |
|---|---:|---:|
| control (pre-0206) | 14,891 | 1,974,783 |
| candidate (0206) | 10,727 | 1,928,711 |
| delta | **-4,164 (-27.96%)** | **-46,072 (-2.33%)** |

Counts are exact and identical across repeated runs and rebuilds; 4,164
matches the Start/Empty element count of the corpus `content.xml` (one
name allocation + copy + validation + free removed per element).

## Verdict

**Banked.** Claim scope: source-open allocation count -27.96%
(14,891 → 10,727 allocations per open) and allocated bytes -2.33%
(1,974,783 → 1,928,711) on the `ods-media-publication` corpus — a
deterministic count measurement, not subject to the layout noise floor.
Latency: neutral on every workload (all accepted statistics below the
calibrated floor; the single adverse both-directions reading, one-percent
commit p99 -12.18%, is within its 13.5% floor). The profiling-attributed
~4-6% latency expectation did not materialize above the floor — the
removed work is real but small relative to layout noise at this
magnitude. Raw artifacts: `docs/performance/results/*-0206-*`.
