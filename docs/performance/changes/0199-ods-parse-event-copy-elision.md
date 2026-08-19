# Change 0199: ODS parse loops drop per-event `into_owned` copies

Date: 2026-08-18

## Decision

Remove the per-event `Event::into_owned()` deep copies from the two
full-document `NsReader` parse loops on the ODS hot path. Accepted on all
five measured selectors; no regression pattern fired on any statistic.

## Mechanism and invariants

Two `NsReader` parse loops in `litchi-ods` deep-copied every XML event with
`Event::into_owned()` before matching on it:

- `worksheet::codec::parse_impl` — the full content.xml worksheet parse,
  reached at open (eager and source-backed) and again inside every
  source-backed commit readback (`codec::parse` after the row splice).
- `settings::codec::locate` — the full content.xml scan that locates the
  spreadsheet host and calculation-settings child during
  `SourceBackedSpreadsheet::from_package`.

`into_owned()` allocates and copies the event payload (start-tag bytes,
attributes, text) for every event — hundreds of thousands of copies for a
multi-megabyte content.xml — purely to detach the event from the reader
buffer's lifetime. Both loops compile without the copy: each event is only
read while the buffer borrow is live, and every value that outlives an
iteration (`Span::qname`, attribute values, cell text) is already copied into
its own allocation explicitly. Removing the calls changes no observable
behavior; event bytes, error paths, and error texts are untouched.

The `scan()` span builder in `worksheet/package.rs` intentionally keeps its
`into_owned()` calls: it stores element payloads in the retained
`ContentLayout`, so the copies there are load-bearing.

Profiling motivating the change (post-0198 source-backed open of the 16.8 MB
media corpus): quick_xml namespace/event machinery ~22%, `parse_impl` self
6.33%, `settings::codec::locate` self 4.55%, `memmove` ~9.3% across the
per-event copies and name decodes.

## Matched release timing

Two frozen release binaries differ only in the two `into_owned` removals;
both carry changes 0193-0196 and 0198 as baseline and the identical 341-case
selector matrix. Control SHA-256
`4e5662b270a30fc93af0b72bbe8c80d03adc82a73612a85085e6e3ac020954e1`
(bit-identical to the 0198 candidate, confirming build reproducibility),
candidate SHA-256
`6acae4dbfcac07185cba1047c5c157d52b096d0bdcc5642083cc6055e977c8ff`.
Fresh CPU-2-pinned processes ran `A1 control, B1 candidate, B2 candidate,
A2 control`, 30 warmups and 500 retained samples per leg. The predeclared
p50/mean/p95/p99 drift ceilings are 5%/5%/10%/15%; a statistic is accepted
only when both paired directions are lower and both drifts pass its ceiling.
Every leg ran clean: the open selectors verify the corpus hash and the
in-harness semantic/read-evidence gates (a failed gate aborts the leg), and
the edit selectors report all embedded verification flags true.

### Source-backed open (`ods_file_source_open`)

Elapsed p50/mean/p95/p99 all accepted at 6.42%-9.70% lower (drifts within
1.06%-3.90%).

### Eager open (`ods_file_eager_open`)

Elapsed p50/mean/p95/p99 all accepted at 9.46%-15.17% lower (drifts within
-4.54%-1.23%).

### One-edit guardrail (`ods_source_backed_one_edit_save`)

Lifecycle p50/mean/p95 accepted (0.97%-3.42% lower); commit p50/mean/p95
accepted (6.62%-9.43% lower). The p99 tails straddle zero and are withheld
as neutral.

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

Lifecycle p50/mean/p95 accepted (0.05%-2.71% lower); commit p50/mean/p95
accepted (4.16%-7.52% lower). The p99 tails straddle zero and are withheld
as neutral.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| total p50 | 2.33% | 2.14% | -0.82% | -0.62% | accept |
| total mean | 2.65% | 2.29% | -0.99% | -0.63% | accept |
| total p95 | 3.87% | 3.34% | -1.37% | -0.83% | accept |
| total p99 | 5.28% | 2.33% | -3.74% | -0.74% | accept |
| stage p50 | 0.86% | 2.37% | -0.36% | -1.88% | accept |
| stage mean | 1.15% | 1.94% | -0.66% | -1.45% | accept |
| stage p95 | 2.05% | 2.48% | -0.39% | -0.83% | accept |
| commit p50 | 12.46% | 10.99% | -2.71% | -1.07% | accept |
| commit mean | 12.74% | 10.93% | -2.90% | -0.88% | accept |
| commit p95 | 13.70% | 12.04% | -3.01% | -1.14% | accept |
| commit p99 | 14.80% | 8.61% | -4.82% | 2.09% | accept |
| publication p50 | 0.17% | 0.24% | -0.40% | -0.47% | accept |
| publication mean | 0.52% | 0.42% | -0.62% | -0.51% | accept |
| publication p95 | 2.20% | 1.92% | -1.39% | -1.11% | accept |
| publication p99 | 3.03% | 1.38% | -3.11% | -1.47% | accept |

Only stage p99 is withheld (candidate drift +16.26% against the 15% ceiling;
its paired directions straddle +6.73%/-11.41%). No statistic shows a
regression pattern.

The commit-phase wins come from the readback `codec::parse` inside every
source-backed commit; the open wins come from the worksheet parse plus the
settings host scan. The claim is scoped to the measured ODS selectors; other
format families' quick_xml loops are untouched by this change and are not
re-measured here. No allocation/RSS, physical-I/O, cold-cache, producer, or
broad ODF claim is made.

## Verification

```text
cargo test --locked -p litchi-ods --all-targets    # 342 passed, 0 failed
cargo clippy --locked -p litchi-ods --lib --all-features --no-deps -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --all-features --no-deps
cargo fmt --all -- --check
```

No manifest changed, so `cargo sort` is not implicated. The full-workspace
`--all-features` debug gate does not fit on this host's disk; the change is
private to two `litchi-ods` parse loops with no signature changes.

This change also corrects the derived-summary provenance blocks for changes
0196-0198: their `compiler` field carried a stale `rustc 1.97.1` template
string while every leg JSON records the pinned `rustc 1.95.0 (59807616e
2026-04-14)`. The summaries and their manifest hashes were regenerated in
place; no leg data or verdicts changed.
