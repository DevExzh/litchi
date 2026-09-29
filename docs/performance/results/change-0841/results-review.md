# 0841 results review

The 0841 packet passes its final analysis and remains diagnostic-only.  It
must not be used to adopt the rejected 0782 `Option<&str>` PPT candidate.
`analysis.json` reports `156` reports, `3,636` samples, zero native flags,
and zero spread flags.  The disposition is
`diagnostic-only-restore-production`.

## What this run establishes

The current-source paired native lane did not reproduce the old many/unicode
latency regressions.  The write p50 ratios (after candidate / before
production) are:

| shape | write p50 ratio (95% bootstrap interval) | lifecycle p50 ratio |
| --- | ---: | ---: |
| tiny | `1.02461` (`[1.01525, 1.03607]`) | `1.01140` (`[1.00042, 1.01942]`) |
| many | `0.99924` (`[0.98013, 1.01221]`) | `0.99608` (`[0.98730, 1.00421]`) |
| payload | `0.77742` (`[0.75444, 0.82742]`) | `0.81525` (`[0.78266, 0.85168]`) |
| unicode | `0.69301` (`[0.68476, 0.71272]`) | `0.70796` (`[0.70240, 0.71648]`) |
| rich | `0.99115` (`[0.96014, 0.99453]`) | `0.98152` (`[0.97817, 0.98687]`) |

The tiny write increase is about `+2.46%`, below the packet's `+5%` guard.
The many interval crosses one, and the payload and unicode reductions are
large in this current-source pair.  These are observations about the 0841
pair, not evidence that the candidate is safe: the 0841 source includes
post-0782 changes, the probe adds a shared `#[inline(never)] profile_write`
wrapper and barriers, and the before/after release binaries are new builds.
The wrapper is shared between the two legs, which supports pairing within this
packet, but it does not make the binaries or workload comparable with 0782.

The operation-scoped Callgrind lane gives the same localization signal in
guest instructions (Ir):

| shape | repeat 0 ratio | repeat 1 ratio | useful self-cost observation |
| --- | ---: | ---: | --- |
| many | `0.98642` | `0.98635` | plain textbox self Ir `199,000 -> 199,000`; shape-child `797,000 -> 798,000` |
| unicode | `0.94518` | `0.94518` | plain textbox self Ir `49,778,569 -> 49,778,569`; `memcpy` `3,603,876 -> 634,572` |
| payload | `0.62490` | `0.62489` | plain textbox self Ir `531,712 -> 531,712`; `memcpy` `5,450,619 -> 2,890,159` |
| rich | `0.99862` | `0.99859` | formatted textbox self Ir `12,328 -> 12,328`; shape-child `3,232 -> 3,228` |

The `UserShapeData` drop and allocator rows also decrease in the unicode and
payload profiles.  That identifies the conversion/copy/drop neighborhood as
the next attribution target, while the unchanged plain-textbox self row and
nearly unchanged shape-child rows do not support claiming that the text-writing
loop itself caused the difference.  Callgrind Ir is guest instruction
accounting; it is not native time, physical copy bytes, or a causal allocation
attribution.
Missing separately emitted names can have been inlined or folded into their
callers.

## Why the old rejection remains

Frozen 0782 evidence recorded guard failures for `many/write` (`+12.897%`),
`many/lifecycle` (`+11.089%`), `unicode/write` (`+20.650%`), and
`unicode/lifecycle` (`+19.069%`).  The 0782 payload cases improved, but its
operation allocation evidence confirmed the expected removal, and Heaptrack
attributed 192 removed allocation events / 7,680,000 requested bytes to the
conversion path across three payload writes. That evidence and whole-process
`perf stat` did not explain the latency regressions. The 0841 native lane does
not reproduce those old failures, but it cannot invalidate them: current
source, build configuration and the `profile_write` wrapper differ from 0782.  Specifically,
0841 uses current HEAD `307434b454f622023818f86f724b56ae5c99ee15`, whereas the
0782 production base was `f1df64d15a8f2711e554f58acb413b06d56ef1ca`; 0841's
release builds use Rust/Cargo `1.95.0`, LLVM `22.1.2`, and LTO.  Its retained
before/after probe binaries have different SHA-256 values
(`0cc6b3cd12ab09cdb088d8d4a9ecfcb55a1e092527d229a2108438f6224585d9` and
`7a057a936c6777c23cc2295f6f033ba8ba0ebe20c42147a7652391522bcfd6bf`).  The
current result therefore supports investigation only; it does not support
another representation change or production adoption.

The reader and analyzer repairs are additive history.  `reader-repair.json`
reports native parity with all 20 qualification rows unchanged, and
`analysis-repair.json` reports unchanged metrics and thresholds.  Neither
repair replaced a workload run (`workload_runs_replaced: 0`).  The repaired
analysis and the source restoration receipt are part of the result record;
the production source is restored after capture.

## Next diagnostic action

Use the already admitted 0841 binaries for a narrow symbol and assembly
attribution pass before considering any new candidate.  Query or annotate
these boundaries separately, treating absent symbols as inlining/unknown:

* `convert_shape_to_escher_with_sound_mapping` and the generated
  `UserShapeData` move/drop path;
* callers of `__memcpy_avx_unaligned_erms` in the selected write owner;
* `append_shape_container_children` and `build_shape_properties`;
* plain and formatted client-textbox builders and their inline call stacks.

`nm -C`, `objdump -d -C`, DWARF inline information, and address-to-source
mapping from the same before/after binaries can compare prologues, spills,
loads/stores, branches, and clone/allocation/drop calls without changing the
diagnostic executable.  Do not add standalone struct-layout instrumentation:
that would change the measured binary.  If this attribution finds a concrete
boundary, a later candidate still needs a fresh operation-scoped native run on
current source and build, exact output/readback checks, and the 0782 guard
policy.  No new workload run is justified by the present Ir-only signal.

## Highest-impact compatible OOXML lane

The next production work-elimination opportunity is opened-package
source-backed save, rather than another PPT representation trial.  In
`crates/litchi-opc/src/source_backed.rs:11721-11935`,
`write_changed_overlays_with_appended_inner` currently:

1. allocates scratch and rebuilds a `PreservationIndex`;
2. allocates and sorts replacement/omission lookup vectors;
3. scans the index to count replacement and omission entries; and
4. allocates a `PreservationPlan` and scans the same entries again before raw
   copy/regenerated-member emission.

The concrete next profile should use one opened source-backed OOXML package
with one changed Part and measure those phases separately, including scratch,
lookup, index, and plan allocation, followed by the actual emission.  Only a
measured dominant phase should lead to a candidate, such as bounded reuse of
immutable central metadata.  Exact no-op raw-source passthrough already exists
and is not a new target.  DOC's remaining first-touch/page-fault path also
needs measurement before proposing a buffer or spare-capacity rewrite; the
current DOC write path already moves final stream buffers into CFB and emits
large slices directly.

Any opened-package candidate remains bounded by:

* ADR 0005's I/O and memory budgets;
* ADR 0006's exact-preservation and typed-refusal rules;
* ADR 0030's source freshness and lazy-decode contract; and
* ADR 0032's prohibition on an unbounded parsed-index memo.

This review records the measured disposition and next diagnostic step.  It
does not authorize source adoption or claim a causal explanation from the
current profiles.
