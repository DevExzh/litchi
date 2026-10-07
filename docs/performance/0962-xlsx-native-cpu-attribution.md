# 0962 — XLSX native CPU attribution

Status: diagnostic frame-pointer attribution qualifies. Production is unchanged;
no optimization, speedup or new ordinary-native baseline is claimed.

## Motivation and custody

The 0959–0961 candidates reduced particular allocation costs without a
qualifying end-to-end benefit. Requested allocation bytes are not a CPU or
live-memory ranking. This batch investigates native execution within the
managed benchmark's edit/commit phase before selecting another candidate.

The previous turn made progress: 0961's complete screen was independently
verified, rejected and committed as `dfb777544c`; disposable caches were removed.
All 10,932 files in the current committed source export (excluding performance
and report evidence) match the retained 0957 source in content and executable
bit. All 35 normative inputs are unchanged. The retained uninstrumented native
benchmark is SHA-256
`5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`,
43,259,328 bytes. The external evidence root is
`/home/zhuhe/litchi-performance-0962/`.

## Unprofiled phase bound

The existing 0961 **baseline-only** native packet supplies six processes and
360 measured samples per case. The summary computes each phase's sum divided
by the total-time sum per process, then reports the median process fraction.
Independent arithmetic review reproduced every result. Candidate rows are
excluded. These are existing native timing observations, not profiled timings.

| Case | Edit/commit fraction | Ideal total speedup if that phase vanished |
| --- | ---: | ---: |
| Medium 1% | 61.39% | 2.590× |
| Dense-sparse 1% | 58.68% | 2.420× |
| Medium one edit | 57.21% | 2.337× |

The Amdahl bound is `1 / (1 - f)`, assuming every other phase stays unchanged
and no overlap. It is not an observed speedup or a prediction that the entire
edit/commit phase can be removed. It rules out attributing a 10× end-to-end
improvement in these cases to edit/commit acceleration alone.

## Phase attribution contract

`timed_iteration` times constructor ingress, selector staging and
`stage_timed_commit`, publication including returned-snapshot drop, and
commit teardown. Corpus construction, preflight, warmups, per-iteration setup,
output verification, semantic readback and inverse fixture checks also run in
the process. Whole-process CPU samples therefore cannot be labeled edit/commit
samples. Exact ancestor and inline-callsite qualification, unresolved coverage,
event-period weighting and nonoverlapping attribution are required. Profiled
wall times and RSS are diagnostics, not baseline or speedup measurements.

The first 8 KiB DWARF capture proved unsuitable: the timed-iteration function
alone has an approximately 10.9 KiB stack frame, and almost no sampled stack
retains its phase ancestors. Increasing frequency without changing stack size
does not solve that problem. The 65,528-byte follow-up also fails: only 3 of 21,487 medium samples
retain both phase markers, and roughly two thirds contain unresolved frames.
This is a reconstruction problem; those samples cannot rank phase CPU owners.
The original schedules and their timestamp correction remain external. Their
manually authored creation times were inaccurate and are not freeze evidence.
Failed qualification remains evidence, not an optimization result.

A separate frame-pointer/debug-level-two build uses the verified retained
source, Rust 1.95.0, locked/offline resolution, opt-level 3, fat LTO, one codegen
unit and two Cargo jobs. Source verification passes before and after the build.
Its retained executable is 104,819,848 bytes, SHA-256
`843ae35a0f5131a8cba965f917b2a4c5fccb6926ae3d66862e5f78512820fb33`.
The changed code generation makes this an attribution-only build. Earlier
unqualified build attempts are excluded, including an attempt from the dirty
main checkout. No executable from those attempts was profiled.

The corrected pilot plan is frozen before launch. It uses 997 Hz user-cycle
sampling with frame-pointer callchains, zero warmups, 200 measured iterations,
one worker and CPU affinity 12–15. All inherited threads remain in the raw
population; ancestry qualification is thread-local. The six-child follow-up
wave requires passing semantics, matching ELF mapping/build ID, paired physical
and inline sample identity, zero lost records, complete sample/period
conservation, and at least 100 resolved phase samples / 100,000 period weights
in the pilot. Each follow-up child must also pass those gates.

## Decoder corrections

The pilot has paired physical and inline exports. Early reductions remain
provisional: one mislabeled zero-stage rows as multiple-stage rows; another
failed closed because it ignored the inline marker on source-location lines.
Neither error changes the raw capture. Review also rejected heuristic stack
tags that matched type-name substrings and required source locations to
disambiguate short inline names. Follow-up exports request full source paths.
Only the final source-qualified reduction may support an owner ranking.

The final pilot reduction is `analyzer-paired-v5.json`; independent admission
`review/fp-pilot-admission-v2.json` passes all 28 checks. Both streams contain
20,998 samples and 91,320,259,987 period weights, with zero lost, unmatched or
duplicate records. Exact physical timed ancestry plus the inline stage selects
1,921 samples / 8,443,033,125 period weights. The 1,193 timed-without-stage rows
and 16 unresolved-stack rows remain excluded. Pilot owner ranking is withheld;
162 selected rows lack a deepest source location.

The prospective `fp-wave-execution-v1.json` addendum binds the frozen helper,
pilot admission and original plan. Six captures run serially with semantic and
terminal validation before each next launch. Offline decoders run only after
the capture lane ends. Any failed attribution gate rejects the entire wave;
no child is replaced. All six captures completed with 1,200 passing semantic
iterations. The independent review also confirms exact command/helper/binary
bindings and clean process exits. Offline owner admission is separate.

After the wave was launched, the reviewer corrected a display-only workload
PID field in its admission receipt. The post-freeze correction was preserved
separately, and the original bytes were restored exactly against the plan
hash. `fp-pilot-freeze-restoration-v1.json` records that event; no gate,
selection, command or workload changed and no capture was repeated.

## Qualified frame-pointer wave

All six captures pass exact paired-stream identity, period/sample conservation,
zero lost events and the per-child 100-sample / 100,000-period minimum.
Independent CPU-header reduction reproduces all 134,969 events and
583,643,582,118 period weights. Qualified stage ancestry contains 12,319 events
and 52,852,646,778 period weights. Every child sampled only its observed main
PID/TID, on one CPU within the allowed 12–15 set. Thread-local qualification
remains the contract; this does not prove complete phase coverage.

| Child | Case | Qualified samples | Qualified period weights |
| --- | --- | ---: | ---: |
| 00 | Medium 1% | 1,899 | 8,347,974,970 |
| 01 | Dense-sparse 1% | 3,722 | 15,898,003,220 |
| 02 | Medium one edit | 499 | 2,195,986,095 |
| 03 | Medium one edit | 470 | 2,068,701,002 |
| 04 | Dense-sparse 1% | 3,747 | 15,992,963,474 |
| 05 | Medium 1% | 1,982 | 8,349,018,017 |

The following full-symbol groups are **exclusive physical self including
inlined callees**, within the qualified denominator. Out-of-line descendant
costs are excluded. The ranges describe the two observed captures, not
confidence intervals or a production CPU/time fraction.

| Physical function | Medium 1% | Dense-sparse 1% | Medium one edit |
| --- | ---: | ---: | ---: |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 7.16–9.01% | 7.16–8.49% | 5.32–10.61% |
| `quick_xml::name::NamespaceResolver::resolve_event` | 7.06–7.11% | 6.86–7.87% | 7.82–9.79% |
| `litchi_xlsx::cell_values::validation::worksheet_xml_and_parse_source` | 6.88–7.27% | 5.99–6.64% | 6.17–7.01% |
| `litchi_xlsx::cell_values::validation::scan_attributes` | 4.42–8.83% | 3.69–6.94% | 7.22–7.22% |

Full source paths do not resolve every inline owner: some DWARF locations are
still relative, and others absent. Under the stricter absolute-path rule,
55.13–64.09% of qualified period remains source-unresolved. It stays in every
denominator. The separate 90.84–93.19% “any source location” coverage includes
relative paths and must not be presented as unique source-owner coverage.
No deepest-source-file ranking or heuristic inclusive-tag percentages are
adopted. Physical function groups retain full symbol and exact executable DSO.

The initial offline controller observed transient nonempty process groups
immediately after successful inline decoder exits. Recovery v2 verifies later
clearance and stable export hashes, preserves original terminal records, and
reuses the completed inline exports. It then completes missing physical
exports and parsing with bounded process-group cleanup. No capture or inline
export was replaced. Separate metadata corrections preserve their older
receipts; CPU-summary v2 and mapping-metadata v3 are authoritative.

## Source admission and rewritten readback

A separate replay uses the same qualified samples and exact full physical
symbols. `Snapshot::from_source_selected`,
`Snapshot::from_rewritten_value_source`, neither, and ambiguous are disjoint
buckets; their counts and period weights conserve every selected denominator.
There are zero ambiguous rows. `validate_xml` beneath the rewritten snapshot
is a **nested subset**, not a fourth additive bucket.

| Physical ancestry | Medium 1% | Dense-sparse 1% | Medium one edit |
| --- | ---: | ---: | ---: |
| Original-source admission | 61.63–62.00% | 60.83–61.27% | 60.74–61.52% |
| Rewritten-value readback | 35.01–35.95% | 35.98–36.03% | 34.01–36.25% |
| Nested output XML validation | 28.35–28.69% | 29.01–29.28% | 29.76–31.05% |

These are sampled period shares of qualified stage ancestry in the diagnostic
build. The first two rows are disjoint; the third is contained in the second.
Pooling the six captures gives 61.29% original-source admission, 35.77%
rewritten readback and 2.94% neither. The nested output validator contributes
29.05% of pooled qualified period. Pooling gives denser cases more weight;
the per-case repeat ranges above remain the primary description.

Original-source admission already shares its validator, facts observer and
semantic parser in one XML-reader traversal. The rewritten path performs a
separate full output XML validation before semantic parsing or reduced
readback. That is a concrete next investigation, not permission to remove an
authoritative check. Any fusion or reuse must preserve namespace/MCE behavior,
malformed-input error ordering, source freshness, omission provenance, staged
semantic readback, cancellation, budgets, deterministic output and inverse
restoration. Native latency/tails, allocator traffic/live peak and RSS still
require a fresh matched candidate screen. Prior regression flags remain open.

Fusion is not a drop-in solution for reduced readback: that path removes
omitted-cell byte ranges before parsing and later merges source records. A
full-output reader would materialize omitted cells again unless a new
omission-aware parser contract preserved complete-cell/row ranges, error order
and MCE offset semantics. First census the actual omission routes for the
measured cases. A bounded no-omission fusion experiment must not be presumed
to benefit sparse reduced-readback cases. The source review records both this
constraint and the narrower proof hypothesis without adopting either.

## Evidence and cleanup

The bounded [evidence packet](results/change-0962/EVIDENCE.md) contains helpers,
plans, terminal records, independent reviews and summaries. Its
[external inventory](results/change-0962/external-artifact-inventory.json)
binds every retained artifact by filename, bytes and SHA-256. Raw captures,
large decoded streams and the compiled executable remain outside Git. Build
and capture receipts record their regeneration commands and dependencies.
The early DWARF command logs use `benchmark`/`binary` as placeholders for
`/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2`;
their schedules bind its exact size and hash. Corrected FP start records
contain the complete executable path and argv.

Validation covers the verified diagnostic build, 200 pilot plus 1,200 follow-up
semantic iterations, 11 analyzer controls, capture-validator corruption
controls, independent mapping/CPU totals, and exact selected-population and
partition conservation. Production Rust was not changed, so no new library
test suite or adoption qualification is claimed. Disposable build caches
covering 5,994,821,702 logical bytes were removed after binary verification;
legacy `/tmp` diagnostics were retained externally. Canonical evidence remains.

## ADR boundaries

| Constraint | This batch |
| --- | --- |
| 0001 / 0005 | Measure actual execution; no claim from requested-byte traffic alone. |
| 0002 / 0024 | External tooling only; no production dependency or ownership changes. |
| 0003 / 0006 | Existing benchmark semantic, preservation and inverse checks remain. |
| 0008 | Profile custody, phase qualification and independent reduction precede conclusions. |
| 0031 | Explicit one-worker context and CPU affinity; no new ambient executor. |

No iWork work or unrelated worktree edits are included. The broader GOAL
remains active.
