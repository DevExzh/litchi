# 0948 — multi-window XLSX audit rejected on allocation safety

The candidate reduces publication p50 by **14.25% / 14.63%** on the medium /
dense-sparse one-percent workloads, but increases publication allocation calls
by **67.73% / 129.31%** and requested bytes by **7.30% / 14.70%**. Both primary
cases fail the frozen allocation safety gates in every repeat. **The candidate
is rejected; production remains unchanged.** The unapplied patch and tests are
retained in the [evidence packet](results/change-0948/README.md).

## Mechanism and boundary

[0947](0947-xlsx-replacement-window-coverage.md) identified small changed cells
spread across worksheets that defeat the existing single-window proof. This
candidate carries exact source/output cell spans from the value-only writer,
through operation-local OPC replacement metadata, into the XML auditor.

The original XML is still completely audited first. Multiple hints must name
ordered, disjoint, in-bounds, balanced non-root elements. Every corresponding
outside gap must match exactly. The auditor establishes inherited namespace
and depth state from the original, validates each replacement element, then
subtracts removed counters and adds replacement counters with checked arithmetic.
Whole-replacement byte, depth, event, attribute, and text limits still apply.
Any ambiguous mapping or failed proof uses the canonical complete replacement
audit, preserving its error and offset. One hint uses the existing route;
exact no-ops retain the original audit behavior.

The maximum is 4,096 hints. Namespace snapshots have an 8 MiB conservative
logical admission charge, **not a physical RSS bound**. Reservations are
fallible; over-bound hints or state fall back. Replacement windows covering
more than half the output also fall back. XLSX metadata is admitted only for
multiple existing-cell updates with unchanged dimension/membership, checked
against source/output allocation identity and length, and cleared or rebound
when those bytes change. It does not retain an additional original payload.
Managed and owned-budgeted publication retains the existing fallible-copy route.
No parser checks are disabled, and no cross-operation parsed-XML cache or
stopped-observer shortcut is introduced.

| Accepted contract | Candidate obligation and review |
| --- | --- |
| ADR 0003 — snapshots, edits, patches | Immutable source identity, exact no-ops, publication freshness, and patch behavior remain; focused stale/repeated-edit tests and semantic comparison pass. |
| ADR 0005 — I/O, memory, performance | Bounded hints/state, fallible reservation, explicit ownership, and prospective measurement; allocation performance gates fail. |
| ADR 0006 — validation, security, compatibility | Complete original audit, namespace/depth/aggregate limits, exact outside gaps, and canonical full-audit errors; differential/adversarial XML checks pass. |
| ADR 0011 — physical OOXML package ownership | OPC owns replacement validation; XLSX supplies untrusted range metadata without taking over package validation. |
| ADR 0030 / 0031 — lazy OPC access and execution budgets | Existing fallible payload access and budgeted fallback remain; no ambient I/O or parallel execution is added. |
| ADR 0032 — small derived-value memos | No new memo admission is claimed; the metadata is an operation-specific proof hint and does not skip the original audit. |

## Prospective protocol

The frozen `plan-v3.json` binds baseline commit
`b6c5c80ffbcb60736d6c4c2f3b590838640008a9`, complete 10,789-file role censuses,
binary/build identities, fixed C/UTC environment, CPU 12, and four cases.
The host is x86-64 AMD EPYC 9R45 with 32 logical CPUs. Candidate builds use
Rust 1.95.0, release opt-level 3, fat LTO, one codegen unit, debug level 1,
incremental compilation disabled, and two build jobs. The post-capture host
receipt's default Rust 1.98.1 is not the measured build toolchain.
Baseline executables are reused from 0945 after exact source-equivalence checks.

All captures run serially: 16 qualification children × 2 samples, 48 native
children in six alternating paired blocks × 100 samples after 10 warmups,
and 24 allocator children in three alternating repeats × 5 samples after
2 warmups. All **88 children** pass capture/semantic custody. No child is
retried. These are named synthetic warm serial XLSX workflows, not cold-cache,
concurrent, native-producer-wide, or general Office performance claims.

For each primary, the median paired publication-p50 ratio must be at most
0.95 and the deterministic 10,000-draw bootstrap upper 95% endpoint below 1.0
(seed 948009). Every block and all four cases must meet the 1.05 p50/p95
publication and workflow safety bounds, with aggregate safety also checked.
Each of five allocator/RSS metrics must meet 1.05 both per repeat and by median
paired ratio. Allocation has no improvement target or bootstrap claim.

## Results

All ratios are candidate/baseline; smaller is better. Intervals describe paired
publication-p50 ratios. Absolute block quantiles, workflow intervals, p95/p99,
and every pair are retained in the machine-readable summary.

| Case | Publication p50 ratio | 95% interval | Workflow p50 ratio | Native verdict |
| --- | ---: | --- | ---: | --- |
| Medium 1% | 0.85747 | [0.85633, 0.86118] | 0.93839 | Pass |
| Dense-sparse 1% | 0.85366 | [0.81812, 0.90975] | 0.93110 | Pass |
| Medium one-edit | 1.00504 | [0.99679, 1.00839] | 0.99858 | Pass |
| Medium managed one-edit | 1.00000 | [0.99547, 1.00478] | 0.99970 | Pass |

The five allocator gate columns below are median paired ratios. Primary byte
and call failures are identical in all three repeats. All other gated pairs,
and all metrics in both controls, pass.

| Case | Requested bytes | Allocation calls | Region peak live bytes | Live bytes after | Maximum RSS |
| --- | ---: | ---: | ---: | ---: | ---: |
| Medium 1% | **1.07301 fail** | **1.67734 fail** | 1.00014 | 1.00014 | 0.99217 |
| Dense-sparse 1% | **1.14697 fail** | **2.29310 fail** | 1.00018 | 1.00019 | 0.98964 |
| Medium one-edit | 1.00014 | 1.00000 | 1.00001 | 1.00001 | 0.99103 |
| Medium managed one-edit | 1.00014 | 1.00000 | 1.00001 | 1.00001 | 0.99406 |

Medium publication requests 2,495,613 → 2,677,821 bytes and 406 → 681 allocation
calls. Dense-sparse requests 2,765,840 → 3,172,336 bytes and 406 → 931 calls.
These counters measure allocator requests, not physical copies. The small RSS
ratios do not override allocation failures. Additional non-gated deallocation
and reallocation flags remain disclosed in the summary.

Independent review recomputes the paired statistics and bootstrap from the
reducer's retained sample rows and agrees with the failed decision. It is a
gate-calculation audit, not a second independent raw-report parser. The full
reducer validates the raw reports, semantics, source census, and custody.

## Quality, corrections, and next step

The final quality composition is **26 raw passes and 4 documented baseline
failures**, not an all-green workspace claim. The failures are unchanged XML
all-target Clippy findings in `attribute_summary_probe.rs:266` and OPC formatting
drift in `source_fingerprint_quality.rs`. Candidate-focused tests, doctests,
Clippy, documentation, formatting where applicable, and boundary checks pass
through the recorded composition. Independent semantic review found no blocker.

Measured source v5 equals candidate-build-v3. Final source v7 changes only
`xml-minifier/tests/source_replacement.rs` for test lint/format fixes; all
472 retained release dependency records (236 per lane) exclude that test file.
The v7 focused default/all-feature tests and Clippy pass 4/4. The source bridge
supports reuse of the measured runtime bytes; it does not claim a v7 rebuild.
Earlier compilation, quality-fixture, and v6 lint failures remain retained.

Before capture, checks corrected draft scheduling and portable working-directory
bindings. After qualification, reducer v3 rejected equivalent source-summary
representations. Versioned reducer v4 independently rehashes both complete trees
and strictly normalizes full-stage versus focused child-end summaries. Its
correction receipt binds the original plan, both reducers, capture helper, and
unchanged gates. Independent negative controls pass. Existing qualification
children were reused; neither this correction nor the final rejection caused
a workload retry. Superseded helpers and failed attempts remain external.

Next attribute per-window namespace/range/parser allocation sites before
proposing another candidate. The counters establish a regression but do not
identify its dominant allocation site. Any subsequent optimization needs a new
prospective comparison; this batch's failed gates must remain failed. The
broader performance goal stays active, with iWork excluded.

## Retention and replay

The complete external archive contains 217,466 selected files and is bound by
size and SHA-256 in the [packet README](results/change-0948/README.md).
Every selected file's size, mode, and digest matches the extracted archive;
its three metadata files are checked separately. The full relocated reduction
is byte-identical to the original rejected result.

The unchanged reducer's absolute invocation-path check fails after relocation,
so an additional replay-only adapter binds the original authorization, plan,
and reducer hashes and reconstructs historical argv/cwd while reading relocated
artifacts. Independent actual-leaf controls reject altered cwd/argv and a
missing report. A file-syscall trace of relocated qualification finds no access
to the original packet root. Full replay is separately byte-identical; it was
not the traced invocation. Adapter, review, and cleanup evidence produced after
the main seal are explicit supplements rather than additions to that archive.

Owned target intermediates and replay scratch are cleaned after verification;
required historical executables, complete source versions, failures, and both
verified archives remain external. The committed subset satisfies the 1 MiB
per-file evidence limit and contains no binaries. No push is performed.
