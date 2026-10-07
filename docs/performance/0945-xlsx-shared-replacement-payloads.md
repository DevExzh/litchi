# 0945 — share owned XLSX replacement payloads

Adopt immutable sharing of detached XLSX replacement payloads. Publication
requests **9.75% fewer allocation bytes for medium** and **15.55% fewer for
dense-sparse** one-percent edits. Publication p50 is nearly unchanged.
Whole-workflow p50 increases **1.21–1.46%** across the four measured cases;
all frozen latency and memory guardrails pass. This is an allocation improvement,
not an end-to-end speedup claim. iWork is excluded.

## Change and preservation contract

`SourceTopologyPlan::try_replace_part_shared` accepts `Arc<Vec<u8>>` through
the ordinary OPC replacement path. The existing Vec route and new shared route
retain the same root-URI, operation-count, duplicate-name, and fallible-reserve
checks in the same order. Ordinary replacements still carry `source_xml: None`
and undergo the full original-first XML audit, source freshness checks, and
publication limits.

XLSX worksheet/workbook replacements share detached `Owned` payloads.
`Managed` and `OwnedBudgeted` payloads retain the fallible copy path and budget
lease behavior. Calculation-chain additions remain unchanged. Added tests cover
Arc identity, copy-on-write isolation, refusal order, output equality, managed
fallback and release, actual edit/publish/reopen, typed values, and exact
preservation of an unrelated styles part. This does not implement the XML
proof-reuse shortcut excluded by [0944](0944-xlsx-post-attribute-publication-profile.md).

## Experiment

The baseline is `e037befa6ce52b3988c8ae2032ae652196f2a09d`. Isolated Git exports
bind 10,788 source files per role; the final candidate changes only
`crates/litchi-opc/src/source_backed.rs` and
`crates/litchi-xlsx/src/cell_values/snapshot.rs`. The applied files match the
measured candidate byte for byte. Full source manifests and binaries remain
external; the [packet](results/change-0945/README.md) retains their descriptors.

The baseline native executable is reused from 0944 only after exact selected
source equivalence, including path, bytes, hash, Git object, and mode. The other
three executables were built offline with Rust 1.95.0, release optimization,
fat LTO, one codegen unit, debug level 1, incremental compilation disabled, and
two Cargo jobs. No compiler wrapper or RUSTFLAGS was supplied.

The primary case is `xlsx_source_backed_cell_values_one_percent_edit_save`,
medium and dense-sparse: 93/9,216 and 178/17,792 updated cells respectively.
Each generated package contains four worksheets, seventeen ZIP members, and
eight 512 KiB media members. One-edit medium workflows, unmanaged and managed,
are controls. This covers the [CRUD taxonomy](../CRUD_Scenario_Checklist.md)
update-and-save scenario on an in-memory source-backed package, not physical
cold-cache, remote-provider, or broad producer-corpus performance.

All processes ran serially on CPU 12 of the shared AMD EPYC 9R45 host. The
frozen plan has 16 qualification processes (zero warmups, two samples), 48
native processes (six paired blocks, ten warmups, 100 samples), and 16 allocator
processes (two paired repeats, two warmups, three samples). Odd blocks reverse
case and role order. All **80 processes / 4,880 measured samples** completed;
no benchmark process was retried or replaced. Qualification passed before
formal capture. Both formal stages finished before comparative gate review.

The primary allocation gate requires candidate/baseline publication requested
bytes at most 0.95 in **each** allocator pair. All four cases require workflow
and publication paired p50 median ratios at most 1.05, with two-sided 95%
bootstrap upper endpoints below 1.05 (10,000 draws, seed 945009). Allocation
calls, region peak live bytes, live bytes after publication, and process RSS
have 1.05 paired-median safety gates. Individual regressions, p95/p99 changes,
and spread triggers remain review flags. No allocator bootstrap is inferred
from two pairs. Publication phase vectors are independently sorted for their
quantiles; phase/workflow alignment is separately checked.

## Results

Ratios are candidate divided by baseline. Allocation counts below are the
per-child medians and are identical in both repeats.

| Case | Publication requested bytes, baseline → candidate | Reduction | Workflow p50 ratio [95% interval] | Publication p50 ratio [95% interval] |
| --- | ---: | ---: | ---: | ---: |
| One-percent medium | 2,765,098 → 2,495,613 | 9.75% | 1.01354 [1.01215, 1.01787] | 0.99968 [0.99659, 1.00249] |
| One-percent dense-sparse | 3,275,001 → 2,765,840 | 15.55% | 1.01371 [1.00129, 1.03039] | 0.99922 [0.97639, 1.02320] |
| One-edit medium | 1,065,966 → 1,002,065 | 5.99% | 1.01214 [1.01008, 1.01701] | 0.99484 [0.98504, 1.00108] |
| Managed one-edit medium | 1,066,046 → 1,002,145 | 5.99% | 1.01456 [1.01161, 1.01727] | 0.99899 [0.99499, 1.00424] |

Both primary rows reduce publication allocation calls from 416 to 406.
Region peak live bytes fall from 27,028,722 to 26,759,276 for medium and from
35,755,804 to 35,246,682 for dense-sparse. Live bytes after publication are
39 bytes higher in each row. These separate measurements must not be described
as a 9.75%/15.55% live-memory reduction. Process RSS paired ratios are 0.83979,
0.89734, 0.84886, and 0.85163 in table order, but two process pairs do not
establish a general RSS saving or physical-copy-volume claim.

Every allocation, latency, semantic, and safety gate passes. Both reduction
review-flag lists are empty; the smaller workflow regressions above remain
explicit. Reopen/output identities, untouched members and relationships,
source cache checks, sink writes bounded by 65,536 bytes, typed refusals,
managed release, and allocator conservation pass. The independent auditor
recomputed raw quantiles, paired statistics, gates, schedule, and artifact
custody for all 80 processes and agreed with the main reducer.

## Quality and retained failures

The corrected matrix passes **12,584 test executions**, summed over eight
baseline/candidate × OPC/XLSX × default/all-features test commands. This counts
repeated executions, not unique tests. Clippy, doctests, rustdoc, boundary
checks, and scoped formatting are separately recorded. One external ZIP64 source
test requiring `LITCHI_0415_PYTHON_ZIP` is ignored in each OPC test command;
that separate corpus was not supplied.
Package-wide OPC formatting still reports pre-existing, byte-identical drift
in `tests/source_fingerprint_quality.rs`; both edited files pass scoped rustfmt.

The first 38-check matrix retained 20 passes and 18 failures. Three missing
isolated test fixtures were restored from hash-bound sources. Candidate v1 had
two test-only errors: reuse after a consuming publication and an unnecessary
`drop` of a Copy view. Candidate v2 fixes those tests without changing the
production candidate. The corrected runner later failed while assembling its
summary after every child had completed; offline recovery used the retained
terminals without repeating those children.

The original measurement controller rejected its own absolute plan path before
launching a workload. Its empty first-child directory and frozen v1 plan/helpers
are preserved. Versioned v2 helpers correct internal descriptor paths, leaving
source, binaries, schedule, and thresholds unchanged; independent fake-process
controls exercise terminal and capture sealing. The first offline qualification
reducer invocation also used the wrong working directory and could not import
its semantic helper. The identical frozen reducer passed from the packet root,
without a workload rerun. These limitations are retained in the packet rather
than hidden by replacement records. Archive v1 also stopped before writing output
because it assumed the historical quality manifest was sorted; the v2 checker
accepts its recorded order while retaining duplicate, file-set, size, and hash checks.

## Evidence and scope

The full external reduction `analysis.json` is 5,439,340 bytes, SHA-256
`1246d3644ada2a12f6f3d2c580664e51fa35aea2e892d2cd87f960f1102db840`.
Regenerate it at the original packet location, from its root, with
`python3 -B tools/reduce_0945_v2.py all --output /path/to/new-analysis.json`.
The [retained packet](results/change-0945/README.md) provides compact results,
source/build and failure records, reproduction commands, and archive identity.
Every archived file passed extraction/hash verification. The independent raw
auditor reproduces the original results after relocation, and the compact
summary reproduces byte for byte. The frozen main reducer itself rejects moved
invocation paths; that failed replay is retained and is not presented as portable.

This adoption is scoped to sharing ordinary owned replacement payloads under
the existing audit and preservation contract. It does not change the dominant
XML/writer work identified by 0944, establish cold-cache or concurrent behavior,
or claim progress of an order of magnitude for the whole workflow.
