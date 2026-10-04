# 0901 — current XLSX publication instruction attribution

The current one-percent XLSX publication profile identifies source-compatible
XML auditing as the largest guest-instruction branch: **58.57%** of the
publication owner for medium and **61.88%** for dense-sparse. Preservation
writing is the other substantial branch at **40.88%** and **37.69%**. These are
scoped Callgrind instruction counts, not native time shares or predicted gains.
No production change is adopted. iWork is excluded.

## Current-source scope

[0900](0900-xlsx-edit-save-phase-refresh.md) measured publication as the largest
native phase in every child, at 43.88–44.33% / 46.26–46.81% of its timed
edit/save interval. This batch uses the same retained 0894 release executable,
source-backed one-percent selector, and medium/dense-sparse corpus. All 10,771
selected source files still match the original build manifest; the executable,
build receipt, tool identities and prior native context are hash-bound.
Evaluation HEAD is `7371a3a094` and production is unchanged from `506aa2edae`.
The [packet](results/change-0901/README.md) retains the exact commands and data.

Four fresh processes run sequentially on CPU 12, in medium/dense-sparse then
reverse order. Each has zero warmups and one measured workload iteration.
The first planned child passes semantic and raw-profile admission before the
remaining three run. There are no retries or replaced observations. The shared
AMD EPYC host and in-memory warm provider do not establish physical cold-cache,
exclusive-host, concurrency, or cross-machine results.

Callgrind collection is toggled, reset, and dumped around
`SourceBackedEditor::publish_multi_commit_to_stream`. This scope excludes the
returned snapshot's destruction, which is inside 0900's native publication
clock. It also excludes opening, planning, staging, commit, and the untimed
output-reopen/oracle work. Instrumented elapsed durations are retained only as
raw observations and do not enter native conclusions.

The binary contains four specialized publication symbols with the same
demangled name. Address-bounded disassembly identifies the benchmark runner's
call target. Callgrind records a shared demangled owner ID; it does not prove
which machine-code instantiation ran. The unique positive incoming edge from
`run_xlsx_cell_values_edit_save` selects the measured dump. Each child also
contains five separately classified lifecycle publications and a zero-cost
termination dump. All twenty lifecycle calls are retained and excluded from
the four measured-owner totals. Five was the prior prediction, not authority
to silently discard a different count or unknown parent.

## Reconciled guest-instruction counts

| Shape / repeat | Publication owner Ir | Topology preparation Ir | XML audit child Ir | Preservation writer Ir |
| --- | ---: | ---: | ---: | ---: |
| Medium / 0 | 104,397,410 | 61,430,681 | 61,141,763 | 42,678,394 |
| Dense-sparse / 0 | 189,501,698 | 117,558,775 | 117,269,833 | 71,428,233 |
| Dense-sparse / 1 | 189,500,855 | 117,557,185 | 117,269,700 | 71,428,650 |
| Medium / 1 | 104,396,992 | 61,430,667 | 61,141,728 | 42,678,034 |

The XML audit is nested inside topology preparation; those columns must not
be added. Preparation and preservation writing are distinct immediate children
of `write_topology_to_stream`. Its other children and self cost remain explicit
in the machine-readable partition. The publication owner also includes small
plan construction and source-identity costs outside that topology call.

Preparation accounts for 58.84% / 62.04% of the measured owner. Almost all of
that branch is five calls to `validate_source_part_replacement_xml`, covering
four changed worksheets and the workbook. This delegates to
`xml_minifier::audit::verify_source_replacement`, preserving original-first
failure ordering and the complete source policy. The immediate writer branch
includes raw transfer, regenerated-member Deflate, framing, and accounting;
it is not a compression-only total. Nested Deflate and parser costs overlap
these branches and must not be added to them.

Exclusive function costs and a remainder independently reconcile to each
measured total. The leading self function is
`zlib_rs::deflate::longest_match::longest_match`; XML token reading, auditing,
attribute checks, delimiter searches, and memory operations also contribute.
A function's self count excludes its callees. Neither self nor inclusive Ir
is a latency bound, and no native phase percentage is multiplied by a guest
instruction percentage to predict a speedup. Two repeats per shape do not
establish a population confidence interval.

## Semantics and capture limitations

Every report matches 0900's corpus, output and semantic hashes, untouched-member
identity, source/cache counters, and sink observations. The source-backed
workload updates 93 / 178 cells across four sheets, retains twelve untouched
members, and emits writes no larger than 64 KiB. Logical source reads remain
4,233,170 / 4,258,604 bytes in 183 calls; these are not physical I/O or
recompression counts. No allocator or RSS result is inferred from Callgrind.

Every numbered dump has `events: Ir`, matching summary and totals, conserved
exclusive self costs, and one known positive publication-parent edge. The
measured edge has one call and its inclusive Ir equals the dump total. Each
owner's self plus direct children reconciles independently. The termination
dump is zero. Per-process stderr collection totals reconcile with the six
scoped dumps. Forked environment probes repeat the same collected total under
other PIDs; those repeated lines are not additional work and are not summed.

All four runs retain Valgrind's `brk segment overflow` warning. They exit
successfully and pass semantic and instruction-accounting checks, but the
warning remains a profiling limitation; no memory, native-time, or universal
cost claim is based on these runs. Runtime Git metadata is null because the
binary's original isolated build root was cleaned; runtime compiler metadata
is null because the fixed child PATH lacks rustc. External source/build/binary
receipts establish provenance.

## Bounded next hypothesis

The [source audit](results/change-0901/source-audit.md) identifies a small,
separate work-elimination hypothesis inside `verify_source_replacement`.
When the differing span already exceeds half the replacement, the function
marks window search stopped before scanning the original, yet still calls the
scanner specialized for a window observer. The source loop retains a dynamic
observer check; the unit-observer path represents a compile-time false check.
A candidate could dispatch that upfront-stopped case to ordinary source audit
of the original followed by ordinary source audit of the replacement, returning
`Complete` under the same limits and original/replacement error mapping.

This must preserve both mandatory audits, errors and offsets, failure ordering,
all budgets, and behavior at the exact half-size boundary. Stops reached later
in window search remain on the existing path. The profiles do not count this
specific branch or isolate its overhead; source redundancy alone does not
prove redundant machine instructions or qualifying end-to-end savings.
Correctness qualification and matched native measurements remain required,
with speculative changes reverted if they do not meet the benefit gate.

Changing compression policy is not justified by its prominent self function.
The 0837/0838 durable-inverse evidence requires preserving physical restoration
and source provenance, not merely equivalent decompressed bytes. Larger proof
reuse or parallelism would require separate ADR-compatible design and evidence.

## Verification and cleanup

The first child has independent raw admission. An optional annotation replay
failed on a source-path display format; its partial exports and failure receipt
are retained and excluded from the raw analysis. A separate root invocation
used a removed CLI flag and failed before creating output; the corrected
raw-only replay passes. Neither failure caused a workload retry.

The full analysis reuses the
retained 0845 raw reader, including its source-file-switch handling, and keeps
nested costs separate from disjoint partitions. Independent review, final
byte-identical replay, and root closure receipts bind the completed packet.
The helper author's first temporary analysis and root's duplicate replay are
removed. No production edits, builds, worktrees, or build roots were created;
unrelated workspace files and other teams' worktrees remain untouched.
The broader non-iWork goal remains active.
