# 0951 — XLSX publication after transient multi-window auditing

Preservation writing is now the largest named publication branch: **57.70–60.95%**
of Callgrind guest instructions on the measured dense-sparse/medium cases.
The nested source/replacement XML audit accounts for **38.61–42.05%**.
Deflate longest-match search is the largest exclusive function. These are
current instruction shares, not native-time shares, speedups, or allocation
measurements. No production change is made; iWork is excluded.

## Current source and measurement scope

The adopted source is `ce6f1fbea9902d533b97f6830994d4bc3fd497a1`.
The retained 0950 native executable is 266,090,728 bytes, SHA-256
`198c2e87471a088cf054a23785e4068eaf7f205ffb1fb25cad5ba0b7a6db5bec`.
It was built from the frozen 0950 candidate based on `b6c5c80f...` plus its
12-file delta, before adoption; this is not a claim that it was compiled at
current HEAD. All 10,789 tracked selected files match current committed Git
objects. The additional ignored root `Cargo.lock` matches the frozen manifest
and current checkout by hash. The actual standalone build used its tracked
`tools/perf-baseline/Cargo.lock` with `--locked --offline`.

Reusing this verified binary avoids a rebuild or a source change. Its build
uses Rust 1.95.0, release optimization, fat LTO, one codegen unit, debug level 1,
and incremental compilation disabled. Source/build/binary checks bind the
retained 0950 archive. Unrelated working-tree changes are excluded from the
comparison. Goal, scenario checklist, and accepted ADR hashes are unchanged.

Four fresh processes ran serially on CPU 12: medium, dense-sparse, dense-sparse,
medium. Valgrind 3.26.0 collects `Ir` with zero warmups and one sample per child,
a 900-second process-group deadline, fixed C/UTC environment, and no retries.
The first child passed semantic and raw-profile admission before the others.
The case is `xlsx_source_backed_cell_values_one_percent_edit_save`: 93/178
scalar updates on deterministic medium/dense-sparse four-sheet XLSX packages.
Opening, planning, commit, and reopen checks lie outside the selected
`SourceBackedEditor::publish_multi_commit_to_stream` owner. Destruction of the
returned snapshot is also outside it, although the native publication timer
includes that destruction. This is an in-memory, source-backed edit/save
scenario; no cold-cache or remote-provider behavior is measured.

## Reconciled attribution

| Shape / schedule index | Publication Ir | Preparation Ir | Nested XML audit Ir | Writer Ir | Other publication Ir |
| --- | ---: | ---: | ---: | ---: | ---: |
| medium / 0 | 70,017,755 | 27,323,146 | 27,035,195 (38.61%) | 42,678,061 (60.95%) | 16,548 |
| dense-sparse / 1 | 123,788,788 | 52,340,914 | 52,052,412 (42.05%) | 71,431,254 (57.70%) | 16,620 |
| dense-sparse / 2 | 123,788,193 | 52,341,011 | 52,052,790 (42.05%) | 71,431,112 (57.70%) | 16,070 |
| medium / 3 | 70,018,237 | 27,323,847 | 27,035,765 (38.61%) | 42,677,646 (60.95%) | 16,744 |

Preparation and writing are disjoint immediate children of
`SourceBackedPackage::write_topology_to_stream`. XML auditing is nested inside
preparation and must not be added to it. The residual subtracts only the two
immediate branches from publication; it retains owner self work, plan creation,
other calls, and drops. Every measured dump's exclusive self rows reconcile to
its complete instruction total. Two repeats per shape do not establish a
confidence interval.

The current XML owner is `audit_source_part_replacement_xml`, rather than the
older `validate_source_part_replacement_xml` wrapper. Each measured dump has
five incoming audit calls across two callsites. Their summed edge cost equals
the audit function's self cost plus direct children. The initial path diagnostic
retained one representative route, which is insufficient for total attribution:
in medium repeat 0 it showed only the one-call 65,464-Ir edge, while the other
four calls contribute 26,969,731 Ir. The supplementary reducer retains both
callsites and reconciles the complete 27,035,195 Ir. The initial diagnostic is
preserved, not used as the total, and no capture is rerun.

`MultiWindowSearch::finish`, `audit_replacement`, and `prove` are observed on
the current path. Their costs overlap the audit owner and are retained as
nested diagnostics. The complete original scan and provisional replacement
windows share this owner; this profile does not separate an original-only
cost. Original errors remain authoritative even when window work interleaves.

The largest exclusive row, `zlib_rs::deflate::longest_match::longest_match`,
uses 18,584,283 Ir in both medium repeats and 35,695,609 Ir in both dense-sparse
repeats. This row is nested inside the preservation writer; it cannot be added
to that inclusive branch. Writer cost includes changed-member compression,
unchanged-member transfer, framing, and accounting.

## Semantics, review, and limits

Every child has one measured publication, five separately retained lifecycle
publications, and an empty termination dump. Numbered dumps, owner parents,
terminal status, sample count, chronology, artifact hashes, raw header commands,
and collected-instruction totals are checked. The terminal receipt precedes
report consumption. Output reopen, updated-cell semantics, topology and
relationships, source/cache behavior, sink bounds, and twelve untouched members
pass against retained 0950 context. Maximum writes remain 65,536 bytes.

All four runs retain Valgrind's `brk segment overflow` warning. Successful
exit, semantic admission, and instruction conservation do not erase that
profiling limitation. Instrumented elapsed times are retained but excluded
from conclusions. The 0950 native control-p99 observation remains separate;
this profile neither explains nor requalifies it.

The prior 0944 profile predates both shared replacement payloads (0945) and
transient multi-window auditing (0950). It is historical context, not a matched
before/after experiment. Do not multiply these instruction shares by historical
native phase shares or turn them into a predicted speedup.

## Next work boundary

Investigate the changed-member compression path and its longest-match work
before another XML micro-optimization. Any candidate needs separate current
allocation and native qualification, including both one-edit controls and tail
review. Preserve compression/container contracts, unchanged-member passthrough,
source checks, exact no-ops, reversible patches, resource budgets, typed
refusals, complete original auditing, and canonical fallback. The older
rejected shortcuts remain rejected. This evidence ranks work; it does not
justify a compression-policy change or introduce a candidate implementation.

## Reproduction and retained evidence

Independent four-child reconciliation passed. The deterministic reduction is
`analysis/reduction-0951-v1.json` (112,004 bytes, SHA-256
`395bdaa7167bdb120e0850ef1c171168ac9439c013a58987a478a0abe6f9cb31`).
The reducer is `tools/reduce_0951_v1.py` (29,192 bytes, SHA-256
`c4074267ec9f320671a34fc2ba31461173a68f1d1a3891e1f9b851d974cf6bc2`).
Run `python3 -B tools/reduce_0951_v1.py /tmp/reduction-0951.json` from the
retained 0951 packet; the 0944 parser and 0950 source/build/binary dependencies
must be present. Captured absolute commands are historical metadata.

The raw adapter's `lifecycle_prediction` is null because it reads a top-level
key while the frozen plan stores the expectation under `callgrind`. The
observed five lifecycle calls are independently checked; no capture changes
or retries were needed. Archive identity, relocated replay, negative controls,
and cleanup are recorded in the accompanying final archive receipt outside
this packet, avoiding a self-referential archive hash.


The external archive `archive-0951-v1.tar.zst` is 459,597 bytes, SHA-256
`66f5668e4d8299b1d21519bb83b7d6a2c04bb17e631c1b2cd047628fbc1c2aef`,
under `/home/zhuhe/litchi-performance-0951-archive/`. Its 28,615-byte manifest
has SHA-256 `aacca68cc6d39fe0df017c3e00c50e4eb12f5220e01a34a2d05c6267a9cdc6f8`:
126 files and 4,923,164 logical bytes. Raw captures and profiler traces remain
external. The archive references the retained 0944 parser and 0950 source/build
archives; it does not duplicate those dependencies or claim to be self-contained.

The [compact summary](results/change-0951/summary.json),
[independent review](results/change-0951/independent-review.json), and
[source/binary receipt](results/change-0951/source-binary-reuse.json) are
committed. The copied [reducer](results/change-0951/reduce.py) is an exact source
reference; execute it in the archived `tools/reduce_0951_v1.py` layout.


[Final archive/replay closure](results/change-0951/archive-closure.json) records
exact reproduction commands, hashes, extraction checks, and cleanup. All 126
files passed hash/mode verification. Relocated replay reproduced the summary
byte for byte, with zero original-0951 reads observed in file tracing. The
frozen reducer needed an external path adapter for historical descriptors;
the original failure and three intermediate adapter failures are retained.
Use the retained `replay_0951_v5.py` and its hash-bound v4 parent at the archive
paths recorded in the closure. Their committed copies are source references.

The initial v4 negative controls only asserted mutations differed and were
rejected in root review. Corrected v5 controls invoke the actual command and
manifest validators: altered argv/cwd/artifact paths, missing and extra files,
symlinks, and manifest path escape all reject. The old receipt remains
historical evidence, not accepted validation. Owned extraction and bytecode
scratch were removed after successful replay; canonical evidence and required
0944/0950 dependencies remain. The archived draft's reducer byte count was
29,159; the verified file is 29,192 bytes with the same stated hash.
