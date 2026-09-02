# Change 0374: DOCX story-hyperlink retained snapshot

## Scope

Change 0374 removes semantic recapture from the source-backed DOCX
story-hyperlink publication path. `ForwardOnlyPatch` retains the validated
story-hyperlink `Snapshot` produced during planning. Publication validates the
execution token, source version, source lineage, complete artifact
fingerprint, post-fingerprint version, and equality with the snapshot's stored
artifact fingerprint before writing; it then reuses the planned snapshot
rather than recapturing semantic story state.

The optimization remains inside the owning `litchi-docx` story-hyperlink
boundary and uses the existing OPC publication substrate. No public CRUD
signature, dependency edge, archive handle, runtime handle, lock, unsafe code,
or parallel execution path is introduced. Existing exact no-op output,
redaction member locality, source immutability, deterministic output,
freshness, signature, cancellation, and failure-atomicity gates remain in
force.

## Focused capture proof

The focused unit test `publication_reuses_the_planned_story_snapshot` passed
`1/1`. Its test-only counters reported `capture_source = 1` and
`load_story = 1` after planning and remained `capture_source = 1` and
`load_story = 1` after publication. This independently proves reuse of the
planned snapshot; it is not an allocation, RSS, or timing inference.

## Serial validation

Validation used one Cargo process at a time, `CARGO_BUILD_JOBS=1`, disabled
incremental and debug compilation, one test thread, an 8 GiB build virtual-
memory cap, and a 10 GiB `MemAvailable` admission gate. No parallel build or
worktree lane was used. The `litchi-docx` library suite passed `935/935`; all
other integration targets passed, including `source_backed_story_hyperlinks`
`23/23`. The exact pre-existing test
`replacing_the_path_reports_source_changed_without_retargeting` was excluded
after failing in isolation and is unrelated to this change.

Production-library Clippy with `-D warnings` passed with the named
pre-existing allowances `unnecessary_lazy_evaluations`, `double_must_use`,
`needless_borrow`, `redundant_closure_call`, and
`unfulfilled_lint_expectations`. The all-test run additionally exposed the
unrelated pre-existing `needless_question_mark` and `needless_lifetimes` debt.
The crate-boundary gate passed. The release build took 15m09s; 1.1 GiB of
validation targets and 938 MiB of release targets were reclaimed afterward,
leaving only the primary worktree.

## Clean ABBA evidence

The clean release ABBA used CPU affinity 2, one logical CPU, one execution
worker, 20 warmups, and 500 samples per case in `A1-B1-B2-A2` order. Control
was revision `73fabbcf707279e09d7a4de62a8190cb4075dd41` with tree
`bec7e7a2d2919d2f0bdfce4f50fd3a1cbfb82adc` and binary SHA-256
`1bf48f11ad9b600d0145dde0d30baee4bd05d25900cb77ff6721e7bcbba6f97e`.
Candidate was revision `7cb61e3938c040c01e4c48960d2ce881404cf63a` with tree
`7d22585dfe8c6c66ae16d36452106acfde2988c0` and binary SHA-256
`7b9d807392cb6171ef53b78f9acd0a6325ac3d352e0a588e47155cdc3db62600`.
Both implementations and all four legs were clean.

The corpus has seven story kinds, 15 OPC Parts, 24 ZIP members, 9,900 archive
bytes, and 39,564 uncompressed bytes. Its SHA-256 is
`457421e8f86ec8eb52fbe181cebe7d0821ce1e794a08142ff01a4c4e03df0cac`.
No-op output is byte-identical to the corpus. Redaction output is 9,757 bytes,
uses 125 writes, and has SHA-256
`7f776cb8087680f390c271bee7ff35da31e8ecf1f4469a6d63a89f3e9ded327c`.
All 14 correctness/refusal gates were true in all eight case-leg rows.

| Case and ABBA legs | p50 reduction | mean reduction | p95 reduction | p99 reduction |
| --- | ---: | ---: | ---: | ---: |
| No-op A1/B1 | 30.577956% | 29.506205% | 25.142442% | 20.667777% |
| No-op A2/B2 | 30.390954% | 33.137655% | 37.435278% | 58.413028% |
| Redaction A1/B1 | 16.592503% | 15.753115% | 15.442193% | 10.006419% |
| Redaction A2/B2 | 18.591303% | 18.832824% | 22.640355% | 24.036492% |

The complete compact evidence is [the machine-readable ABBA result](../results/docx-story-hyperlink-publication-0374-abba.json).

## Claim boundary and decision

The authorized claim is limited to the named benchmark cases on this exact
corpus and protocol: no-op p50 and mean, and redaction p50, mean, p95, and
p99. No-op p95 and p99 are withheld because control tail drift exceeded the
predeclared gate. Capture/load reuse is established separately by the focused
counter test and static review, not inferred from the timing cells.

No claim is made for reads, decompression, materialization, allocation volume,
RSS, physical I/O, cold-cache behavior, throughput, fixed memory, general OOM
prevention, all DOCX inputs, unmeasured selectors, or parallel execution.
The historical story-hyperlink planning and publication coverage records
remain historical and are not rewritten by this result.
