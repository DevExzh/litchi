# Change 0966: bounded DOC SPRM census

Status: the sealed diagnostic build and authorized 48-child capture passed the
independent reduction and checker. This is a diagnostic result only. It is not
a timing, allocation, causality, or production-adoption decision.

This packet is read-only evidence for the current DOC lifecycle. It asks how
often the existing `parse_sprms_two_byte` path sees empty and short SPRM input
and what logical returned-vector capacity it reports. It does not change parser
semantics, choose a capacity policy, or authorize a production edit. The
instrumentation is gated by the existing `performance-diagnostics` feature and
is retained in the external sealed source tree only.

## Fixed protocol

The matrix has two fixtures (`NoHeadFoot.doc` and `FloatingPictures.doc`), four
edit counts (1, 2, 8, and 32), and three retention routes (`zero`,
`default_8mib`, and `release_8mib`). It has 24 cells, two fresh serial child
processes per cell, and therefore 48 children. Each child uses `warmups=0` and
`samples=1`; there is no retry or replacement process. The diagnostic is a call
census, so `timing_claim=false` is required. The five lifecycle phase buckets
are `snapshot_open`, `edit_open`, `staged_replacement`, `final_commit`, and
`output_copy`, with an `other` bucket.

The observer records fixed bounded counters and histograms only: calls,
successes, errors, empty/nonempty input, input bytes, successful output count,
logical returned-vector capacity, and a fixed error-kind histogram. The
capacity-byte field is `successful_output_capacity_total * size_of::<Sprm>()`;
it is a modeled logical value and does not measure allocator traffic, RSS,
peak live memory, or removable work. Existing output, semantic, retention,
atomic-refusal, and preservation gates remain required for every child. The
independent checker and reducer must reject missing, extra, failed, saturated,
or non-conserving child reports and must require byte-identical fresh repeats.

## Captured result

The external run completed 48 of 48 children with zero failures and zero
retries. The reducer accepted all 24 cells, exact forward/reverse repeat
identity, counter conservation, and all existing correctness gates. Its
representative-per-cell totals were 119,775 parser calls and 119,775
successes, with zero errors; 5,703 calls had empty input, and the modeled
logical returned-vector capacity totaled 49,424,112 bytes. Those totals are
descriptive counters from the 24 representative cells, not allocator bytes or
saved work.

The per-cell call-count trend was identical across the three retention routes
within each fixture/edit cell:

| Fixture | 1 edit | 2 edits | 8 edits | 32 edits |
| --- | ---: | ---: | ---: | ---: |
| `NoHeadFoot.doc` | 170 | 187 | 310 | 1,162 |
| `FloatingPictures.doc` | 4,742 | 5,220 | 8,109 | 20,025 |

Each table value is one representative cell; each has two fresh identical
children and each retention route is represented separately in the reduction.
The phase breakdown was:

| Phase | Calls | Empty successful calls | Logical capacity bytes |
| --- | ---: | ---: | ---: |
| `snapshot_open` | 20,640 | 1,092 | 8,712,000 |
| `edit_open` | 11,736 | 552 | 4,798,080 |
| `staged_replacement` | 66,231 | 2,967 | 26,999,280 |
| `final_commit` | 21,168 | 1,092 | 8,914,752 |
| `output_copy` | 0 | 0 | 0 |
| `other` | 0 | 0 | 0 |

The phase table is also one representative per repeat-equal cell. The
retention equality and edit-count relationship are route observations only;
they do not establish a faster route, removable parser work, or a capacity
policy.

The reduced 24-cell result is
[`analysis/census-summary-v1.json`](analysis/census-summary-v1.json),
11,652 bytes, SHA-256
`53339c9a9c9e22af0f1773fdd2621351948ef0b207263ac3744a04f63d4121d5`.
The per-cell details are
[`analysis/census-cell-details-v1.json`](analysis/census-cell-details-v1.json),
83,816 bytes, SHA-256
`6b45e9a2ddce1f5edde696619dea9edb8dcd7be772f129ef25179cfef9d56c1b`.
The 48-child terminal/report custody manifest is
[`analysis/child-terminal-report-manifest-v1.json`](analysis/child-terminal-report-manifest-v1.json),
35,266 bytes, SHA-256
`b23f755a37fe7d5edb06130647ea8b3d16b91e4cf411a3e5bc02c6ff5a0dd96f`.
The 48 report JSON files total 1,721,162 bytes and the raw child stdout files
total 984,650 bytes; they remain external at
`/home/zhuhe/litchi-performance-0966/capture/doc-sprm-census-v1` and are
bound by the terminal/parse receipts and post-capture audit. No raw capture
files were copied into this repository.

## Sealed source and build

The source revision is
`a499d34f5119a034bc92abde1e1dfcf529159b8d`. The sealed source directory and
all capture artifacts are external:

`/home/zhuhe/litchi-performance-0966/source-head-v1`

The canonical patch is
`/home/zhuhe/litchi-performance-0966/diagnostic-combined.patch` and is copied
to [`repro/diagnostic-combined.patch`](repro/diagnostic-combined.patch). It is
Git-relative and applies cleanly to an archive of the source revision. The
four-file reproduction set includes `tools/perf-cfb-profile/README.md` as a
control; an exhaustive byte comparison found that README unchanged. The
actual patch deltas are:

| File | Sealed SHA-256 |
| --- | --- |
| `crates/litchi-doc/src/sprm.rs` | `66cacbd9a4f10bdead0405094ac5adacd42e523af220acf5990ae2bc335be680` |
| `tools/perf-cfb-profile/Cargo.toml` | `8f440b06101df67672296b5fbf7dd9e857b435efa17adfcaf95d730e077022ef` |
| `tools/perf-cfb-profile/src/lib.rs` | `b6547fb048820b6898443908bc4b81f9895c058ee6091126fe1ee2450e75cc79` |

Patch SHA-256 is
`ce5407f309d2495139401d0e0998a4bfd1456a816819e6c6bf1ece963a067142`
(24,773 bytes). The older absolute `/tmp` patches remain external and are
preserved as authoring receipts; they are not the replay binding.

The external release executable is not copied into this repository. Its build
receipt records 125,930,408 bytes, SHA-256
`d257020c302f92096c34d8bc821f0a88878900ab270f52a189b4e2889eaaf375`, and ELF
build ID `bb0e4bb2f8514bb0350165006209a694e6aaa54b`. The exact build command
was:

```text
RUSTFLAGS= CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=2 cargo +1.95.0 build --offline --locked -j2 --release --bin doc_sprm_census_0966 --manifest-path tools/perf-cfb-profile/Cargo.toml --target-dir /home/zhuhe/litchi-performance-0966/target-0966-census
```

The ignored workspace lock used by that external source export is retained as
[`repro/root-Cargo.lock`](repro/root-Cargo.lock), 138,443 bytes, SHA-256
`d80aa535dcf880a85fa31e92fe2e5ce78b91d93a55953ae734e513aed83e0061`.
It is an overlay receipt, not a newly tracked workspace lock. The standalone
harness lock remains tracked in the sealed source at
`tools/perf-cfb-profile/Cargo.lock` with SHA-256
`343c5464ae1144cbb72284cbb3e9b52b5dff2b798d09462d3c0d6a74b317bac7`.

The build receipt reports 24 SPRM parser tests, two census-control tests, 1,130
DOC tests passed with two ignored, a passing diagnostic release build, a
passing harness compile-only check, and a passing ordinary no-feature check.
It also retains the initial failed check attempts (`cargo-check-initial`,
`cargo-check-borrow-fix`) and the two incomplete broad-test attempts
(`cargo-test-broad-incomplete-overlay`, `cargo-test-broad-overlay-v1`) with
their exit codes. Their full logs remain external under `build-0966/`; this
packet does not copy command logs.

## Reproduction and review bindings

[`receipts/reproduction-v1.json`](receipts/reproduction-v1.json) is preserved
as the historical scratch receipt. Its earlier attempt used the main checkout
as the effective apply cwd while the temporary archive remained unchanged, so
it is not the replay binding. The main checkout was restored to `HEAD` and its
three diagnostic file hashes were checked before and after the corrected run.
[`receipts/reproduction-v2.json`](receipts/reproduction-v2.json) records the
corrected procedure: `/usr/bin/patch --dry-run` and then apply with explicit
`--directory <temporary-scratch> -p1 --fuzz=0`, followed by byte comparison of
all four files and scratch removal. It passed without changing the main
checkout; no build or workload was invoked. The v2 receipt is 4,047 bytes with
SHA-256 `7c0f92b9b0c0fdb1fab62f4635df357678211a499f418aaededaf5f764204987`;
it records the leftover scratch path and timestamps that explain the historical
cwd incident.

The sealed protocol, pre-capture review, corrected source design note
([`analysis/source-design-v2.md`](analysis/source-design-v2.md)), source
provenance, accepted-ADR digest, independent checker, runner, reducer, and
plan-preparation helper are retained in `protocol/`, `review/`, `analysis/`,
`provenance/`, `tools/`, and `receipts/`. The source provenance binds the two
fixture bytes and hashes, source commit, diagnostic source hashes, toolchain,
build receipt, and the external fixture overlay. The accepted ADR digest is a
binding record; it does not amend any ADR. The repository copy of the
corrected design note normalizes four historical links relative to its packet
path; it is 9,825 bytes with SHA-256
`525524a7d907324cb46454d875db93c98c3ced94658cd314e7a4218e50af3a5a`.
The sealed external v2 note remains at
`/home/zhuhe/litchi-performance-0966/doc-sprm-design-v2.md`, 9,704 bytes with
SHA-256
`57f65b245ef4d476f0198c1e45835228501ff7e41a65a28a7a1f732a19242444`; the
original v1 note remains external as a preserved authoring receipt.

The sealed plan is [`protocol/plan-doc-sprm-census-0966-v1.json`](protocol/plan-doc-sprm-census-0966-v1.json),
SHA-256 `18b2915f692567cd830edf9254326431767cb9b0faf79da9eaad1bc9cf0be79b`
(23,187 bytes). It binds the external source tree, executable, checker,
fixture identities, schedule, serial/no-retry policy, and capture root. The
independent sealed review receipt is
[`review/doc-sprm-census-review-receipt-v1.json`](review/doc-sprm-census-review-receipt-v1.json),
with status `pass` and SHA-256
`653beac50c00dc2e5892f006d424f2ebcae72a1365b252466c2cdc17eacba532`.

The independent post-capture audit is
[`review/postcapture-audit-v1.json`](review/postcapture-audit-v1.json),
10,630 bytes, SHA-256
`f6fe42f9896e9895e9d9669bd36daae95ba87dc4d59691871357cbec15ffe293`.
It independently checked all 48 reports and 24 cells, exact repeat identity,
counter and phase conservation, logical-capacity equations, null allocation
payloads, zero errors, false timing claims, raw stdout/report equality, exact
child file sets, and the complete 290-file capture manifest. The 290 capture
files total 2,842,199 bytes; the manifest digest is
`9a5ef1c8d36e5b76d0f222cb58c769b4f0c4c5b7d3bfbc8c9d0f753102dbeb4e`.
It reports `status=pass` and no Office workload beyond the authorized
diagnostic run. The independent checker receipt is 1,003 bytes with SHA-256
`3b6ccb9e5f88a9c4f119f91b5ee2ea605957a8a1711c6da8a3d381efb9d28910`.

The root-owned GO receipt authorized capture; [`protocol/root-go-v1.json`](protocol/root-go-v1.json)
is 358 bytes with SHA-256
`383fcf328f43ef2b9345dffdf8bd3229e9662ba512d17aa406672451695f47a4`.
The external runner first
validated the sealed plan, then retained each child's raw stdout, stderr, JSON
report, and terminal receipt before parsing. The independent reducer consumed
the retained capture root and the checker independently validated all 48 child
reports. [`receipts/checker-receipt-v1.json`](receipts/checker-receipt-v1.json)
records the checker pass; its 48 report inputs totaled 1,721,162 bytes and
the temporary checker directory was removed. The root GO, run-start,
run-terminal, and cache-cleanup receipts are retained in `protocol/` and
`receipts/`; cleanup freed 2,008,809,472 external disk bytes while preserving
the bound executable byte-for-byte. The run-terminal receipt is 181 bytes with
SHA-256 `0385b9ed31b8ebaae84fb822e50b869114d8e051f38d193f681c520036ab5133`.
The supplement helper
[`tools/supplement_doc_sprm_census_0966_v1.py`](tools/supplement_doc_sprm_census_0966_v1.py)
is retained with SHA-256
`5271004dc1dbbd92126c81ea466816019ab86bba1440f42c790109a898018f91`.

To replay the retained analysis, use the sealed plan and external capture root:

```text
PYTHONDONTWRITEBYTECODE=1 python3 tools/run_doc_sprm_census_0966_v1.py --plan /home/zhuhe/litchi-performance-0966/protocol/plan-doc-sprm-census-0966-v1.json --check-plan
REPLAY_DIR=$(mktemp -d /tmp/litchi-0966-replay.XXXXXX)
PYTHONDONTWRITEBYTECODE=1 python3 tools/reduce_doc_sprm_census_0966_v1.py --plan /home/zhuhe/litchi-performance-0966/protocol/plan-doc-sprm-census-0966-v1.json --capture /home/zhuhe/litchi-performance-0966/capture/doc-sprm-census-v1 --output "$REPLAY_DIR/census-summary-replay.json"
```

The replay output must be a new path because the reducer refuses to overwrite
an existing retained file.

The 48 raw child outputs remain external and are bound by the terminal/parse
receipts and post-capture audit; no raw capture files were copied into this
repository.

The bounded packet inventory is
[`receipts/packet-files-v1.json`](receipts/packet-files-v1.json). It lists every
packet file except itself with its byte count and SHA-256; all listed files are
below 1 MiB and no executable is included.

## Interpretation boundary

The strongest permitted statement is a bounded current-source call/capacity
census for this named fixture and lifecycle matrix, including its repeat and
phase observations. Empty calls are not
allocator savings, logical capacity is not allocator traffic, and a count
relationship is not a latency or causality claim. A production candidate would
need separate source-preserving correctness, adversarial/error, native timing,
and allocator qualification; this diagnostic packet provides no qualification
for those claims.
