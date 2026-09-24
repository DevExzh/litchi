# ODS data-style graph preflight evidence

This directory retains the matched before/after allocation receipt for the
bounded `put_extended_style_graph` preflight optimization. It records the exact source hashes under review and does not establish a general latency claim.

## Change under measurement

The candidate reuses the source catalog's already identified direct
`office:automatic-styles` range, content length, and opening-tag shape when
sizing the candidate. The old path reparsed all content spans for the preflight
size calculation and then parsed them again for insertion. The candidate still
runs the existing source-qualified scan, insertion seam, output limit checks,
and graph readback verification. No cache or full-package copy was added.

The baseline is commit `8e3ad310426a534c0bb17a789eb2c42a96e61310`. The candidate
is that same clean worktree with only `advanced.rs` overlaid at hash
`20eb69aa8402ac5c621efdf767a1069b89712f327e83d03d2f0a305a55e0822a`.
`data_style/source.rs` was unchanged at
`97121798e76fda01d7d12a1be10c4a34457d6447e1fb5de6184a86e5b2629261`.

## Harness and capture

`harness/` is the exact harness used for both runs. Its lockfile is retained;
source and lock hashes are in `hashes.txt`. Both runs used the process-global
counting allocator, scales 8, 128, and 512, and seven iterations per operation.
The operations were source query, snapshot clone, metadata no-op, scalar patch,
graph put, graph replacement, same-value graph replacement, and graph removal.
The raw receipts are `raw-before.jsonl` and `raw-after.jsonl`. Fixture archive
hashes are in `fixture-sha256.txt`; before and after fixture hashes match.

At scale 512, the deterministic `graph_put` allocation fields were:

| field | baseline | candidate |
| --- | ---: | ---: |
| median elapsed microseconds | 11,327.612 | 10,255.260 |
| allocations | 75,437 | 67,018 |
| requested bytes | 26,382,083 | 24,208,708 |
| peak live-byte delta | 1,743,415 | 1,743,415 |
| source bytes | 8,610 | 8,610 |
| result bytes | 8,244 | 8,244 |
| explicit copy bytes | 8,610 | 8,610 |

The allocation reduction is 8,419 calls (11.16%) and 2,173,375 requested
bytes (8.24%), with unchanged peak live bytes and output/copy fields. Elapsed
time is included as a matched observation; the allocator fields are the
bounded resource evidence.

The exact commands and temporary capture paths are in `commands.txt`, with
binary/toolchain details in `toolchain-and-binaries.txt`.

## Isolated validation

The candidate was validated in the same clean baseline worktree after the
profile capture:

- `data_style_vocabulary`: 47 passed, 0 failed;
- `cargo clippy -p litchi-ods --lib --offline -- -D warnings`: passed;
- rustfmt and `git diff --check`: passed for the owned source files.

The corresponding logs are retained under `gates/`. The root worktree has
other unrelated ODS `sheet_metadata` changes, so these isolated logs bind the
result specifically to the baseline-plus-`advanced.rs` candidate described
above.

## Root verification and measurement limits

The root independently checked both 168-row receipt sets (24 groups, seven
repetitions each). Allocator balance identities hold and every net live-byte
delta is zero. Graph-put allocation counts, deallocation counts, requested bytes and released bytes changed; source/result/copy fields and logical peak were unchanged. Among the selected deterministic metrics checked, changes were confined to graph put; see `root-receipt-verification.json`.
An independent checkout based on `5d21e069d`, with only the candidate
`advanced.rs` overlaid, passed 298 library tests, 47 data-style integration
tests and strict library Clippy without warning suppression. Its exact hashes,
commands and logs are under `root-gates/`.

Peak memory here means allocator-observed logical live bytes. It excludes RSS,
allocator headers and transient old-plus-new storage inside `System::realloc`.
Deallocation counts cover explicit deallocation calls. Timings include the
harness validation, parsing, edit staging and member extraction; they do not
measure durable publication. Explicit copy bytes count only instrumented
harness copies. Graph lanes do not assert every unrelated package member or
an exact inverse. Functional tests cover selected XML preservation cases and some graph replacement inverses; they do not establish a full graph-put inverse or preservation of every foreign binary member. This matched run supports the stated allocation
reduction on these fixtures; its elapsed times are observations, not a broad
throughput guarantee.

## Reproduction

Run from a checkout containing the baseline commit and this evidence:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-data-style-graph-preflight-performance/replay.py /var/tmp/ods-graph-preflight-replay-results
```

The output path must not exist. The script creates two clean baseline
checkouts and applies the retained `candidate.patch.gz` only to the candidate.
It verifies both source hashes, uses the captured harness and lockfile, runs
seven repetitions, retains receipts and binary hashes in the requested output,
and removes its temporary checkouts and build targets. Offline dependency
availability is required. The root ran this command successfully; both
168-measurement captures matched the original deterministic metrics and all
nine fixture hashes. Its receipts and verification are under `root-replay/`.

## Capture notes and review

Independent review approves the matched allocation characterization and the
source optimization within the scopes above. The original commands used the
sibling `ods-data-style-source-performance/harness` path. Its three harness
files are byte-identical to the retained copy here; the replay script copies
this retained harness into that same path in each clean checkout.

Historical harness comments are retained byte-for-byte for hash fidelity.
The comment about warming parser caches describes one discarded invocation;
no process parser cache is established. The comment attributing inverse proof
to scalar warmup is inaccurate: `check_scalar_inverse` runs after the measured
lanes. Neither comment changes the retained measurement scope.
