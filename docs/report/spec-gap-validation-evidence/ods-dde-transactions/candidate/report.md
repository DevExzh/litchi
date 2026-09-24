# ODS inert DDE transaction profile — final candidate

This receipt measures the frozen DDE transaction candidate applied as a scoped overlay to baseline commit `5347801fc2e8086d67fc6c140a023aeedd260c91` (`perf(ods): binary search sheet metadata cell selectors`, committed `2026-09-13T02:35:01Z`). The candidate transaction source hash is `bce59b24064e017d0a3f3cf945594fdd9870e84cdeda513d076ea9c86fc1bc4c`; the focused transaction test hash is `5bad85510fcf7b2444e9b37ea254bd16d14c42b13ff6b6ab97b982da2660064b`. The candidate binary is recorded in [`binary.sha256`](binary.sha256), the 18-lane receipt hash is `6566b32570dac184057810fb5a2dd39e42fca0441def6940329f5ce6a8723e58`, and source/overlay manifests are [`../gates/source-after-gates.json`](../gates/source-after-gates.json) and [`source-hashes.sha256`](source-hashes.sha256).

The overlay contains only the ODS sources and the retained harness; unrelated OPC/XLSX working-tree changes were excluded. Final repository gates (all-target ODS, clippy, rustdoc, doctests, formatting, and the API authoring runner) completed with exit status 0. The isolated overlay test run in [`scoped-tests-final.log`](scoped-tests-final.log) ran 12 facade and 33 transaction tests, all passing.

## Scope and method

The standalone harness generates valid ODF 1.4 `content.xml` streams with required `table:table-column` children and one empty worksheet cell. It uses named cache tables with attribute-only cells and inert `file:///never/...` DDE topics; it never contacts a producer and does not build or publish a ZIP package. The three shapes are:

- `small`: 1 worksheet source, 2 links, 8×8 cache, 9,845 source bytes.
- `large-cache`: 1 worksheet source, 1 link, 256×256 cache (65,536 cells), 4,132,098 source bytes.
- `many-links`: 4 worksheet sources, 2,048 links, 1×1 caches, 863,375 source bytes.

Each lane used three warmups and 15 measured iterations. The process was pinned to CPU 2 with `taskset`; `/usr/bin/time -v` supplied whole-process maximum RSS. The host was shared with other services and was not otherwise isolated, so wall-clock values are scoped observations rather than machine-wide performance claims. The full command list and per-lane stdout, status, allocator, and RSS receipts are retained beside [`raw.csv`](raw.csv).

`parse` times `Snapshot::parse_with_context` and includes XML parsing/indexing, validation, cancellation checks, and budget accounting. `read` parses before the timer, then reads all worksheet-source fields, link-source fields, cache XML lengths, and table names; it measures accessor traversal only. `stage` parses and constructs the caller's replacement `LinkSpec` before the timer, then times `Snapshot::edit()` and one source-order link replacement (the many-links lane also reads the sheet-source count). `commit` parses and stages before the timer and times commit rendering and readback. `noop` parses and creates an edit before the timer and times its no-op commit. `edit` constructs the typed replacement before the timer, then times parse, stage, and commit; it is a metadata edit workflow and excludes caller spec construction and package publication.

The harness's counting allocator reports requested/released bytes and live-byte deltas for the timed region. `Work` and `Memory` are execution-budget counters, not allocator bytes. A staged `LinkSpec` is caller-built before the stage timer; its retained cache payload is admitted to the transaction's `Memory` budget during staging even when the move itself makes few allocator calls. `Maximum resident set size` is the operating-system whole-process value and is separate from the budget counters. `probes` is cumulative over all 15 measured iterations, while each Work/Memory and allocator p50 is per iteration; the raw file retains before/after values and maxima.

## Final candidate measurements

Times are `mean / p50 / p95 / p99`; byte and count pairs are `p50 / max`. Requested and released bytes are shown as `requested p50/max ; released p50/max`.

### Parse and read

| lane | input B | mean / p50 / p95 / p99 ns | alloc calls p50/max | requested / released B p50/max | allocator peak live Δ B p50/max | Work Δ p50/max | Memory Δ p50/max | Output B p50/max | RSS KiB | probes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| parse-small | 9,845 | 79,800 / 78,640 / 91,360 / 91,360 | 217 / 217 | 25,997 / 25,997 ; 14,515 / 14,515 | 12,938 / 12,938 | 10,034 / 10,034 | 10,751 / 10,751 | 0 / 0 | 2,252 | 0 |
| parse-large-cache | 4,132,098 | 29,961,434 / 29,905,033 / 30,672,347 / 30,672,347 | 65,604 / 65,604 | 8,332,898 / 8,332,898 ; 4,199,205 / 4,199,205 | 4,135,149 / 4,135,149 | 4,198,169 / 4,198,169 | 4,132,719 / 4,132,719 | 0 / 0 | 14,416 | 0 |
| parse-many-links | 863,375 | 10,015,213 / 10,005,854 / 10,059,714 / 10,059,714 | 45,184 / 45,184 | 4,579,628 / 4,579,628 ; 3,329,246 / 3,329,246 | 1,251,838 / 1,251,838 | 881,845 / 881,845 | 1,463,501 / 1,463,501 | 0 / 0 | 4,544 | 0 |
| read-small | 9,845 | 26 / 30 / 30 / 30 | 0 / 0 | 0 / 0 ; 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 2,244 | 60 |
| read-large-cache | 4,132,098 | 50 / 30 / 180 / 180 | 0 / 0 | 0 / 0 ; 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 14,348 | 45 |
| read-many-links | 863,375 | 4,443 / 4,430 / 4,530 / 4,530 | 0 / 0 | 0 / 0 ; 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 0 / 0 | 4,584 | 30,840 |

### Transaction phases

| lane | input B | mean / p50 / p95 / p99 ns | alloc calls p50/max | requested / released B p50/max | allocator peak live Δ B p50/max | Work Δ p50/max | Memory Δ p50/max | Output B p50/max | RSS KiB | probes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| stage-small | 9,845 | 324 / 320 / 370 / 370 | 5 / 5 | 520 / 520 ; 8 / 8 | 520 / 520 | 4 / 4 | 2,657 / 2,657 | 0 / 0 | 2,240 | 0 |
| stage-large-cache | 4,132,098 | 348 / 340 / 470 / 470 | 5 / 5 | 360 / 360 ; 8 / 8 | 360 / 360 | 3 / 3 | 2,103,553 / 2,103,553 | 0 / 0 | 12,592 | 0 |
| stage-many-links | 863,375 | 16,740 / 16,620 / 17,580 / 17,580 | 5 / 5 | 327,880 / 327,880 ; 8 / 8 | 327,880 / 327,880 | 2,050 / 2,050 | 327,833 / 327,833 | 0 / 0 | 4,788 | 0 |
| commit-small | 9,845 | 168,230 / 167,621 / 173,071 / 173,071 | 856 / 856 | 132,179 / 132,179 ; 119,925 / 119,925 | 35,544 / 35,544 | 21,221 / 21,221 | 11,089 / 11,089 | 10,171 / 10,171 | 2,336 | 0 |
| commit-large-cache | 4,132,098 | 72,537,184 / 72,528,513 / 72,792,094 / 72,792,094 | 393,381 / 393,381 | 68,176,652 / 68,176,652 ; 64,042,195 / 64,042,195 | 25,243,341 / 25,243,341 | 9,315,827 / 9,315,827 | 4,133,057 / 4,133,057 | 4,132,424 / 4,132,424 | 36,640 | 0 |
| commit-many-links | 863,375 | 20,155,358 / 20,145,870 / 20,268,790 / 20,268,790 | 112,924 / 112,924 | 448,348,123 / 448,348,123 ; 447,096,969 / 447,096,969 | 2,360,524 / 2,360,524 | 1,528,777 / 1,528,777 | 1,463,839 / 1,463,839 | 863,701 / 863,701 | 8,296 | 0 |
| noop-small | 9,845 | 267 / 250 / 390 / 390 | 1 / 1 | 176 / 176 ; 0 / 0 | 176 / 176 | 0 / 0 | 0 / 0 | 0 / 0 | 2,244 | 0 |
| noop-large-cache | 4,132,098 | 256 / 250 / 340 / 340 | 1 / 1 | 176 / 176 ; 0 / 0 | 176 / 176 | 0 / 0 | 0 / 0 | 0 / 0 | 14,356 | 0 |
| noop-many-links | 863,375 | 704 / 270 / 6,570 / 6,570 | 1 / 1 | 176 / 176 ; 0 / 0 | 176 / 176 | 0 / 0 | 0 / 0 | 0 / 0 | 4,548 | 0 |
| edit-small | 9,845 | 247,405 / 245,801 / 257,271 / 257,271 | 1,078 / 1,078 | 158,696 / 158,696 ; 134,448 / 134,448 | 47,538 / 47,538 | 31,259 / 31,259 | 24,497 / 24,497 | 10,171 / 10,171 | 2,596 | 0 |
| edit-large-cache | 4,132,098 | 103,772,429 / 103,633,500 / 104,554,365 / 104,554,365 | 458,990 / 458,990 | 76,509,910 / 76,509,910 ; 68,241,408 / 68,241,408 | 29,377,386 / 29,377,386 | 13,513,999 / 13,513,999 | 10,369,329 / 10,369,329 | 4,132,424 / 4,132,424 | 36,748 | 0 |
| edit-many-links | 863,375 | 30,373,105 / 30,389,925 / 30,511,566 / 30,511,566 | 158,113 / 158,113 | 453,255,631 / 453,255,631 ; 450,426,223 / 450,426,223 | 3,938,778 / 3,938,778 | 2,412,672 / 2,412,672 | 3,255,173 / 3,255,173 | 863,701 / 863,701 | 8,292 | 0 |

The corrected large-cache admission is visible in `stage-large-cache`: `Memory` Δ is 2,103,553 bytes for the 65,536-cell payload, while timed allocator requests are only 360 bytes because the payload was prepared before the timer and adopted by the draft. The corresponding `edit-large-cache` `Memory` Δ is 10,369,329 bytes. Changed stage, commit, and edit lanes report positive Work; no-op lanes report zero Work and zero Memory delta.

## Baseline comparison and limits

The read-only baseline used the same three generated source shapes and the same 3-warmup/15-iteration harness, but exposed only `Snapshot::parse` and read accessors. It has no transaction phases or execution context, so no baseline speedup claim is made for stage, commit, no-op, or edit. Candidate `read` also reads table names and therefore has a different probe set; it is reported absolutely rather than treated as a timing comparison.

For the comparable parse operation, candidate p50 versus baseline p50 is:

| shape | baseline `Snapshot::parse` p50 ns | candidate `parse_with_context` p50 ns | candidate / baseline | candidate Work Δ | candidate Memory Δ | candidate RSS KiB |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| small | 27,950 | 78,640 | 2.81× | 10,034 | 10,751 | 2,252 |
| large-cache | 8,203,595 | 29,905,033 | 3.65× | 4,198,169 | 4,132,719 | 14,416 |
| many-links | 5,755,525 | 10,005,854 | 1.74× | 881,845 | 1,463,501 | 4,544 |


The largest scoped parse regression is `large-cache`: 29,905,033 ns p50 versus 8,203,595 ns, or 3.65× on this run. `small` is 2.81× and `many-links` is 1.74×. The candidate path adds context-aware validation and budget accounting, which is a possible contributor to these observations; this harness has no CPU profile, so it does not establish causation. Large-cache allocator peak-live Δ is 4,135,149 bytes versus the baseline's 4,134,342 bytes, and candidate maximum RSS is 14,416 KiB versus 14,388 KiB. These are measurements of the three synthetic streams on one shared host.

Replay the candidate from the baseline with a separate worktree:

```sh
EVIDENCE=/absolute/path/to/ods-dde-transactions
BASE=/var/tmp/ods-dde-replay
TARGET=/var/tmp/ods-dde-replay-target
git worktree add --detach "$BASE" 5347801fc2e8086d67fc6c140a023aeedd260c91
git -C "$BASE" apply --unidiff-zero "$EVIDENCE/candidate/dde-api.diff"
mkdir -p "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/candidate/harness/src"
cp "$EVIDENCE/candidate/harness/Cargo.toml" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/candidate/harness/Cargo.toml"
cp "$EVIDENCE/candidate/harness/Cargo.lock" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/candidate/harness/Cargo.lock"
cp "$EVIDENCE/candidate/harness/src/main.rs" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/candidate/harness/src/main.rs"
CARGO_TARGET_DIR="$TARGET" cargo build --manifest-path \
  "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/candidate/harness/Cargo.toml" \
  --release --locked --offline
taskset -c 2 "$TARGET/release/ods-dde-transaction-profile" \
  --workload stage --case large-cache --warmups 3 --iterations 15
git worktree remove --force "$BASE"
```

This evidence covers inert metadata parsing and the transaction phases exercised above. It does not establish DDE producer execution, refresh behavior, native interoperability, complete package publication cost, or general ODS performance.
