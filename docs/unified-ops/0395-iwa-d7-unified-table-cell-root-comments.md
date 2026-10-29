# 0395 — D7 unified table-cell root comments

- Status: bounded D7 implementation slice on `uo2/iwa-d7-root-comments`, cut
  from coordinator root `1bc8f22499d2691e85ebed37686c8381f0c985da`; the final
  source is `a24f121a1421ed596e07a5f6fbe8c67adce5f33d` (clean). Runtime/proof
  history is `6be0814e1b5151945eba926a3574d6666f54c73f` through
  `b6bbb77202227a1be4202fb63537a53fd121edc2`; the bounded marker probe is
  `d4de8d6adc24924c8c071ee6ce0367f45982df54`, with its observed-record repair
  at the final tip.
- Scope: expose the existing focused root-comment leaf transactions through
  the selector-first `cell.comment` property on Pages and Keynote, and make
  the existing Numbers route truthfully declare `null` clearability.
- Contract: `cell.comment` reads and sets text; `null` clears it. A set on a
  rootless in-grid cell creates the root. Replies, authors, timestamps,
  standalone comment nodes, and drawable/text-box comments remain outside
  this route. Focused refusals are terminal and never fall through to the
  legacy host.

## Implementation

- `crates/litchi-ops/src/format/iwork/pages.rs` projects sparse table comments
  into `tc.comment` and stages set/clear through
  `edit_body_table_cell_comment`.
- `crates/litchi-ops/src/format/iwork/keynote.rs` and `keynote/write.rs`
  provide the matching slide-table projection and table-ordinal selector
  adapter through `edit_slide_table_cell_comment`.
- `format/iwork/error.rs` carries the Pages/Keynote comment errors through
  `Source::Iwork` with explicit limit, allocation, source, patch, protected,
  and focused-refusal classifications. Missing tables remain `not_found`,
  while ambiguous table names/selectors remain `invalid_value`.
- The shared registry/spec helper marks only the existing `comment` property
  as clearable-writable. No new `Comment` element or compatibility alias was
  added. Numbers runtime dispatch is unchanged.
- `crates/litchi-ops/src/open/bytes.rs` uses a bounded central-directory-only
  iWork marker probe in an all-iWork build. It checks exact raw member names,
  preserves ODF/OOXML precedence, caps directory bytes and observed records,
  and leaves malformed/false markers to the existing typed detector route.

## Proof and boundaries

The Pages and Keynote unified tests cover populated reads, genuine File-level
rootless-cell creation, root replacement, `null` clear, exact no-op bytes,
source-bound inverse, reopen/readback, stale-recipient refusal, atomic
multi-cell batches, untouched comment metadata/replies, an untouched ZIP
member, and missing-selector terminal refusals. Numbers retains its existing
read/set/create/clear regression and adds the clearable capability assertion.
The host-backed leaf differential
cases remain the oracle: Pages `pages_body_table_comment_*_parity`, Keynote
`keynote_slide_table_comment_*_parity`, and Numbers
`numbers_table_cell_comment_*_parity`.

No ledger rows were changed in this D7 dispatch slice. The nine unified
Pages/Keynote/Numbers read, set, and null-clear routes were appended to
`0363-mac-queue.md`; all remain pending. No native E3/E4 acceptance is
claimed, and no raw captures or binaries were added.

## Full-route cost probes

Eight ignored probes now cover the full `File` route against the direct
focused leaf route: Pages and Keynote each have read, rootless create, root
replace, and root clear. Read probes compare the exact semantic text output;
write probes compare the complete published bytes before any counted loop.
The measured symbols include only `open_bytes` plus the selector-first File
operation and its output, or `Package::from_bytes` plus the focused leaf
operation and corresponding output. Fixture preparation, source hashes,
warm-up equality, no-op/readback/inverse checks, and source immutability stay
outside those symbols. Pages replacement/clear uses a root-only source
authored from the tracked table fixture; Pages/Keynote creation uses a
rootless stored cell. Numbers has no new performance claim because its D7
change is schema-only; its existing comment read/set/create/clear regression
remains required.

Each capture must run release mode on a coordinator-pinned CPU with Callgrind,
ABBA order `[unified, leaf, leaf, unified]`, one exact ignored test selected,
exit 0, positive selected-symbol Ir, and immutable binary/source hashes. The
instruction acceptance bar is candidate/direct `<= 1.05`; a 20-microsecond
allowance requires separate valid wall-time evidence and is not inferred from
Ir. Raw captures stay outside the repository and final qualification must use
the integrated coordinator commit.

The final source-bound qualification used CPU3 and the release build command
below. The binaries are external, immutable copies: Pages is 400,814,536 bytes
(SHA-256 `d5b6c92c7bb49ba23640ac8563c656680e8e8f612697dfa8171812c542cec02c`),
and Keynote is 394,295,392 bytes (SHA-256
`adf4d93cbd21114095727cd09ff3301ab06559045528848ed9aadaf9f25c01f4`).

```text
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/d7-comments-perf-b6bbb77202/target CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 CARGO_BUILD_JOBS=4 RUST_TEST_THREADS=1 CARGO_NET_OFFLINE=true cargo test --release --locked --offline -p litchi-ops --all-features --test pages_unified --test keynote_unified --no-run
```

The build log is external at
`/home/zhuhe/code/litchi-worktrees/.coord/d7-comments-perf-b6bbb77202/build-a24f121a14.log`
(SHA-256 `c2b8447f8b843f5e7506d1a6e614cddf23074db0dbfcd3b5f4e4cbbb8287d30d`).
The complete compact receipt, including all 16 raw capture/log/annotation
hashes, is
`/home/zhuhe/code/litchi-worktrees/.coord/d7-comments-perf-b6bbb77202/coordinator-review-a24f121a14.json`
(SHA-256 `a40afa4d618c3579b8602151e25a2d2887e1c4fec347145fd3327ec149a9199b`).
The raw captures and threshold-100 inclusive annotations are under the same
external directory in `callgrind-a24f121a14/`; the earlier d4 captures remain
preserved as nonqualifying history because they predate the observed-record
boundedness repair.

| full-file route | unified/direct Ir | ratio | bar |
| --- | ---: | ---: | --- |
| Pages read | 34,161,006 / 33,354,393 | 1.024183111 | pass |
| Pages rootless create | 100,835,958 / 99,991,031 | 1.008450028 | pass |
| Pages root replace | 91,154,644 / 90,331,540 | 1.009112033 | pass |
| Pages root clear | 165,344,083 / 164,391,172 | 1.005796607 | pass |
| Keynote read | 124,150,720 / 122,796,077 | 1.011031647 | pass |
| Keynote rootless create | 430,888,063 / 429,766,675 | 1.002609295 | pass |
| Keynote root replace | 418,344,439 / 417,134,881 | 1.002899681 | pass |
| Keynote root clear | 484,925,916 / 483,453,742 | 1.003045119 | pass |

All 16 legs exited 0, selected exactly one test, reported positive Ir, and
showed exactly two calls to the selected route wrapper in the raw Callgrind
function records. No wall-time or native E3/E4 claim is made.

## Verification

The scoped checks run on the final source include:

```text
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo check --offline -p litchi-ops --all-features --lib --tests
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test pages_unified body_table_cell_comments_round_trip_through_the_unified_route -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=2 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test pages_unified body_table_cell_comment -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test keynote_unified cell_comment -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=2 cargo test --offline -p litchi-ops --features pages,keynote,numbers --lib format::iwork::error::tests::table_comment_selection_errors_keep_missing_and_ambiguous_codes -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test numbers_unified numbers_cell_comment_schema_is_clearable -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test numbers_unified cell_comments_are_read_set_added_and_removed -- --nocapture
LITCHI_BLESS=1 CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test registry --test capabilities -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --all-features --lib marker_tests
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --all-features --test pages_unified --test keynote_unified
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo clippy --offline -p litchi-ops --all-features --all-targets -- -D warnings
```

The focused Pages and Keynote suites passed 21 and 25 non-ignored tests on the
final source; the five marker tests and strict all-target/all-feature Clippy
also passed. No raw captures or binaries were added to the repository. The
coordinator still owns ledger disposition, native queue execution, final
integrated requalification, and the 53-row gate.
