# 0395 — D7 unified table-cell root comments

- Status: bounded D7 implementation slice on `uo2/iwa-d7-root-comments`, cut
  from coordinator root `1bc8f22499d2691e85ebed37686c8381f0c985da`; runtime
  and tests are committed at `6be0814e1b5151945eba926a3574d6666f54c73f`.
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
  and focused-refusal classifications.
- The shared registry/spec helper marks only the existing `comment` property
  as clearable-writable. No new `Comment` element or compatibility alias was
  added. Numbers runtime dispatch is unchanged.

## Proof and boundaries

The Pages and Keynote unified tests cover populated reads, root replacement,
rootless-cell creation, `null` clear, exact no-op bytes, source-bound inverse,
reopen/readback, unrelated-cell locality, and a missing-selector terminal
refusal. Numbers retains its existing read/set/create/clear regression and
adds the clearable capability assertion. The host-backed leaf differential
cases remain the oracle: Pages `pages_body_table_comment_*_parity`, Keynote
`keynote_slide_table_comment_*_parity`, and Numbers
`numbers_table_cell_comment_*_parity`.

No ledger rows were changed in this D7 dispatch slice. The nine unified
Pages/Keynote/Numbers read, set, and null-clear routes were appended to
`0363-mac-queue.md`; all remain pending. No native E3/E4 acceptance is
claimed, and no raw captures or binaries were added.

## Verification

The scoped checks run on the final source include:

```text
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo check --offline -p litchi-ops --all-features --lib --tests
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test pages_unified body_table_cell_comments_round_trip_through_the_unified_route -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test keynote_unified cell_comment -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test numbers_unified numbers_cell_comment_schema_is_clearable -- --nocapture
CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test numbers_unified cell_comments_are_read_set_added_and_removed -- --nocapture
LITCHI_BLESS=1 CARGO_NET_OFFLINE=true CARGO_BUILD_JOBS=4 cargo test --offline -p litchi-ops --features pages,keynote,numbers --test registry --test capabilities -- --nocapture
```

Full coordinator-owned performance qualification and the integrated 53-row
gate remain required after merge; this record makes no File-vs-leaf
performance claim.
