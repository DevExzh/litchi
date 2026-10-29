# 0395 — D7 unified table-cell root comments

- Status: bounded D7 implementation slice on `uo2/iwa-d7-root-comments`, cut
  from coordinator root `1bc8f22499d2691e85ebed37686c8381f0c985da`; runtime,
  proof tests, and ignored cost probes are committed at
  `e0622be0e300ca4e405790c4b309b8e230a6f63b`.
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
```

Full coordinator-owned performance qualification and the integrated 53-row
gate remain required after merge; this record makes no File-vs-leaf
performance claim.
