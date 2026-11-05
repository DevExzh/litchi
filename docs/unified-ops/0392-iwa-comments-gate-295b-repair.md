# 0392 — W1-G comments gate repair on the 295b candidate

- **Base:** `295b421bcd3d9ec0b0b9208e4039de22623685bc`
- **Implementation commits:** `5fc2f905a0`, `56020e9214`, strict-clear
  admission checkpoint `c4a823fc58`, mirror repair `9d0d581e8f`, proof
  merge `d34aacac9a`, and final runtime repair `67427c20f1`. The
  proof-only follow-up `6d45e19a61` was merged as `6eebe0bab3` on
  `uo2/iwa-comments-gate-295b-repair`.
- **Status:** the scoped Numbers/shared-engine repair, 13-case
  Pages/Keynote/Numbers clear-safety proof, original comment suites, scoped
  gates, and current-source instruction qualification are green on this
  clean WP. Coordinator integration, final integrated requalification, and
  native acceptance remain pending.

## Scope and root causes

The 295b full gate exposed five regressions in the existing Numbers comment
route. The repair stays in `litchi-iwa-table/src/comments_native.rs`; the
Numbers facade continues to publish the selector-first API and the shared
root operation.

1. Numbers comment roots live in the DataList component, not necessarily in
   the selected table-model component. Existing-list admission now validates
   the selected roots through the source census without rejecting this valid
   multi-component profile.
2. `ListKeys::Reuse` applies to a newly-created Numbers list. An existing
   list still allocates from its source `next_list_id` watermark, avoiding a
   key-2 collision on populated lists.
3. A single registered author is reused. When a multi-author Numbers graph
   has no existing `litchi-iwa` author, creation allocates a new generated
   author and appends it to the existing registry. Payload-only native author
   registries retain their header shape.
4. Existing table-list headers vary between aggregate-only, generic
   `FieldInfo`, and keyed `FieldInfo` mirrors. List metadata updates now
   preserve those source shapes while checking selected reference counts and
   removing only the selected object mirror.
5. A destructive clear admits exactly one type-3056 message with matching
   `ArchiveInfo` metadata, rejects empty table-info candidates, and rejects
   unknown top-level comment payload fields before mutation. Unknown fields
   remain supported by non-destructive read/text paths.
6. Clear admission now pre-charges and runs a package-wide
   `RejectUnknownMetadata` archive-reference census, binds the deleted root
   to the selected list message, checks metadata UUID/external ownership,
   validates selected storage/list mirrors, and strictly decodes table-info,
   model, tile, and list payloads. Generic and keyed list mirrors are
   preserved; only the selected keyed/ordinal mirror may be removed.

No public selector, convenience API, ledger disposition, or native evidence
claim was added by this repair. Unsupported reply authoring and native
acceptance remain as previously documented in the W1-G records and queue.

## Regression and focused verification

The original gate failures now pass without weakening the existing assertions:

| Check | Result |
|---|---:|
| `comment_authors` generated-author reuse/new-author cases | 2 passed |
| `comment_corpus` corpus and workbook round trips | 2 passed |
| `table_data_list_reader_integration` | 25 passed |
| `table_cell_comment_reply` | 31 passed |
| `table_cell_comment_metadata_native` | 5 passed |
| `numbers_focused_comment_reads` | 17 passed |
| `litchi-iwa` clear-safety parity module | 13 passed |

The clear-loss controls retain their exact typed outcomes and source bytes:

- noncanonical metadata framing → `InvalidSource`;
- an opaque sibling message in the selected storage object →
  `UnsupportedDependency`;
- an empty type-6000 table-info candidate → `InvalidSource`;
- unknown top-level type-3056 payload field → typed refusal.

The valid root-only clear, inverse, candidate reopen, locality, no-op, and
reply-bearing refusal cases remain green. Host-backed Pages/Keynote clear
baseline and negative proof is recorded in
[0393](0393-iwa-comments-clear-safety-proof.md) and runs against the repaired
shared engine.

## Scoped gates

Passed on the clean implementation source (release/debug settings are stated
per command):

- `CARGO_BUILD_JOBS=2 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0
  CARGO_PROFILE_TEST_DEBUG=0 cargo fmt --all -- --check`;
- `cargo clippy --offline -p litchi-iwa-table --all-targets --all-features
  -- -D warnings`;
- `cargo clippy --offline -p litchi-numbers --all-targets --all-features
  -- -D warnings`;
- `RUSTDOCFLAGS=-Dwarnings cargo doc --offline --no-deps --all-features
  -p litchi-iwa-table -p litchi-numbers`;
- `python3 tools/check_crate_boundaries.py`;
- `python3 tools/iwa_exit_ledger.py --check`;
- focused `tools.test_check_crate_boundaries`, `tools.test_iwa_exit_ledger`,
  and `tools.test_no_3rdparty_references`;
- locked Numbers and iwa-protos fuzz-workspace `cargo check --bins` checks.

The affected generated documents were checked against `target/api-glossary`
with `RUSTC_BOOTSTRAP=1` using `--profile iwa-table`,
`--profile numbers-minimal`, and `--profile numbers-full`; all three passed.
The current external receipts are under
`(deleted-coord-artifact)`.
The original six Numbers comment suites and the 13-case clear-safety run were
rerun after the proof-only merge; their external receipts are
`original-six/run.log` and `gates/clear-safety-13.log` in that directory.

## Current-source instruction qualification

The authorized 12-route matrix was rerun on runtime source
`67427c20f10dc51f45e772f08da6c1763a04505d` using the feature-enabled release
parity test binary. The immutable binary is external at
`(deleted-coord-artifact)/parity-67427c20f10dc51f45e772f08da6c1763a04505d`, size 41,380,128 bytes, SHA-256 `bf8ff72a505f712b5886faba58200582da8881c2ca3159ccb56dc76b0c860ae1`. The build receipt and runner result are in the same external directory; the runner used CPU 3, Callgrind ABBA, release mode, and fresh external output `measure-12-final`.

All 12 routes passed, yielding 24 callgrind captures and 36 exact-test logs; every route stayed within the 5% slowdown bar:

| Route family | Host Ir | Leaf Ir | Delta |
|---|---:|---:|---:|
| Keynote set/read/create/clear | 658,091,493 / 20,345,349 / 904,991,498 / 58,125,869 | 417,208,237 / 8,042,759 / 429,841,475 / 41,952,904 | −36.603% / −60.469% / −52.503% / −27.824% |
| Pages set/read/create/clear | 2,591,604,039 / 1,234,965,546 / 2,735,193,618 / 2,741,075,485 | 90,331,205 / 33,549,613 / 99,981,254 / 165,487,524 | −96.514% / −97.283% / −96.345% / −93.963% |
| Numbers set/read/create/clear | 73,280,745 / 3,711,055 / 169,352,648 / 128,335,767 | 74,009,105 / 2,967,400 / 82,626,646 / 118,322,350 | +0.994% / −20.039% / −51.210% / −7.803% |

The later `6eebe0bab3` merge changes only proof tests and records; it does
not change the runtime source bound to this matrix. These are
engine/facade instruction-cost measurements, not Apple native acceptance
evidence. Final integrated-trunk requalification remains coordinator-owned.

## Evidence boundaries and remaining qualification

This repair has no native E3/E4 evidence. The current matrix above is the
source-bound host-vs-leaf instruction-cost qualification for runtime
`67427c20f1`; the earlier failed `b4fc6f2884` Keynote capture remains
historical regression context and is not qualification evidence. The
coordinator must run the final host-vs-leaf instruction-cost matrix again on
the integrated candidate after all wave merges, then update `PROGRESS.md` and
perform the complete 53-row gate on one trunk commit.
