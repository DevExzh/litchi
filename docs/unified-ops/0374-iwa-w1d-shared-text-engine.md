# 0374 — W1-D: the shared text engine out of `litchi-iwa`

- Status: wave W1-D landed on branch `uo2/iwa-w1d`. Trunk
  `feat/unified-ops-api` was merged at `bea5fdddf9` and `6641a0f94e`. The
  coordinator merged the branch back in batches; the last merge is
  `b7278b5cb8`.
- Date: 2026-10-26.
- Plan: wave row W1-D of [0363](0363-iwa-monolith-exit-plan.md) (§0), API
  shape §3.2 and parity §4. The harness and ledger are described in
  [0369](0369-iwa-exit-parity-harness-and-ledger.md), the engine contract in
  0351 (required `WireLimits`, typed errors), and native evidence in the
  [macOS queue](0363-mac-queue.md).
- `performance_claim`: every engine route is cheaper than its host route by
  32–51% in callgrind instructions (§5).

## 1. What landed

The text-storage transformations of the host's `text/**` are now
storage-scoped engines in `litchi-iwa-text-wire`. They are expressed over
the semantic values of `litchi-iwa-text`.

- An engine owns wire payloads only. It reads the storage, style, theme or
  object payloads it is given and returns a typed plan or rewritten
  payloads.
- The format package publishes them: it allocates identifiers, inserts and
  removes objects, and keeps archive metadata (object references, object
  UUIDs, external references, the identifier watermark).
- The leaf routes are added by W1-B, W1-C and W2. Here every engine is
  parity-tested through the host route it replaces. The leaf side is
  published by an independent test glue (`tests/parity/text_glue.rs`)
  through the public `litchi_iwa::Edit`.
- The host keeps every route and serves as the oracle. No host behaviour,
  manifest, crate edge or ops dispatch changed.

| Commit | Module | What |
|---|---|---|
| `c9f6c40604` | `paragraph_style` | Character and paragraph properties: `effective` and `plan` (`Change::{Set, Reset}` → `Plan::{Unchanged, Rewrite, Derive, Collapse}`), `uniform_style`, `repoint`, `stylesheet_with_variation`/`stylesheet_without_variation`. A private native wire model decodes with merge semantics and encodes canonically (no Prost in production). |
| `5d19cbd56e` | `ranged`, `smart_field` | The `ObjectAttributeTable` engine for the smart-field (11), bookmark (15) and highlight (23) tables: `entries`, and `rewrite` with `Add`, `Move` or `Remove`. Also the hyperlink and bookmark codecs. |
| `f330c64a57` | `language` | Language runs: `runs`, `at`, `rewrite` (`Set`, `RemoveBoundary`, `Reset`). |
| `d5389eed44` | `smart_field` | Date & Time codec: `date_time`, `encode_date_time`, `date_time_with_settings`. |
| `2ae86bad42` | `number_attachment` | `anchors`, `anchor`, `settings`, `encode`, `with_settings`. |
| `2d2474630d` | `drop_cap` | Plain-text Drop Caps: `placements`, `model`, `variation`, `plan` (`Rewrite`, `Derive`, `Detach`), `point`. |
| `24fb21ec67` | `list` | Paragraph lists: presets, placements, levels, numbering and per-level properties. Plans are `Rewrite`, `Derive` and `Collapse`, and boundaries can be rewritten. |
| `282ce7c2d2` | `paragraph_style::named` | Named paragraph styles: `stylesheet`, `theme_presets`, `catalog`, `applied`, `plan_apply` (`Reference` or `Replace` of an exclusive override chain), `named_copy`, `renamed`, `stylesheet_with_named`/`stylesheet_without_named`, `theme_with_preset`/`theme_without_preset`, `redefine`. Refusals are typed (`paragraph_style::Refusal`). |
| `29b8559209` | `annotation` | Highlights, text comments and replies: `annotations`, `annotation` and `graph` through an `Objects` lookup, with `TextHighlight`, `TextComment` and `TextCommentReply` views. Also the highlight and comment-storage codecs and the lossless body and reply rewrites. |

The test commits that pair with each engine are listed in §2.

Every public engine function takes a required `WireLimits` and is pinned
as a `required_signatures` entry of the `iwa-text-wire` glossary profile.
`docs/API_GLOSSARY.md` describes the rule.

## 2. Ledger

154 `text/**` rows are now `migrated`, all with destination
`shared`/`text.*`:

| Rows | Family | Parity tests | Test commit |
|---|---|---|---|
| 72 | character/paragraph properties (set/reset per property) | `text_paragraph_property_{set,rewrite,reset}_parity`, read parity | `d6973214e6` |
| 8 | hyperlinks, bookmarks | `text_smart_field_{add,update,remove,read,refusal}_parity` | `ed5ad65873` |
| 5 | languages | `text_language_{set,remove,reset,read}_parity` | `23448d023a` |
| 9 | Date & Time fields, `replace_text`/`set_text`, storage reads | `text_date_time_{add,insert,update,remove,read}_parity`, `text_replace_parity`, `text_storage_read_parity` | `f84def8f78` |
| 4 | number attachments | `text_number_attachment_{insert,update,remove,read}_parity` | `66d39f6ec9` |
| 4 | Drop Caps | `text_drop_cap_{set,rewrite,remove,read}_parity` | `7d5732ef72` |
| 32 | paragraph lists | `text_list_*_parity` | `b9b36d9e7f` |
| 8 | named paragraph styles | `text_named_paragraph_style_{apply,create,rename,delete,redefine,read}_parity` | `c653f0a412` |
| 12 | highlights, comments, replies | `text_{highlight,comment,comment_reply}_{add,edit}_parity`, `text_annotation_read_parity` | `08c9def6a0` |

None of these routes has host native evidence, so each has a W1-D row in
`0363-mac-queue.md` (`Evidence::Queued`).

Five rows are retired (`2962ac0669`):

- `text/date_time_field.rs from_object_id` is a raw native object-ID
  constructor, which ADR 0028 keeps off a supported facade. Date-time
  fields are addressed through the ranged-table entry the engine reads.
- The four `text/extractor.rs` rows are methods of the crate-private
  `TextExtractor` behind the host's document text read. They are not a
  storage transformation, and they leave with that route.

Four `text/editor.rs` rows stay `host-only`: `open`, `from_bytes`, `save`
and `to_bytes`. They are package I/O of `text::Draft`, which the leaf
`Package` of each format owns (family `io`).

## 3. Tiers and parity findings

- **Tier A, byte-identical.** Rewrites, collapses, moves, body updates and
  removals. This includes the removal of an exclusive override chain when a
  named style is applied, and the removal of the host's generated
  annotation author.
- **Tier B, known fields.** Every case that creates an object the host
  stamps with a random value: property variations and list styles (object
  metadata UUIDs), new named styles, highlights (text-attribute UUID), and
  comment storages (storage UUID and current time). The per-case
  `known_fields` mask exactly those fields, and `leaf_structural_metadata`
  accepts the glue's deterministic UUIDs.
- **Theme preset insertion.** The host re-encodes the whole theme
  canonically. A source-preserving reference-list rewrite instead appends
  the new preset after the text extension's last field. The engine
  therefore splices the preset in field-number order and replaces the text
  extension through the source-preserving `litchi-iwa-protos` theme codec.
  The output is byte-identical to the host's on all fixtures.
- **Theme codec limits.** The theme codec caps each resource at 64 MiB.
  The engine clamps the caller's `WireLimits` to that ceiling instead of
  passing the 512 MiB default through, which the codec would reject.
- **Annotation authors.** All four text fixtures carry an empty
  annotation-author storage. The first comment therefore creates the
  host's generated author (author storage entry, author object, external
  reference), and the last removal deletes it again. The glue reproduces
  both steps, and the tier A removal cases prove them.

## 4. Gaps and follow-ups

- **Leaf routes.** W1-B, W1-C and W2 add the per-format routes (Pages text
  boxes, Keynote slide text boxes, Numbers sheet text boxes). The
  `pages/`, `keynote/` and `numbers/` named-style, list, highlight and
  comment rows stay `host-only` until they do. Each of those routes
  publishes the plans the way `text_glue.rs` does.
- **Author storage.** The generated annotation author and its storage are
  shared comment infrastructure (`crate::comments`). They are also used by
  table-cell and drawable comments, so the leaf crates still need a common
  owner for them. The engine leaves author choice to the caller.
- **Overlapping highlights.** Storages with overlapping highlights
  (field 25) are refused (`ranged::Refusal::OverlappingHighlights`), as the
  host refuses them.
- **Native evidence.** Every W1-D route waits on its macOS queue row.

## 5. Measurements

The probes are in `crates/litchi-iwa/tests/parity/text_cost_probes.rs`:

- Build: release with `CARGO_PROFILE_RELEASE_DEBUG=1`.
- Fixture: `iwork/keynote/basic.key`.
- Pinning: `taskset -c 3`.
- Run: callgrind with `--toggle-collect` set to `*text_cost_probes::host_*`
  or `*text_cost_probes::leaf_*`.
- Order: each route runs host, leaf, leaf, host. A probe covers the whole
  route: open, edit and serialize.

| Route (calls) | Host Ir | Leaf Ir | Δ |
|---|---|---|---|
| paragraph property set / rewrite / reset (12) | 754.9 M | 510.2 M | −32.4% |
| hyperlink add (2) | 173.7 M | 106.1 M | −38.9% |
| language set (2) | 120.7 M | 59.8 M | −50.5% |
| named-style apply (removing an exclusive override) + redefine (4) | 858.4 M | 574.7 M | −33.1% |
| comment add + reply add (4) | 568.9 M | 342.1 M | −39.9% |

The engines themselves account for about 2% of the leaf cost. The rest is
package open, archive edit and serialization, which are shared by both
routes.

Wall-clock ABBA (`text_property_batch_wall_abba`) times sequential
alignment sets on one storage:

| Sets | Host | Leaf |
|---|---|---|
| 1 | 9.75 ms | 6.32 ms |
| 10 | 54.6 ms | 34.9 ms |
| 100 | 500 ms | 314 ms |

The 1k batch was not run because the host is too slow at that size.

Callgrind outputs are kept outside the repository, in
`.coord/iwa-w1d/cg/`. Rerun the commands in the probe file's header to
regenerate them.

| File | Bytes | SHA-256 |
|---|---|---|
| `engine-host.out` | 567,770 | `71af50a627137b7a83108bd91da8a452e01ceafbda9bdf06fd60e34d8418b088` |
| `engine-leaf.out` | 516,132 | `5f61b00bd827133ec2270b9ae95eb52523ba3113ad497e928c3969cff4022a39` |
| `named-host.out` | 626,652 | `cd0ee8dc5120f98cc8827a77cef14ef6ba9224d251ea4738742e19c524615796` |
| `named-leaf.out` | 571,842 | `5df8b02a747015a84c62bed0ff576d3d6bd62efcb73091d6ccd671ed550a9378` |

## 6. Gates

Final run on the branch with trunk merged (`b7278b5cb8`, a fast-forward
of the coordinator's merge), build env of the coordinator rules
(`CARGO_BUILD_JOBS=6`, `RUST_TEST_THREADS=8`).

| Command | Exit |
|---|---|
| `cargo fmt --all -- --check` | 0 |
| `cargo clippy -p litchi-iwa-text-wire -p litchi-iwa-text -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` | 0 |
| `cargo test -p litchi-iwa-text-wire -p litchi-iwa-text` (150) | 0 |
| `cargo test -p litchi-iwa` (1,950; `parity` target 157 passed, 7 ignored cost probes) | 0 |
| `cargo test -p litchi-ops --all-features`: `--lib` (1,022), `--doc` (217), nine groups of ≤20 `--test` targets (1,485) | 0 each |
| `cargo test -p litchi --all-features` (209) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps -p litchi-iwa-text-wire -p litchi-iwa-text -p litchi-iwa` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (71 profiles) | 0 |
| `python3 tools/check_crate_boundaries.py`; `python3 tools/non_iwork_gate.py verify` | 0; 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_iwa_exit_ledger tools.test_check_api_glossary` (1,347) | 0 |
| `python3 tools/iwa_exit_ledger.py --check` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `officecli_corpus.py --check` | 0, 0 |
| `python3 tools/unified_bench_matrix.py --check` (after regenerating `results/0289/bench-matrix.md` to list this record) | 0 |
| `python3 tools/check_feature_matrix.py` (31 legs) | 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `cd tools/perf-baseline && cargo check --locked --all-targets` | 0 |

`cargo test -p litchi` with default features fails seven `lib.rs`
doctests (`litchi::Error` is not in scope without format features). This
branch does not touch `crates/litchi` (no diff against trunk), so the
failure is pre-existing on trunk; with `--all-features` all 209 pass.
Both pre-commit size checks printed nothing for every commit.
