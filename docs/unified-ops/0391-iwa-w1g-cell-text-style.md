# 0391 — W1-G: shared whole-cell text-style substrate

- Status: Numbers-first substrate, bounded tile writer, and header-bucket
  materialization repair remain isolated on the WP. Pages and Keynote facades
  are not started, no ledger rows are migrated, and no native acceptance claim
  is active. The current code tip is `7b8cf978fb`; release-mode instruction
  capture remains coordinator-gated and has not started.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk `e0e297a82b0a534dfd5611ced3a9811851edbf76`.
- Continuation commits: the current WP includes F's reviewed metadata
  transition `b9337962b9` through merge `8662fa8ecd`, staged-budget/readback
  repair `76cfbf3afb`, current frozen-root merge `2eb913a401`, and prepared
  multi-stylesheet/segmented-list plus publication-admission closure
  `74651bb4ba`, planner/cache/probe repair `12f9de7115`, and verification
  budget plus mixed-no-op proof `ea9a17e04d`, direct watermark/cardinality
  proof `ee47ab4896`, reviewed shared tile-writer merge `7f69b0710a`, and
  Numbers writer adoption `43608b8faa`, bounded shared table-fact selection
  `4f014f7463`, reviewed header-helper merge `ea1bab6270`, and header-adapter
  integration `7b8cf978fb`. Earlier budget, parser,
  exclusive-path, mutation-output and identity/unknown-preservation
  checkpoints remain in history.
- Scope: shared `litchi-iwa-table::cell_text_style` graph reads and direct
  `CharacterProperties` planning, plus the Numbers read and single-cell
  adapter path. Existing public Numbers selectors and vocabulary remain
  unchanged.

## Shared contract

`cell_text_style` owns the format-neutral `Target`, bounded source/table/BNC
inspection, style-list and parent-chain resolution, effective character
properties, inherited-value reduction, exclusive-versus-derived variation
planning, stylesheet/list payload rewrites, fresh-object descriptions,
metadata registration requests, and typed wire/limit/allocation failures.
The plan contains neutral rewrites, insertions, registration operations and
cell-key updates; the leaf supplies identifiers and owns `Working`, metadata
publication, exact ZIP reassembly, candidate reopen/readback and public error
mapping.

The current continuation repair makes the budget contract explicit:
`Limits::from_source` derives wire fields, nesting, input/output and work
ceilings from the caller's inner archive profile; raw wire structure is
preflighted before `WireView` allocation; cumulative fields, references,
work, collection allocations and retained copies are admitted through the
operation ledger; and style depth follows the caller profile rather than a
fixed constant. Zero object/reference ceilings are preserved. Character
property and parent decoders now have separate source-bound nested preflights,
including span allocations, aggregate nested work/input/fields, output/font
buffers, and inherited-font copies. BNC style-key reads charge residual
input/fields/work before `BncCell::parse`; exclusive-style metadata and
reference traversals are charged on every pass. Raw-wire scans, `WireView`
passes, and the private text-wire decoder views each debit their own
input/fields/work and span allocations; a cumulative regression fixes the
boundary at the third parser pass. Typed decoder failures remain terminal.
The Numbers adapter remains shared and maps `BatchUnsupported` to typed
`UnsupportedDependency` without a legacy fallback.
Batch/list varint failures and character-style encoder resource failures retain
typed terminal errors. Single-cell exclusive rewrites skip legitimate
stylesheet owners, retain the actual parent chain, and use exact flag
semantics for unmentioned properties; batch planning keeps this path
single-cell-only and requests fresh identifiers for batch publication.
Mutation writers now pre-admit bounded output bytes, retained copies,
allocations, and writer parse/work passes across repeated batch merges; the
shared planner has an exact cumulative-output regression. In-place reuse is
restricted to the canonical generated variation wire shape, so unknown or
unrelated paragraph fields force safe fresh derivation rather than loss.

This is still a Numbers-only substrate slice: Pages/Keynote, ledger rows,
native queue changes, timing and release-mode performance qualification remain
out of scope.

## Ownership

- W1-G owns `crates/litchi-iwa-table/src/cell_text_style.rs`, the shared
  `cell_text_style_headers` and tile-writer exports, shared planner tests, the
  Numbers adapter in `crates/litchi-numbers/src/package/table_cell_text_style.rs`,
  and its focused parity/cost controls.
- Existing helper consumers (`table_cell_text`, `table_cell_list`,
  `storage_list`, `table_resize`, and conditional highlights) remain on the
  compatibility helpers until their common graph primitives are extracted in
  the next bounded step.
- Pages/Keynote facades, new ledger rows, native queue changes, E-owned style
  chart/font paths, and F-owned formula-owner admission are out of scope.

## Verification and remaining work

The current focused checks pass: 60 shared table unit tests (including
duplicate/unknown locality, nested-font pre-admission, residual BNC fields,
zero ceilings and low field/work/allocation cases), nine Numbers
`cell_text_style` tests, formatting, and strict Clippy for both crates. The
workspace-level glossary, boundary, ledger and fuzz gates remain coordinator
checks on the integrated candidate. The broader
`test_no_3rdparty_references`, ledger, boundary, glossary, rustdoc and
fuzz-lock results from the prior checkpoint are not reissued as current-source
qualification here; the coordinator must rerun them on the integrated
candidate. The generated fuzz targets remain outside the checkout.

The prior Numbers proof rerun covered the 17-property host set/reset/read
vectors and the populated-reader differential (595 direct semantic reads);
those receipts are historical until rerun on the integrated candidate. The
current nine-case Numbers suite includes the shared-key publisher path and
the mixed no-op/changed batch control: two cells share one style key, both are
changed, the
candidate is reopened and read directly, the old key is released, an
unrelated cell remains unchanged, inverse bytes equal the source, and an
effective-style no-op preserves candidate bytes. The shared planner test also
checks the combined list rewrite removes the shared key after both releases;
its ownership control confirms an incoming model reference blocks in-place
rewrite. Batch publication uses fresh derivations while repeated single-cell
updates reuse the existing style object with its actual parent chain.
All Numbers cell-key publication now routes through the shared bounded tile
writer, including existing cells, empty slots and missing rows; the former
default-limit/private writer paths are removed.
The header adapter parses the DataStore row-header envelope into source-indexed
bucket references, selects the checked global row bucket, and publishes row
and column count deltas only when a previously absent cell is materialized with
a style key. Existing cells and no-op updates leave both buckets byte-exact;
missing bucket references do not trigger a fallback writer. The focused
`cost_probes::numbers_whole_cell_style_materialization_header_controls` test
decorates header buckets with a retained unknown field and exercises empty G22
plus the first missing row in an existing tile. It compares both bucket counts
and unknown fields, tile row/cell presence, host and leaf semantic styles, and
the exact leaf inverse. The current-source log is retained at
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-header-7b8cf978fb/header-controls.out`.
The package proof mutates an otherwise exclusive style with unknown root and
unrelated paragraph data (asserting both the nonempty paragraph field and root
unknown field); the canonical-shape gate derives a fresh style and the old
object bytes survive exact inverse/readback.

Before Pages/Keynote work begins, the coordinator must independently review
the nested budget/admission contract and run final integrated qualification.
No wall-clock or instruction-cost result is claimed from this slice; timing
remains held until the coordinator grants the slot. Only after that acceptance
should the six Pages/Keynote facade rows and native queue entries be added.

Historical parity and corpus receipts from the pre-repair checkpoint are
retained as context only and are not requalified evidence for this source.
Native evidence, ledger migration, Pages/Keynote facade work and performance
remain coordinator-owned follow-ups.

## Whole-cell cost-route preparation

The current source contains three dedicated, non-inline ignored routes in
`crates/litchi-iwa/tests/parity/cost_probes.rs`:

- `numbers_whole_cell_style_single_cost_abba`: one B2 whole-cell set;
- `numbers_whole_cell_style_exclusive_repeat_cost_abba`: a source-authored
  exclusive style followed by a repeated single-cell update;
- `numbers_whole_cell_style_batch_1000_cost_abba`: a source-built 50×20 table
  with 1,000 materialized cells sharing one style key, one batch leaf edit,
  and a host loop over the same cells.

The source builder is outside the measured wrappers. It retains a valid
Numbers registry, uses the modern tile storage fields, and leaves the
selected cells/style graph source-relative without quota-padding fixture
members. The reviewed F metadata transition is present in this WP; final
integrated qualification still owns the combined metadata/publication budget.
Each route performs host/leaf reopen/readback,
locality, refcount/identity, and inverse checks before its exact
host/leaf/leaf/host wrapper calls. The batch route now prepares the shared
table/tile/style graph once, batches tile-key publication and batches
readback preparation. A temporary current-source leaf-only checkpoint passed
the unpadded 1,000-cell route after `74651bb4ba`; the host loop has not been
rerun after these repairs. The last explicit full route on this source was:

```text
cargo test --offline -p litchi-iwa --test parity \
  cost_probes::numbers_whole_cell_style_batch_1000_cost_abba \
  -- --ignored --nocapture --test-threads=1
```

The historical run passed in 416.09 s, including setup/correctness and the
four ABBA calls, but predates the current budget/allocator repairs and is not
current-source qualification. Immutable binaries, captures and
instruction-count receipts are intentionally absent until the coordinator
accepts this fixture/probe checkpoint and grants the CPU slot.
