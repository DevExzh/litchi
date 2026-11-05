# 0391 — W1-G: shared whole-cell text-style substrate

- Status: Numbers-first shared substrate, bounded tile writer, and
  header-bucket materialization repair are complete and accepted as a scoped
  source-bound slice at `c9f74916c7`. The three dedicated whole-cell Numbers
  instruction routes passed the accepted Callgrind qualification. Pages and
  Keynote facades are not started, no ledger rows are migrated, and no native
  acceptance claim is active. Final integrated-trunk requalification remains
  coordinator-owned.
- Handoff tip: current trunk `bc9a37fd75` (the sparse-example relocation) is
  merged into this WP as `c52439a057`; the runtime qualification and external
  performance receipts remain explicitly bound to clean source `c9f74916c7`.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk `e0e297a82b0a534dfd5611ced3a9811851edbf76`; the scoped evidence is
  bound to clean source `c9f74916c73772bc57b9218d99220170ddb7504d`.
- Continuation commits: the current WP includes F's reviewed metadata
  transition `b9337962b9` through merge `8662fa8ecd`, staged-budget/readback
  repair `76cfbf3afb`, current frozen-root merge `2eb913a401`, and prepared
  multi-stylesheet/segmented-list plus publication-admission closure
  `74651bb4ba`, planner/cache/probe repair `12f9de7115`, and verification
  budget plus mixed-no-op proof `ea9a17e04d`, direct watermark/cardinality
  proof `ee47ab4896`, reviewed shared tile-writer merge `7f69b0710a`, and
  Numbers writer adoption `43608b8faa`, bounded shared table-fact selection
  `4f014f7463`, reviewed header-helper merge `ea1bab6270`, and header-adapter
  integration `7b8cf978fb`, row-header admission closure `d39eb1740c`, and
  cross-bucket/missing-ownership closure `6d5f02c0a1`, and selection-scan
  admission closure `6ec4e6becf`, regression arithmetic correction
  `33731d7056`, and final source-bound documentation/probe binding
  `c9f74916c7`. Earlier budget, parser,
  exclusive-path, mutation-output and identity/unknown-preservation
  checkpoints remain in history.
- Scope: shared `litchi-iwa-table::cell_text_style` graph reads and direct
  `CharacterProperties` planning, plus the Numbers read, single-cell and batch
  adapter paths. Existing public Numbers selectors and vocabulary remain
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

This remains a Numbers-only substrate slice: Pages/Keynote facades, ledger
rows, native queue changes, and unified `File` qualification remain out of
scope. The accepted instruction qualification below measures the three
dedicated whole-cell Numbers routes only; it does not qualify the integrated
trunk or native application acceptance.

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

The current focused checks pass: 62 shared table unit tests (including
duplicate/unknown locality, nested-font pre-admission, residual BNC fields,
zero ceilings and low field/work/allocation cases), nine Numbers
`cell_text_style` tests, two Numbers header-policy tests, formatting, and
strict Clippy for both crates. The current-source logs are retained outside
the checkout under `/home/zhuhe/code/litchi-worktrees/.coord/`.
Workspace-level glossary, boundary, ledger, rustdoc and fuzz-lock gates remain
coordinator checks on the integrated candidate; they are not inferred from
this scoped record. The generated fuzz targets remain outside the checkout.

After the current-trunk merge, `cargo fmt --all -- --check`, strict Clippy for
`litchi-iwa-table` and `litchi-numbers`, the API glossary `numbers-minimal`
and `numbers-full` profiles, the exit-ledger check, crate-boundary check (70
packages, 314 internal declarations, 12 explicit debt items), example-target
check, no-3rdparty suite (75 tests), and the benchmark-matrix check all pass.
All 20 locked fuzz workspaces also pass in the external capped-target run
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-final-c9f74916c7/fuzz-locks-c52439a057.log`
(SHA-256 `e410d50453aecc7cb1e2aad97aa9c405b111f0f22e05145363c22e0b1c07c287ad`).

The Numbers compatibility proof covers the 17-property host set/reset/read
vectors and the populated-reader differential (595 direct semantic reads).
Those 595 reads are correctness/read coverage, not an instruction-cost route
and are kept separate from the three whole-cell measurements below. The
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
`/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-header-6d5f02c0a1/header-controls.out`.
The shared header delta helper also partitions rows that cross the 65,536-row
bucket boundary while retaining global row indices; the adapter rewrites each
owned row bucket and publishes the column bucket once. A source that requires
a materialization count but lacks the corresponding row or column ownership
reference is refused with `UnsupportedDependency`; it is not silently
published through a tile-only fallback. The boundary and missing-ownership
controls pass in the shared header tests and the Numbers adapter unit suite.
The caller-side model, update, tile-index, tile-group, and row-bucket ownership
comparisons now pre-admit their conservative cumulative work. The shared
`table_fact_tile_duplicate_scan_is_cumulatively_admitted` regression exhausts
the work ledger immediately before the tile duplicate scan and expects the
typed work refusal; it passes in the current clean-source focused run.
The package proof mutates an otherwise exclusive style with unknown root and
unrelated paragraph data (asserting both the nonempty paragraph field and root
unknown field); the canonical-shape gate derives a fresh style and the old
object bytes survive exact inverse/readback.

The coordinator accepted the nested budget/admission contract and the scoped
source-bound instruction qualification. On pinned CPU 7, the immutable
release-debug ELF was measured with two selected wrapper calls per side and
the c063 fail-closed raw audit. The accepted inclusive leaf/host ratios are:

- single set: `156,915,444 / 282,826,408 = 0.5548118548`;
- repeated exclusive update: `136,936,777 / 205,615,736 = 0.6659839352`;
- 50×20 shared-key batch: `3,832,042,990 / 461,133,960,244 = 0.008310042895`.

All three are within the `1.05` instruction-cost bar. The batch annotation
also retains its exact partial DWARF file-context attribution row
(`3,827,066,960`) while the inclusive PROGRAM TOTALS remain the qualified
`3,832,042,990`; no normalization or validator change was made. The result
is a scoped Numbers whole-cell qualification, not a wall-clock claim, unified
`File` qualification, integrated-trunk qualification, or native acceptance.
The coordinator will requalify after merge, and Pages/Keynote facade rows and
native queue entries remain later work.

Historical parity and corpus receipts from the pre-repair checkpoint are
retained as context only. The current scoped evidence is external and bound
to clean source `c9f74916c73772bc57b9218d99220170ddb7504d`:

- normalized three-route manifest:
  `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-final-c9f74916c7/measure-6cb1a8fcb7b86dbe9f46e86e535edd503f988b212e02490507c8ec0774294190/callgrind/manifest-3routes-c9f74916c7.json`
  (SHA-256 `f2e6364501648516247f21532141474643296f2098a9194bd51d7008d253f14c`);
- normalized results:
  `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-final-c9f74916c7/measure-6cb1a8fcb7b86dbe9f46e86e535edd503f988b212e02490507c8ec0774294190/callgrind/results-3routes-c9f74916c7.json`
  (SHA-256 `7e6975b99f292de8b9055c848c033dbbac8a792bf368bbf79d10b459c80cf039`);
- c063 derived audit:
  `/home/zhuhe/code/litchi-worktrees/.coord/iwa-w1g-final-c9f74916c7/measure-6cb1a8fcb7b86dbe9f46e86e535edd503f988b212e02490507c8ec0774294190/callgrind/derived-3routes-c9f74916c7.json`
  (SHA-256 `52662a4a1d72632c88d0f819c826737c337c3c3eb7f68e80a0ffb7eb27e1999c`);
- immutable release-debug ELF: 214,625,544 bytes, mode `0555`, SHA-256
  `6cb1a8fcb7b86dbe9f46e86e535edd503f988b212e02490507c8ec0774294190`.

Native evidence, ledger migration, Pages/Keynote facades and integrated
requalification remain coordinator-owned follow-ups.

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
members. Each route performs host/leaf reopen/readback, locality,
refcount/identity, and inverse checks before its exact host/leaf/leaf/host
wrapper calls. The batch route prepares the shared table/tile/style graph
once, batches tile-key publication, and batches readback preparation. The
three routes were run from the immutable archived ELF on clean source
`c9f74916c7`; their exact raw capture and test bindings are in the external
manifest/results above.

```text
cargo test --offline -p litchi-iwa --test parity \
  cost_probes::numbers_whole_cell_style_batch_1000_cost_abba \
  -- --ignored --nocapture --test-threads=1
```

The earlier 416.09-second run is retained only as historical context and is
not used for the accepted numbers above. Immutable binaries, captures and
instruction-count receipts remain outside the repository; the final
integrated candidate must be requalified after coordinator merge.
