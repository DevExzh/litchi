# 0391 — W1-G: shared whole-cell text-style substrate

- Status: Numbers-first substrate and proof repair are complete on the
  isolated WP. Pages and Keynote facades are not started, no ledger rows are
  migrated, and no native acceptance claim is active. The current code/probe
  tip is `4dca2d60b6`; release-mode instruction capture remains coordinator-
  gated and has not started.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk 24ecfb543748885192b8fb5273abb54e669764c6.
- Continuation commits: trunk merge `d220cd76ed`, budget admission
  `beb878ea23`, codec-allocation pre-admission `36be343a42`, residual
  closure `3bfb0c6319`, proof checkpoint `d5abe8856a`, and final residual
  preflight/proof repair `62e6b27aae` plus parent-reference accounting
  `126b1dd7ce`, repeated-parser accounting `3a8313257c`, terminal-error plus
  exclusive-path repair `f616f4315e`, and mutation-output admission
  `a6f1723ac4` plus bounded wire rewrite/batch closure `2e4a784a99` and
  identity/unknown-preservation proof `ef2187916d`, and paragraph-field
  precondition proof `daed5d99c3`.
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

- W1-G owns `crates/litchi-iwa-table/src/cell_text_style.rs`, its crate export,
  shared planner tests, and the Numbers adapter in
  `crates/litchi-numbers/src/package/table_cell_text_style.rs`.
- Existing helper consumers (`table_cell_text`, `table_cell_list`,
  `storage_list`, `table_resize`, and conditional highlights) remain on the
  compatibility helpers until their common graph primitives are extracted in
  the next bounded step.
- Pages/Keynote facades, new ledger rows, native queue changes, E-owned style
  chart/font paths, and F-owned formula-owner admission are out of scope.

## Verification and remaining work

The current scoped gates pass: 47 shared table unit tests (including
duplicate/unknown locality, nested-font pre-admission, residual BNC fields,
zero ceilings and low field/work/allocation cases), seven Numbers
`cell_text_style` tests, formatting, strict Clippy for both crates, and the
affected fuzz locks. `python3 tools/test_no_3rdparty_references.py`,
`python3 tools/iwa_exit_ledger.py --check`, and
  `python3 tools/check_crate_boundaries.py`, the API glossary tests, and strict
  rustdoc for the two affected crates also pass. The litchi fuzz lock
records its existing `litchi-iwa-text-wire` dependency in this source commit;
the generated fuzz targets are kept outside the checkout.

The current Numbers proof rerun passes the existing 17-property host
set/reset/read vectors and the populated-reader differential: 595 direct
semantic reads compare host and leaf values. The two focused parity cases pass
on the current source. A source-built shared-key fixture now drives the actual
Numbers publisher: two cells share one style key, both are changed, the
candidate is reopened and read directly, the old key is released, an
unrelated cell remains unchanged, inverse bytes equal the source, and an
effective-style no-op preserves candidate bytes. The shared planner test also
checks the combined list rewrite removes the shared key after both releases;
its ownership control confirms an incoming model reference blocks in-place
rewrite. Batch publication uses fresh derivations while repeated single-cell
updates reuse the existing style object with its actual parent chain.
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
selected cells/style graph source-relative. Metadata output admission now
pre-charges bounded registry growth for the batch instead of relying on
fixture-sized quota padding. Each route performs host/leaf reopen/readback,
locality, refcount/identity, and inverse checks before its exact
host/leaf/leaf/host wrapper calls. The batch route now prepares the shared
table/tile/style graph once, batches tile-key publication and batches
readback preparation; it completed the 1,000-cell correctness route on this
source. The final explicit run on this source was:

```text
cargo test --offline -p litchi-iwa --test parity \
  cost_probes::numbers_whole_cell_style_batch_1000_cost_abba \
  -- --ignored --nocapture --test-threads=1
```

It passed in 416.09 s, including setup/correctness and the four ABBA calls;
this is a correctness/probe-readiness receipt only, not a Callgrind or
release-mode performance qualification. Immutable binaries, captures and
instruction-count receipts are intentionally absent until the coordinator
accepts this fixture/probe checkpoint and grants the CPU slot.
