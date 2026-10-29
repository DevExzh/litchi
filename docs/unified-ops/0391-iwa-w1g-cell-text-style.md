# 0391 — W1-G: shared whole-cell text-style substrate

- Status: in progress; Numbers-first substrate and proof repair are complete
  on the isolated WP, pending coordinator/W2-A independent review. Pages and
  Keynote facades are not started, no ledger rows are migrated, and no
  performance or native acceptance claim is active. The current source tip is
  `126b1dd7ce`.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk 24ecfb543748885192b8fb5273abb54e669764c6.
- Continuation commits: trunk merge `d220cd76ed`, budget admission
  `beb878ea23`, codec-allocation pre-admission `36be343a42`, residual
  closure `3bfb0c6319`, proof checkpoint `d5abe8856a`, and final residual
  preflight/proof repair `62e6b27aae` plus parent-reference accounting
  `126b1dd7ce`.
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
reference traversals are charged on every pass and typed decoder failures are
terminal. The Numbers adapter remains shared and maps `BatchUnsupported` to
typed `UnsupportedDependency` without a legacy fallback.

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

The current scoped gates pass: 45 shared table unit tests (including
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
checks the combined list rewrite removes the shared key after both releases.

Before Pages/Keynote work begins, the coordinator must independently review
the nested budget/admission contract and run final integrated qualification.
No wall-clock or instruction-cost result is claimed from this slice; timing
remains held until the coordinator grants the slot. Only after that acceptance
should the six Pages/Keynote facade rows and native queue entries be added.

Historical parity and corpus receipts from the pre-repair checkpoint are
retained as context only and are not requalified evidence for this source.
Native evidence, ledger migration, Pages/Keynote facade work and performance
remain coordinator-owned follow-ups.
