# 0391 — W1-G: shared whole-cell text-style substrate

- Status: paused Numbers-first substrate under explicit budget/admission
  repair. Pages and Keynote facades are not started, no ledger rows are
  migrated, and no parity/performance/native acceptance claim is active. The
  current clean tip is 36be343a42.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk 24ecfb543748885192b8fb5273abb54e669764c6.
- Continuation commits: trunk merge `d220cd76ed`, budget admission
  `beb878ea23`, and codec-allocation pre-admission `36be343a42`.
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
fixed constant. Zero object/reference ceilings are preserved. This is a
correctness checkpoint, not substrate acceptance; supported-profile,
malformed/duplicate, locality and host differential proof remain pending.

Numbers currently exercises the shared planner for the existing smoke routes.
The adapter maps `BatchUnsupported` to typed `UnsupportedDependency` and
does not fall through to a legacy publisher. Full supported-profile and batch
proof remains pending.

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

The current smoke gates pass: 40 shared table unit tests (including three new
low-field/work/allocation pre-admission cases), five Numbers
`cell_text_style` tests, formatting, scoped check, and strict Clippy for both
crates. These are not the required 17-property host differential or full
transaction proof.

Before Pages/Keynote work begins, requalify the shared substrate against the
full Numbers 17-property parity/readback and release-mode facade matrix,
including malformed/duplicate references, unknown locality, low-limit refusal
before wire work, inverse/no-op and candidate reopen. Only after that
acceptance should the six Pages/Keynote facade rows and native queue entries
be added. Timing remains held until the coordinator grants the slot.

Historical parity and corpus receipts from the pre-repair checkpoint are
retained as context only and are not requalified evidence for this source.
The next gate must add direct host-vs-shared reads for all 17 properties, an
actual shared-key batch vector, unknown/locality and malformed-duplicate
controls, cumulative/depth/zero-limit transactions, inverse/no-op and
candidate reopen before any Pages/Keynote facade or ledger work.
