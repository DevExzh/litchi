# 0391 — W1-G: shared whole-cell text-style substrate

- Status: paused Numbers-first substrate under explicit budget/admission
  repair. Pages and Keynote facades are not started, no ledger rows are
  migrated, and no parity/performance/native acceptance claim is active. The
  current clean tip is d5abe8856a.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`, then synchronized with frozen
  trunk 24ecfb543748885192b8fb5273abb54e669764c6.
- Continuation commits: trunk merge `d220cd76ed`, budget admission
  `beb878ea23`, codec-allocation pre-admission `36be343a42`, residual
  closure `3bfb0c6319`, and proof checkpoint `d5abe8856a`.
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

The current smoke gates pass: 42 shared table unit tests (including duplicate,
unknown-field and low-field/work/allocation pre-admission cases), six Numbers
`cell_text_style` tests, formatting, scoped check, and strict Clippy for both
crates. These are not the required 17-property host differential or full
transaction proof.

The current Numbers proof rerun also passes the existing 17-property host
set/reset/read vectors and the populated-reader differential: 595 direct
semantic reads compare host and leaf values. The shared wire controls cover
duplicate singular fields, unknown-field retention/locality, zero ceilings,
and low field/work/allocation refusals; the Numbers integration adds an exact
source/no-op/inverse case. The remaining explicit proof gap is a multi-cell
release from a pre-existing shared explicit style key. The tracked Numbers
fixtures do not contain such a key, while the current parity batch creates one
variation per cell, so this route is not claimed until a valid source-built
shared-key fixture or equivalent bounded test input is supplied.

Before Pages/Keynote work begins, requalify the shared substrate against the
full Numbers 17-property parity/readback and release-mode facade matrix,
including malformed/duplicate references, unknown locality, low-limit refusal
before wire work, inverse/no-op and candidate reopen. Only after that
acceptance should the six Pages/Keynote facade rows and native queue entries
be added. Timing remains held until the coordinator grants the slot.

Historical parity and corpus receipts from the pre-repair checkpoint are
retained as context only and are not requalified evidence for this source.
The next gate must add a valid shared-key batch vector, package-level
malformed/duplicate and locality controls over the populated source, cumulative
depth/zero-limit transaction evidence, and candidate reopen before any
Pages/Keynote facade or ledger work.
