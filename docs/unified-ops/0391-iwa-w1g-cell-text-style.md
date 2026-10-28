# 0391 — W1-G: shared whole-cell text-style substrate

- Status: Numbers-first substrate landing; Pages and Keynote facades are not
  started and no ledger rows are migrated.
- Source branch: `uo2/iwa-w1g-cell-text-style`, cut from frozen root
  `1e3978082bd8e576f0afeb7e3eebac1ca3908d86`.
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

The planner preserves the current supported profile: it does not blanket
refuse shared or segmented graphs. Root and segment style-list entries are
read with ownership retained; unsupported or ambiguous graph shapes refuse
typed. Zero ceilings are not raised, and the shared parser does not expose
native identifiers through supported APIs.

Numbers now routes whole-cell style reads and single-cell writes through the
shared planner. Multi-cell staged edits temporarily retain the legacy
compatibility publisher so their existing one-variation-per-cell, atomic
staging and exact inverse behavior remain unchanged while the shared batch
plan is completed. This is an explicit bounded compatibility seam, not a
public alias or a generic-property rewrite.

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

The Numbers `cell_text_style` test target passes all four existing tests,
including the created-workbook transaction/inverse/readback cases and the
33-cell corpus (33 styled, no refusals). Shared planner unit tests cover
minimal inherited-property reduction and host-default emphasis semantics.
Formatting and strict Clippy for `litchi-iwa-table` and `litchi-numbers` pass.

Before Pages/Keynote work begins, complete and requalify the shared batch
planner for existing-key releases and multi-cell staging, then rerun the
full Numbers 17-property parity/cost matrix. Only after that acceptance should
the six Pages/Keynote facade rows and native queue entries be added.
