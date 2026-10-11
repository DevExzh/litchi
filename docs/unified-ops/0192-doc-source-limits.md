# 0192 — DOC source-limit vocabulary and defaults

## Change

DOC physical ingress uses `SourceLimits` and `SourceLimitKind`. The consuming
`SourceLimits::builder()` exposes checked `max_input_bytes`, `max_entry_bytes`,
and `max_aggregate_bytes` setters followed by `build()`. The dimensions mean
complete CFB input, one declared stream, and aggregate declared stream bytes.
The previous positional constructor and `with_max_*` variants are removed.

Construction failures use `Error::InvalidLimit { kind, value, maximum }`;
ingress failures use `Error::Limit { kind, observed, maximum }`. The old
`LimitsError`, `ResourceLimit`, and `ResourceKind` wrappers are removed.
The unified adapter preserves the native cause and maps invalid configuration
to `InvalidArgument`, and exceeded limits to `LimitExceeded`.

`OpenOptions` consumes already checked values through `source_limits`,
`password`, and `leniency`; its fields are private. The obsolete PAPX-padding
compatibility setter is removed. Immutable and raw packages expose retained
`source_limits`. The facade's credential-free `DocOptions` uses the same
`source_limits` spelling. Active security/performance harness consumers move
to the checked builder while retaining their explicit numerical policies.

Shared defaults now match OPC: 512 MiB complete input, 512 MiB per entry,
and 2 GiB aggregate. This intentionally widens DOC's former 128/64/96 MiB
defaults. Hard ceilings remain finite: 1 GiB input, 512 MiB entry, and 2 GiB
aggregate (the aggregate ceiling is raised from 768 MiB). Boundary tests use
small inputs and invalid configuration values; they do not allocate artifacts
near these ceilings. Applications can choose stricter validated profiles.

The canonical and raw ingress gates now preflight every declared CFB stream
from directory metadata before DOC payload reads. The sum includes metadata,
Equation Native, embedded objects, VBA, and opaque entries. Root mini-stream
backing is excluded to avoid charging the logical mini streams twice. This
intentionally broadens the old aggregate policy, which covered only
WordDocument, its selected table, and Data. Scanning sizes allocates no payload
buffers. CFB parsing and shared-owner adoption use the selected source input cap.

## Verification

The combined native DOC and DOC/DOCX/XLSX/XLSB/PPTX/encryption facade suite
passes 2,713 tests and doctests, with 14 ignored. Tests cover exact and one-less
input/entry/aggregate bounds, zero and over-ceiling configuration, inherited
profiles, default values, metadata/equation stream refusal on byte and adopted
OLE ingress, and typed facade dimensions/values. A gated reader disallows
payload reads after container parsing and proves the aggregate refusal happens
before any such read. Password, encryption, signature, transaction, batch,
save and read regressions remain green. The glossary verifier's 21 Python tests
pass. The OfficeCLI capability delta is unchanged.

Warning-denied Clippy (all targets, native DOC, facade, and root DOC feature)
and rustdoc pass. Both DOC glossary profiles pass 34 banned-name checks and
one transaction-family check. Crate boundaries pass for 66 packages and 268
internal dependency declarations, with 11 existing debt items. Scoped rustfmt
and `git diff --check` pass. The performance/security harness library and
binaries, standalone unified DOC paragraph harness, and native DOC fuzz targets
compile. This is a fuzz-target compile check, not a fuzz campaign.

Repository lockfiles are unchanged; checks use an external copied lockfile
with cached cc 1.6.0. No iWork crate is modified.

## Remaining scope

Already-parsed CFB adoption cannot retroactively limit directory/FAT
allocations made by its caller. This precondition is documented; generic
facade detection keeps its own format-neutral limits until it identifies DOC.
The CFB directory/allocation-table defaults remain separate from the three
DOC payload dimensions. The aggregate bounds declared logical payload sizes,
not repeated temporary copies or total process memory.

The focused body-text Snapshot path still opens its strict ObjectEditor
before the public-reader stream preflight. Its outer input gate runs first,
and its existing revision-owner limits remain active, but applying the DOC
entry/aggregate preflight before that strict owner's payload allocations is a
follow-up closed in [0193](0193-doc-strict-owner-admission.md). The primary Package ingress is preflighted before edits
can start. No complete allocation-budget guarantee for focused/raw escape
paths is claimed.

The migration does not establish full-goal completion. DOC focused/raw
vocabulary and semantic coverage, XLS/PPT write loops, tier 3, Python/CLI, and
all required performance/acceptance gates remain open. This configuration wave
does not claim a performance improvement or resolve the large-case timing
uncertainty from record 0191.
