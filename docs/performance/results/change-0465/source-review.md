# ODP checked-default review

The ODP lifecycle runner and fixture generator are unchanged. The code change
adds `OdpExistingAppendLifecycle` to the default list and describes its existing
synthetic corpus in matching Rust/Python catalog metadata. It retains the
64/4,096/8,192-slide, six-member source and the 64 KiB opaque member.

The existing gates cover semantic readback, exactly one appended slide, source
immutability, manifest bindings, untouched members, compressed opaque identity,
forward/inverse patch replay, stale-source refusal, and exact no-op. Input
cloning, append strings and sink construction precede timing; the timer includes
opening, transaction creation, append, commit and sequential output. Readback,
digests, preservation/patch oracles and destruction follow timing. This remains
an owned, fully materialized ordinary ODP lifecycle.

The representative coverage category can contain measured and correctness-only
rows, but must include a measured row to have measured status. Unsupported and
not-applicable rows cannot hide inside a measured category. Only measured rows
enter the report and timed-binding gates. XLSX/RTF streaming creation stays
correctness-only; synthetic ODP timing does not certify independent producers.

The base revision is `161cf53b20d8bb65fe79d4567b9ea7768430de7b`.
Both release build receipts bind the same 7,032-file Rust/TOML/lock manifest,
`0021cdee1035d1dc28ae1506fb991396d26b28e6c48ce4f1452c6be4bd5a263d`.
`source-code.json` archives the two changed Rust files; it is a selected source
archive, not the full source tree. Unchanged sources are recoverable from the
base revision and checked against the complete custody manifest.

The reports must truthfully retain dirty-worktree metadata. The generic
pairwise comparison policy still requires a clean worktree and distinct
revisions; those comparison gates are not weakened. This batch makes only a
descriptive baseline claim tied to source and binary identities, with no
before/after optimization comparison or speedup claim.
