# 0745 log sections (ready to paste)

## For HOTSPOTS.md

## 0745 — PPT slide-order digests deferred; editor opens read owned streams once

[0745](0745-ppt-lazy-artifact-digests.md) moves the two whole-artifact SHA-256
digests out of PPT slide-order commit into `Patch::to_durable`. They are
memoized per snapshot, computed from exactly the retained bytes, and the
durable bytes are unchanged (pinned goldens from the base). Only a mixed
formatting-then-structural commit still hashes once, because it binds an
intermediate artifact that the patch does not retain. Two further duplicates
are removed:

- PPT editor opens no longer read the Document and Current User streams twice.
- Commit reuses its publishing editor for the before-payload capture.

Remaining PPT lifecycle cost is editor opens, finish and reopen. The editor
that `edit()` discards is a validation and is kept. Next candidates found:
`litchi-ole-common` `render_copy_through` opens in-memory sources generically,
so the CFB overlay hashes source and target on several passes. PPT timings on
this host swing up to about ±15% with the early heap layout, which the
length of argv[0] shifts through each binary's own `std::env::args()` copy.
Page faults vary with it, and the glibc mechanism is inferred, not tested.
Randomize that layout in paired measurements. The non-iWork goal remains active.

## For REPORT.md

## 0745 — lazy PPT artifact digests retained

[0745](0745-ppt-lazy-artifact-digests.md) defers PPT slide-order artifact
digests to durable serialization and removes duplicate editor-stream reads and
a duplicate commit editor open. Over 18 paired heap layouts (648 processes,
core 16), the public 45543.ppt slide-removal lifecycle changes:

| Measurement | Deferred digests | Both commits |
|---|---:|---:|
| Median p50 | −33.8% | −39.5% |
| Sealed 0734 probe | −36.2% | −42.7% |
| 41246-1.ppt | −19.0% | −20.9% |

Other lifecycles with both commits:

- The no-op commit changes by −85.9% and formatting-only hide by −52.3%.
- The harness PPT semantic selectors change by −10.1% to −67.0%.

Commit plus `to_durable` is +0.1% after the first commit and −5.8% after both.
Instructions show the work moved, not vanished. DOC controls stay within −1.3%
to +0.1%.

- Durable JSON, outputs and refusals are byte-identical across 32 golden
  scenario entries.
- The first commit alone raises the p95 of commit plus `to_durable` by +8.6%;
  both commits together are −5.9%.
- Flags from layout-sensitive single layouts reach +18.3%.

Gates:

- `litchi-ppt`: 1,229 tests plain and 1,235 with diagnostics.
- Facade: 382 tests.
- Clippy, rustdoc, boundaries and claims checks pass.

`performance_claim: none`. [Evidence](results/change-0745/README.md).

## For GOAL_AUDIT.md

## 0745 — deferred PPT artifact digests and single-read editor streams

[0745](0745-ppt-lazy-artifact-digests.md) keeps the ADR 0003 durable patch
contract byte for byte: tested unchanged in eight unit scenarios pinned to
the eager base, and in 32 golden entries, including refusals, across three
builds. A review fix makes the durable artifact-conflict test reach that check
in both directions. The change retains no additional artifact. Each snapshot
carries a small, bounded digest memo: 48 bytes, plus a 64-byte hex string once
hashed. The bytes and memo are bound by construction, and in-memory apply
still authorizes by exact bytes.

The feature-gated `DiagnosticPhase` loses its two artifact-hash phases and
gains `IntermediateArtifactHash`. That is a breaking change under 0652
trade-off 1.

The survey found no other eager durable-only digest in the OLE2 or OOXML
crates, and lists the authorization and redundant-check sites it left
unchanged. Heap-layout randomization is now the measurement method for PPT
lifecycles. No coverage, claim or timing-contract promotion; the non-iWork
goal remains open.
