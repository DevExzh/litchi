# 0359e — Typed OLE object-capture budgets

- Date: 2026-10-21.
- Base: `7d5de8000a`.
- Status: accepted core capture slice; remaining semantic work and integrated gate below stay required.
- Contract: 0351 error and resource rules; required continuation of 0359c.

The shared object capture, discovery, edit and publication surfaces return
`litchi_ole_common::Error`. Configured object limits report
`Limit { kind, observed, maximum }`; zero configurations report
`InvalidLimit { kind, value, maximum }`. `LimitKind` identifies object and
stream counts, storage/selector depth, directory traversal, stream/object/total
bytes, storage counts and removal batches. Low-level CFB failures retain their
source. Directory visitor exhaustion under an object budget is normalized to
`LimitKind::DirectoryEntries` at the conversion boundary.

Package, presentation and native-stream patch base mismatches use
`PatchConflict`, with a fixed content-free message. DOC and XLS preserve the
shared error in their own `Error::Object` variants. All public DOC callback
bounds, internal callbacks, examples, tests and unified classifiers migrate
in this change. A public callback regression
checks that a typed rejection after modifying its candidate leaves the
original bytes unchanged. Existing CFB/Graph/PPT interfaces that do not use the shared object layer keep their
own errors.

The DOC-to-core conversion previously had no typed conflict or invalid-limit
variant available. `litchi_core::Error` now preserves those cases with
`PatchConflict` and `InvalidLimit`; configured usage still uses
`ResourceLimit`, with the corresponding byte/count/depth resource dimension.
A shared CFB source-change error converts to an I/O error through this path;
its invalid physical policy converts to the typed core invalid-limit variant.
Both direct host and core conversions are checked by unified error tests:
object budgets report `limit_exceeded`, invalid policies `invalid_argument`,
patch base mismatches `invalid_input`, and source changes `io_error`.
Malformed CFB remains a CFB parse failure. The generic sheet read abstraction
keeps its existing content-free refusal for configured XLS backend limits and
uses an invalid-value refusal for an invalid object policy. Nested I/O and
source changes retain their sheet categories. The sheet error gains a distinct
`PatchConflict` so replay conflicts are not mislabeled as malformed data or
invalid values. Native and unified errors retain the precise budget dimension
and numeric counts.

Review also found a depth-admission defect in selected-object discovery. The
policy defines depth from the CFB root, but selected subtree capture checked
only depth below the selected target. Capture now includes the target's
absolute depth. A regression test exercises a selected storage at the ceiling
with a nested child and verifies rejection before retaining that child.
A maximum `usize` depth policy no longer overflows when admitting a stream
selector: its derived depth-plus-one bound saturates. The existing policy also
bounds descendant storage count by `max_storage_depth`; this coupling is now
documented and a wide, shallow object reports `ObjectStorages`, separately
from actual nesting depth. This change preserves that configured ceiling.
Checked aggregate integer-overflow guards remain structural failures: a sum
beyond `u64::MAX` cannot be represented by the error's observed count and is
not a realizable retained-byte sum on this host.

The native probe builds the baseline and candidate from recorded source maps
against identical dependency sources. Synthetic fixtures cover a small object,
a 2 MiB unrelated opaque stream, and 256 unrelated sibling streams. Every
process checks exact no-op, forward and inverse artifacts before timing ingress,
edit creation, no-op commit, changed stream publication and patch replay.
The timing loop makes both owners and returned byte buffers opaque to the
optimizer before consuming their lengths. This matters after inlining: a
length-only consumer can eliminate an otherwise unobserved returned allocation.
The earlier `excluded-unobserved-output.json` capture is retained as an invalid
publication-cost comparison, and is not used for acceptance. The first short
fixed-iteration capture is likewise retained for diagnosis; several legs were
under one millisecond. Final sampling calibrates each route from the baseline
to target at least 150 ms per leg, with a minimum of 100 iterations.
Generated projects, binaries and raw captures remain outside the checkout.

## Validation and measured behavior

The OLE suite passes 315 tests, DOC 1,228, XLS 1,314, PPT 1,236 and unified
error classification 48: 4,141 distinct tests. Five existing external-fixture
or documentation-example tests remain ignored. The 23 XLS error tests were
also rerun after completing nested CFB conversion; they are not counted twice.
All affected-target clippy, the independently locked performance harness,
four glossary profiles, crate boundaries, formatting and 36 tooling tests
pass. The earlier combined host run's two stale XLS message assertions were
migrated; its failed status remains explicit in the receipt. DOC/PPT passed
that run, and subsequent complete DOC/XLS runs pass with the final callbacks.

Successful policy validation now checks fields directly, avoiding the earlier
seven-row error-metadata table on the stack. The replay wrapper is inline,
while exact base comparison and returned-buffer allocation/copy remain.
No structural or configured admission check is removed.

Ratios below are candidate/baseline. Each row reports three isolated ABBA
blocks on CPU 24; measurements are specific to these native object routes.

| Fixture | Route | Median | Range |
| --- | --- | ---: | ---: |
| Small | open | 1.006 | 0.991–1.013 |
| Small | begin_edit | 0.993 | 0.989–1.013 |
| Small | noop_commit | 0.988 | 0.984–1.002 |
| Small | replace_stream | 0.997 | 0.996–0.999 |
| Small | apply_patch | 0.876 | 0.858–0.911 |
| 2 MiB opaque | open | 1.002 | 1.001–1.012 |
| 2 MiB opaque | begin_edit | 1.010 | 0.969–1.015 |
| 2 MiB opaque | noop_commit | 1.001 | 0.995–1.003 |
| 2 MiB opaque | replace_stream | 0.999 | 0.999–1.000 |
| 2 MiB opaque | apply_patch | 1.025 | 1.010–1.028 |
| 256 siblings | open | 0.991 | 0.988–0.994 |
| 256 siblings | begin_edit | 1.001 | 0.998–1.003 |
| 256 siblings | noop_commit | 0.999 | 0.997–1.000 |
| 256 siblings | replace_stream | 0.982 | 0.977–1.018 |
| 256 siblings | apply_patch | 0.991 | 0.982–1.005 |
| 2 MiB opaque | `apply_patch`, longer follow-up | 0.992 | 0.986–1.015 |

The small-open candidate's earlier sustained cost is removed. Large replay's
full-pass range was positive, so a further three blocks used longer legs with
the same binaries. Those cross baseline; all six block ratios have an
unweighted median of 1.012 and range 0.986–1.028. This variation does not support
a stable large-replay speedup or slowdown. No favorable samples were selected
or discarded. Final direct heap profiles show 21.30 MB peak in both builds;
allocation calls are 17,170 before and 17,169 after (the typed conflict removes
one preflight error-string allocation). Both report 544 B runtime leakage.
Profiler wall time is excluded from native comparisons.

Complete samples, source/lock/binary hashes and commands are in
[the native measurements](results/0359e/object-measurements.json),
[the longer replay follow-up](results/0359e/large-replay-followup.json),
[the direct heap receipt](results/0359e/object-heap.json) and
[the validation receipt](results/0359e/object-validation.json). The paired
common source maps and shared dependency source map accompany them. Raw
profiles and command logs stay in the external directory recorded there;
no benchmark binaries are committed.

## Required continuation

This slice closes the configured **core object-capture** budgets. Raw directory
catalog, OLEDS, link, protection and other shared-helper policies/errors remain
required work. Their raw role requires explicit limits; it does not excuse
generic errors or hidden defaults. The shared object `Snapshot` publication,
owner/edit options and source-sharing semantics also remain required. This
record does not accept those paths as exceptions or close the legacy audit.

W4/W4b/W8 integration and the final corpus, performance, inventory and full
workspace gate remain required. Nothing is pushed by this record.
