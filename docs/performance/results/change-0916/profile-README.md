# 0916 — DOCX commit-publication target snapshot profile

This diagnostic profiles the suspected redundant target `Snapshot` parse in
`docx_source_backed_existing_section_layout_edit_save`. The measured
publication captures a fresh current-source snapshot and calls `Patch::apply`
to authorize the patch against it. For a changed section patch, that method
then constructs a fresh target snapshot
with `Snapshot::from_shared_xml`. `Edit::commit` has already parsed and retained
the same target snapshot in its `Commit`. The public `Patch::apply` contract is
broader and still needs to reconstruct a target for callers that hold only a
patch, so this packet makes no production change and does not claim that the
public method can be changed globally.

The retained release binary is
`/home/zhuhe/litchi-performance-0914/target/release/litchi-perf-baseline`,
69,940,744 bytes, SHA-256
`fe56e5192c1198e831583e870beb5e1ba4731fb2077faaffdfc46044bb0ed2db`.
The source archive is
`/home/zhuhe/litchi-performance-0914/source-HEAD-code.tar`, 171,120,640
bytes, SHA-256
`7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`, at
source revision `a4ca85e28cdea151176534b7f8dc912253a1ca1e`. The production
source member used for the line proof has SHA-256
`dcbd516f03a35047d0bf5b28e557c32882f5a4d1624d6dd08c265142d239aa21`.

The binary and source hashes were checked before the first capture. The small
manifest was written after capture and records that fact without backdating a
freeze receipt. All raw Callgrind outputs, raw harness reports, and clock
receipts remain outside Git under
`/home/zhuhe/litchi-performance-0916/profile`; the packet records their exact
sizes and hashes.

## Result

The target trigger was
`litchi_docx::section::layout::Patch::apply`, with four separated caller
levels. The selected part is part 8, whose caller chain reaches the measured
`Package::publish_section_layout_commit_to_stream` call:

| Selected owner | Callgrind Ir | Nested owner | Callgrind Ir |
| --- | ---: | --- | ---: |
| `Patch::apply` | 6,384,252 | `Snapshot::from_shared_xml` | 6,381,596 |
| `ensure_source_section_inventory_xml` | 2,144,436 | `parse_inventory` | 4,237,085 |

The matching whole-publication trigger was
`litchi_docx::source_backed::Package::publish_section_layout_commit_to_stream`.
Its measured part is part 7 at 240,572,558 Ir. Within that same selected
publication dump, `Patch::apply` is 6,394,087 Ir and its nested target snapshot
parse is 6,391,431 Ir. The target parse therefore accounts for 2.656758% of
the selected publication interval, and the complete `Patch::apply` interval
accounts for 2.657862%. In the target-trigger capture, the nested parse is
99.958398% of `Patch::apply`.

The target-trigger capture contains separate setup/control publications,
direct post-timer patch replay, and inverse-publication parts. The selected
part is the one whose caller chain includes both
`Patch::apply` and the measured `publish_section_layout_commit_to_stream`.
The whole-publication capture similarly selects the one part reached directly
from the measured scenario loop; setup and preservation controls are excluded.
The exact part IDs, caller records, excluded part lists, raw hashes, and
machine-readable values are in [`profile-reduced.json`](profile-reduced.json).

These are Callgrind guest instruction counts. `perf_event_paranoid=4` denied
hardware counters, and the dominant source-artifact SHA work executes as
software guest instructions. No native latency, native SHA cost, allocation
byte count, peak-live-memory value, or end-to-end speedup is inferred. The
capture used one zero-warmup sample per bounded diagnostic process, so it is
attribution evidence rather than a statistical estimate.

## Reproduction and verification

The two exact commands are stored in `profile-reduced.json`. They use the
retained binary, CPU 12, Valgrind Callgrind `Ir`, `--warmup 0`, and
`--samples 1`. Regenerate into a new external directory so the retained raw
artifacts are never overwritten. No candidate build is required.

Replay the retained caller proof and all external artifact hashes with:

```sh
python3 -B docs/performance/results/change-0916/profile-replay.py
```

The replay checks the binary, source archive member, corpus identities, raw
artifact custody, Callgrind part/trigger headers, and the selected caller
records. It does not rebuild the executable or modify the external profile
directory.
