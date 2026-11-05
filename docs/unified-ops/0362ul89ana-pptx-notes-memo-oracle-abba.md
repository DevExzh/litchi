# 0362ul89ana — PPTX notes-memo probe oracle and ABBA repair

This follow-up is based on `712fcabf6473461868317a9d0f751b3ab6c6bcf0` in a
new sparse worktree. It is harness-only and does not change the PPTX public
API, dependencies, lockfiles, default case catalog, or fixture data.

## Timing correction

`pptx_notes_memo_cold_capture` now parses the fixed fixture before starting
the operation clock. The timed call is only `Package::opened_presentation`,
which is the capture that admits and scans the notes-owning roots. Warm and
changed rows already prepare `Package::from_bytes` and `opened_presentation`
outside the clock and measure only the detached edit plus `commit`.

## Independent fidelity and raw oracle

Every retained sample serializes outside its clock, reopens the resulting
bytes with `Package::from_bytes`, and repeats the exact three-slide and
three-notes-slide semantic text oracle. A ZIP inventory independently checks:

- identical member identity sets and all untouched member payloads;
- exact expected changed payload set: none for cold, `notesSlide2.xml` for
  the warm notes edit, and `slide1.xml` for the changed-title edit;
- every relationship member's identity and raw payload (therefore every
  relationship ID/target) remains equal;
- all slide masters, layouts, notes masters, and theme payloads remain equal;
- actual slide/notes root counts, changed root indices, and derived untouched
  slide-root count.

The report records independent source/output payload digests, output reopen,
and the raw oracle fields. The changed-root and untouched counts are derived
from the inventory rather than asserted constants; no same-buffer hash or
non-empty-buffer assertion is presented as an independent byte verification.

## Dedicated ABBA validator

`tools/pptx_notes_memo_abba.py` validates the local probe envelope in the
fixed `A1, B1, B2, A2` order. It requires the pinned timing-scope strings,
matching fixture/sample/warmup identity, paired A/B source revisions and
binary identities with distinct control/candidate binary hashes, complete
elapsed statistics/sample-index alignment, deterministic output and fidelity
evidence for every retained sample, all semantic/raw fidelity gates, and either
native unavailable allocation samples
or measured allocator-region vectors. It rejects unknown fields at every
fixed report-object layer, fewer than four samples per round or fewer than
four rounds, and writes a create-new descriptive summary without inferring a
speedup. The direct four-report command is explicitly descriptive-only and
is not the process-boundary acceptance path:

```text
python3 tools/pptx_notes_memo_abba.py --descriptive-only --output SUMMARY.json A1.json B1.json B2.json A2.json
```

Acceptance requires the repeated fresh-process runner, whose manifest has at
least four explicit `A1, B1, B2, A2` rounds (16 child processes minimum):

```text
python3 tools/run_pptx_notes_memo_abba.py --manifest MANIFEST.json --receipt RECEIPT.json
```

Allocator evidence uses `--mode allocator` and requires the exact
`CountingSystemAllocator(std::alloc::System)` identity, counter revision
`serialized_region_peak_v3`, `operation_global_system_allocator` sample scope,
and every numeric field from the shared operation-scoped allocator sample.
Native reports must identify `Rust system allocator` and use the same exact
sample scope without numeric fields. The synthetic protocol tests cover a
passing four-round packet, short packets, false raw-oracle evidence,
source-pair identity mismatch, strict allocator identity/scope failures, and
allocator metric omission:

```text
python3 -m unittest tools/test_pptx_notes_memo_abba.py
```

No Cargo, Rust build, benchmark, Callgrind, native timing, or performance
acceptance was run for this source-only repair. The coordinator must rebuild
both exact endpoints and run the validator over four quiet ABBA legs before
using any timing result.
