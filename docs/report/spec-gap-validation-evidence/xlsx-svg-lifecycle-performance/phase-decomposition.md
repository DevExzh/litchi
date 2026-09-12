# Exploratory SVG phase decomposition

This document defines a separate exploratory harness for the three deterministic
same-drawing attach fixtures. It is a diagnostic boundary plan, not an
acceptance lane and not a performance result. The sealed 69-lane baseline and
its raw receipts remain unchanged.

The harness is `harness/phase_main.rs`, built as the separate
`xlsx-svg-lifecycle-phase-profile` binary. It reuses the existing fixture
builder, public `Workbook::from_bytes`, `Workbook::edit`, semantic worksheet
selection, `PictureSelector`, borrowed `SvgInput`, `Edit::commit`, and the
existing allocator observer. It accepts only 16, 64, or 256 pictures, and the
runner uses three fresh processes, two warmups, and twenty measured samples.

Each measured sample uses the same package bytes and attach intent sequence as
`multi_picture_same_drawing_{16,64,256}`. The fixture is built before the
sample clock. The public operation is divided into these clocks:

| Phase | Timed public work | Objects intentionally retained at the boundary |
| --- | --- | --- |
| `open` | `Workbook::from_bytes` | source workbook |
| `stages` | `Workbook::edit` and every public `attach_svg` call | source workbook and edit |
| `commit` | `Edit::commit().into_workbook()` | source workbook and committed workbook |
| `firstsave` | committed workbook `to_plain_bytes` | source workbook, committed workbook, and first output bytes |
| `reopen_secondsave` | `Workbook::from_bytes(first_output.clone())` and reopened `to_plain_bytes` | source workbook, reopened workbook, and both output byte vectors |
| `validation` | complete multi-picture attach semantic predicate | source workbook, reopened workbook, and both output byte vectors |

The committed workbook is dropped after the `firstsave` clock and before the
reopen clock, outside both clocks. This matches the chained acceptance path,
where the temporary committed workbook is serialized and dropped before
reopen. The source workbook remains alive through validation, as it does in
the acceptance function's lexical scope. Reopened and source objects are
dropped after the validation snapshot. These retained lifetimes are recorded
in the receipt so phase live values are not mistaken for standalone operation
footprints.

Every phase resets only the process-local counter baseline, then records
elapsed nanoseconds, requested/direct/reallocation/deallocated bytes,
`live_before_bytes`, `live_after_bytes`, signed `live_delta_bytes`,
`retained_live_bytes_after`, peak live increase, and allocator validity. The
requested allocation value is the phase-local direct plus reallocation-new
traffic. The live equation is checked independently for every phase. The
receipt never adds phase allocations or subtracts values from unlike retained
live sets. Whole-process RSS remains in the `/usr/bin/time -v` sidecar.
After each warm-up and measured sample, the source workbook, reopened workbook,
and serialized byte buffers are dropped; the driver requires the process-local
live boundary and allocator validity to return to their pre-sample values.

Validation reuses the same complete semantic conditions as the acceptance
attach path: changed output, reopen byte identity, picture count and direct
embedded SVG owners, unchanged raster relationships, unique SVG relationship
IDs, exact SVG media bytes, graph closure, per-picture closure, and opaque
fragment preservation. A phase receipt is rejected if any semantic check,
allocator equation, typed input identity, process identity, or sidecar exit
check fails.

`run_phase_profile.sh` is a separate fail-closed source-capture runner. It
requires the freeze/API gates, rejects Rust flag suppression, requires fresh
external results and target directories, checks the existing 19 production
source pins with `committed_inputs.py`, captures before/after manifests with
the new phase binary and verifier sources, and removes only a target it
created. It runs `phase_summarize.py` and `phase_verify.py` after all raw
receipts are present. No exploratory run is valid from a dirty source tree or
an omitted phase-harness input.

The exploratory report is descriptive. It identifies the largest observed
elapsed and phase-local requested-allocation medians per fresh process for each
fixture size, subject to the three-process uncertainty and retained-live
boundaries. These are same-run absolute observations; they do not attribute
cost to an internal production function, compare unlike clocks to the sealed
baseline, claim a speedup or regression, or imply a scaling law. Production
optimization work remains gated on the semantic/source review and a separate
candidate pin.
