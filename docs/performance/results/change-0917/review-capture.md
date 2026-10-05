# 0917 capture-harness final review

Status: read-only follow-up review, 2026-10-05.  The frozen protocol and
qualification are bound to source revision
`bb19ca74bd52044fbd2c1c7cc9875e975bb334c0`.  I did not build, run a measured
capture, modify production code, or commit.

The design review records that the earlier probe gaps are closed: the exact
no-op control now checks `changed() == false`, `patch.is_noop()`, zero patch
changes, exact source bytes and allocation sharing; same-value and
operation-limit controls compare exact bytes and patch counts with independent
controls while retaining the typed refusal and capacity checks.

## Qualification evidence

The external functional qualification receipt reports 48/48 passed children:
24 native and 24 allocator, covering two fixtures, four edit counts, and
three retention routes.  The packet's updated validator independently accepted
all 48 retained qualification reports (warmups 0, one report sample per
child), including the external source-tree fixture paths.

Temporary copies of qualification reports were then corrupted four ways:

* report schema string;
* fixed-32 stage vector length;
* source SHA-256;
* allocator retained-byte value.

All four were rejected by `validate_report` with the expected schema, stage,
source-identity, and retained-capacity errors.  The existing retained negative
control independently records rejection of a corrupted stage vector.

## Script and custody review

The reviewed script changes close the earlier static issues:

* `sample_stages` now requires each sample's own stage vector;
* protocol validation checks the frozen source revision, schedule, bootstrap
  unit/statistic/interval, and review flags;
* receipt and report provenance are compared with the frozen manifest;
* replay requires exact expected receipt stems and exact reduced-manifest
  provenance;
* `read_receipt` now uses `packet_descriptor`, so packet-local source/build
  receipts are relocatable.

In a temporary copied checkout hierarchy, both copied source and build
receipts produced relative packet descriptors and resolved to the copied
files.  `verify_descriptor` rejected a mutation of the copied source receipt,
and retained-file resolution stayed inside the copied packet.  Frozen packet
files under `change-0917` and the qualification packet contain no file at or
above 1 MiB; large source archives, binaries, and any future raw evidence stay
external.

## Remaining replay constraint

The current `protocol()` implementation still runs `git rev-parse HEAD` and
requires it to equal the pinned protocol `source_head` (see
[`capture.py`](capture.py:168)).  That check passes in the current pinned
checkout and is useful before freeze/capture, but the planned adoption commit
will change the checkout HEAD by design.  If final replay is run after that
commit, `manifest()`/`replay.py` will fail before validating the retained
packet.  Root must either run the final replay while the pinned checkout is
still present or remove/scope this live-HEAD check for replay and rely on the
immutable source/build/probe receipts.  This is a custody sequencing
constraint; it does not invalidate the 48/48 functional qualification.

## Disposition

**Cleared to freeze and start the native/allocator capture from the pinned
checkout.**  The functional qualification, schema corruption controls,
source/build receipt relocation controls, schedule, pairing, bootstrap, and
packet-size checks pass.  Final post-adoption replay requires the replay
constraint above to be honored.  I am idle pending any further request.
