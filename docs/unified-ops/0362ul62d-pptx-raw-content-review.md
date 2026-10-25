# 0362ul62d — PPTX raw-content aggregate and MCE replay correction

Status: source-ready correction to `0362ul62c`, based on the coordinator
replay integration. Native Cargo validation, the workspace gate, and ABBA
remain coordinator work because this leaf has no build slot.

## Corrections

- Package ingress now retains one bounded notes/media relationship index. A
  `PartPatch` publication carries that index forward with changed-part deltas:
  notes aggregate bytes update from changed notes slides, notes masters, theme
  users, and changed theme payloads; media reverse edges update from changed
  slide relationship graphs. Notes aggregate and resource-only media replay
  therefore validate the affected closure without rescanning unrelated package
  parts. The index is retained through facade, metadata, staged-family, and
  opened-candidate publication paths.
- Collaboration MCE offsets are translated back to original source positions
  before active typed validation. A PI-before-MCE regression now proves the
  successful active extension path, not only a refusal path.
- Collaboration parent and direct-child ownership is indexed once per raw
  scan. Presence/threading replay remains one MCE pass and one typed validation
  per active owner, without the former per-owner full-node filters.

## Coverage and integration

The existing exact/one-under notes aggregate and media resource-only tests now
exercise the retained index; collaboration coverage keeps partial-owner and
inactive/opaque URI cases. The coordinator's common physical admitted-byte
accounting remains separate: this leaf's index owns only semantic notes
aggregate and media affected-relationship closure.

No generated evidence, fixture, executable, or `3rdparty/` path is included.
