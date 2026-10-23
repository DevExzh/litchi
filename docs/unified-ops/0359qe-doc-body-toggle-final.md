# 0359qe — DOC body-toggle final integration qualification

- **Status:** deterministic profile complete; native ABBA pending; batch raw-byte
  fidelity is explicitly recorded for coordinator review.
- **Date:** 2026-10-23.
- **Integration:** `3cad59be707c7431bc1a8c115e1ff534a048dc57`.
- **Scope:** the existing private `litchi-doc` body transaction and the existing
  W5 probe. No public API or alternate editing family was added.

## Source assemblies

The candidate is the `3cad` tree as checked out. The baseline is an immutable
source overlay on the same tree: only the body-publication diffs from
`ae094d7d055c`, `6456aa3223`, and `160a00559b` were reversed in
`body_text.rs` and `tracked_revision/package.rs`. The current `PatchConflict`
versus composition `Conflict` taxonomy, current probe, lockfile, and every
other `3cad` file remain unchanged. The baseline is not described as pristine
`81323113be`.

The candidate implementation hashes are body
`614ce42c2dfe8ad985dc37d2c949602916bc155bd51149f2e843f24fdf378224` and
revision package
`0fd67fae678e870f6e7216f0ec5d64ea62a0da7db6df7f87297a965fd9a2da86`.
The baseline overlay hashes are body
`3e09f67abe5c40248959fe35ca90ea9c8ceb52609394267970b1307be09465d5` and
revision package
`613beaf162033c6153906ef2b1299d21e0ece58b4283a43400924c2503dbe40a`.
The exact two-file overlay diff has SHA-256
`10765febb64c4e099f5c3ef560366124dc1d28d540d138768ae0f7cf782da2ab`.

Both legs use probe SHA
`9e1cd649268ad5eefe4314b76493f61cd047bf6ec4114c7e7872b2182b1af689` and
the same release command:

```text
cargo build --offline --locked --release -p litchi-ops \
  --features doc,xls,rtf --example legacy_w5_source_probe
```

The build used toolchain 1.95.0, one Cargo job, incremental compilation off,
release debuginfo off, and the 20 GB memory-capped shim. Binaries and all
profiles remain outside the repository.

## Correctness and fidelity

The ordinary route (`doc/leaf_write`) produced identical raw output SHA
`6bb97058e47f747c29e568395eea01db1226aa5ab8962087e1d78a1b9b82ff11` on both
legs. The mixed-formatting test, refusal atomicity, inverse, and exact no-op
checks passed in the candidate's full DOC library run (1,210 passed, 2
ignored); the baseline overlay also compiled and its retained-render test
passed.

The 1k route preserves rendered output but not raw container bytes. The
candidate emits a compact 182,784-byte OLE file (`bd016a…`); the baseline emits
7,715,328 bytes (`4a5bcc…`) because it retains superseded CHPX/1Table history
from publishing every toggle. LibreOffice produced 18 pages for each file;
the 72-dpi rendered page hash list and extracted text are identical. This is
an explicit fidelity review item, not an unreported equal-byte claim.

## Performance

The selected ordinary route has four calls (expected output, two warmups, one
measured call). Callgrind inclusive Ir is 257,848,943 candidate versus
256,622,212 baseline (1.004785×). `Edit::commit` is 79,843,981 versus
73,325,557 (1.088818×), while the setter falls from 32,195,265 to 22,463,186
(0.697731×). The ordinary Massif peak is 7,910,440 versus 8,228,648 bytes
(0.961329×).

The full 1k commit-plus-write route is 241,104 µs candidate versus 1,319,299
µs baseline in three-round wall diagnostics. Callgrind inclusive Ir is
24,673,833,943 versus 60,148,185,975 (0.410228×); the setter is 0.408337×,
and `write_to` is 0.036455×. Massif peak is 8,769,392 versus 107,252,816
bytes (0.081765×). These measurements account for the historical 1.1–1.7 ms
per-1k residual and include the final `write_to`.

The host load was approximately 37 during native attempts, so the wall
numbers are diagnostic only. A quiet-host native ABBA block remains required;
no native regression or acceptance claim is made here.

Reduced measurements and exact source/binary/profile hashes are in
[`results/0359qe/qualification.json`](results/0359qe/qualification.json).
No large evidence or executable is committed.
