# 0362ul10b — OLE owner qualification validator closure repair

- Date: 2026-10-24.
- Parent packet: 0362ul10a.
- Status: source-only validator/oracle repair; native execution and acceptance
  remain pending coordinator review.

The independent post-run validator now revalidates the selected preparation
against the live source worktrees: Git HEAD/tree/status/diff identities, the
legacy/current CFB signatures, probe call sites and workload markers, every
external driver input, and the manifest pins.  Build receipts must carry the
exact Cargo argv, prepared-driver cwd, target path, source identities and
artifact handoff; ABBA receipts must carry the complete event-ID journal.

Oracle reports are bound to the selected run directory and to their exact
`oracle-{role}-{variant}` journal event.  The event's source identity, role,
variant, executable argv, cwd, environment, output paths and hashes are all
checked.  The oracle now emits a SHA-256 manifest of every output artifact on
stdout; the validator compares that invocation manifest with the independently
  rehashed files, so changing a source/edit/history artifact without changing
the actual oracle invocation is rejected.  OLE inventory proofs include the
root/storage metadata and every expected sibling; DOC covers all 256
paragraphs, XLS covers all 256 cells plus the 96 KiB opaque stream, and the
retained-history proof is required to cover 129 packages and 128 patches in
both directions.

ABBA measurements carry the exact clone, route and allocator event IDs.  The
validator checks their commands, source identities, paths, stdout rows and
ABBA order, requires all 268 referenced events, and admits only explicitly
shaped calibration events as additional journal entries.  Allocator routes
remain non-claimable and require ordered route rows, conservation, retained
history cardinality and a live event link; the Rust oracle keeps the retained
history tuple alive through the counter snapshot.

Source-only checks performed for this repair:

```text
python3 -m unittest tools.test_prepare_ole_owner_inputs tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run
python3 -m py_compile tools/native_qualification/*.py tools/test_ole_owner_paired_runner.py tools/test_prepare_ole_owner_inputs.py tools/test_validate_ole_owner_run.py
rustfmt --edition 2024 --check tools/perf-baseline/examples/ole_owner_oracle.rs
```

The updated oracle source SHA-256 is
`49df705ccd92e498efc51388cb3d00310de453a12e84d1919128c52ffeddee95`.
No Cargo build, native run, timing result, or performance acceptance claim is
made by this record.
