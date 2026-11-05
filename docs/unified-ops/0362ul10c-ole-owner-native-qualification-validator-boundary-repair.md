# 0362ul10c — OLE owner validator boundary repair

- Date: 2026-10-24.
- Parent packet: 0362ul10b.
- Status: source-only false-acceptance repair; native execution and acceptance
  remain pending coordinator review.

The validator and paired runner now close five remaining evidence gaps:

- each build executable is required at the exact
  `target_root/{role}-{variant}/release/examples/` path and exact expected
  filename, and build argv must use the recorded regular Cargo executable
  path from the toolchain receipt;
- allocator evidence must contain the six routes exactly once and in the
  manifest order;
- DOC and XLS proofs must enumerate source and edited indices exactly once in
  order `0..255`, rather than merely having the expected row count;
- OLE inventory proofs must contain the exact `Root Entry` storage identity
  (`entry_type=5`) on both sides, matching source/replacement metadata, while
  retaining all stream checks; and
- build event output paths are checked as an exact ordered path list.

Focused build-free mutations cover each false-acceptance case: a misplaced
executable, an argv/toolchain mismatch, a duplicated allocator route,
duplicated DOC/XLS indices, and altered Root Entry metadata.

Source-only checks performed:

```text
python3 -m unittest tools.test_prepare_ole_owner_inputs tools.test_ole_owner_paired_runner tools.test_validate_ole_owner_run
python3 -m py_compile tools/native_qualification/*.py tools/test_validate_ole_owner_run.py
rustfmt --edition 2024 --check tools/perf-baseline/examples/ole_owner_oracle.rs
```

No Cargo build, native run, timing result, or performance acceptance claim is
made by this record.
