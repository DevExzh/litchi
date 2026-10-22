# 0359o — XLS query-table fixture path repair

- Date: 2026-10-22.
- Scope: make the two query-table corpus tests independent of the process
  working directory.
- Status: focused validation passed; no performance claim.

The two tests for POI workbooks `45365.xls` and `57456.xls` previously passed
`../../test-data/...` directly to `std::fs::read`. Cargo runs these tests with
the workspace root as the current directory, so those paths escaped the
crate checkout and both tests failed with `fixture readable`. The fixtures
are tracked at `test-data/poi/test-data/spreadsheet/`; the test helper now
anchors that directory at `env!("CARGO_MANIFEST_DIR")` and accepts only the
fixture filename.

Validation used one capped Cargo process and the tracked fixtures:

```text
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-retag-20261022/target-xls-final
cargo test --locked -p litchi-xls --lib -- --nocapture
test result: ok. 1316 passed; 0 failed; 0 ignored; 0 measured; 0 filtered out
```

No `3rdparty/` path is referenced, and the earlier two failing attempts were
retained outside the checkout with the other validation logs.
