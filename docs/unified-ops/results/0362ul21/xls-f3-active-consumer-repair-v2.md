# XLS F3 active-consumer repair v2

This follow-up preserves the earlier failed build and repair packets and
records a fresh source-only candidate endpoint after review found that the
old repair overlay changed active consumers outside the isolated F3 probe.

- Candidate: `/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/xls-f3-candidate-472b-repair-v2`
  at `472b6705475b073e7fe08f6e11ee17fb1e29f838`, tree
  `ebc1cee05f124c0c63977aff07a23288c19e7d00`; status and diff are clean.
- The endpoint uses the exact tracked 472b `perf-baseline/src/lib.rs` and
  `xls_numeric.rs`. The implicit `perf-baseline` library target and
  `mod xls_numeric` are therefore covered; all five xls_numeric opens and all
  seven F3 opens use `litchi_cfb::OleFileLimits::default()`.
- Fixture: tracked Apache POI `test-data/poi/test-data/spreadsheet/54016.xls`,
  984576 bytes, SHA-256
  `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a`, with
  the tracked provenance and NOTICE files. No `3rdparty/` reference is used.
- Driver: external branch `uo2/xls-f3-driver-v2`, commit
  `7280d5ad8724012dbd0c89bff0a000578d285435`; it pins the harness and
  candidate source to 472b and retains production order 8d, ef, 157, c7, a443.
- Source identity and options-only `validate_options` passed for the complete
  3-route × 2-corpus matrix. No Cargo, native, timing, allocation, or
  qualification driver ran; the options check used existing distinct baseline
  paths only to satisfy path existence.

Packets:

- `/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/xls-f3-repair-source-preparation-472b-v2.json`
  SHA-256 `6612b3408991a234797ec9ae20d89b864ab8971aa3cb2c5e0e8e79ca6b1c739f`.
- `/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/sheet-native-readiness-472b-f3-repair-v2.json`
  SHA-256 `2e57e9e20caf9e2994b0abab92c62f15c1d212d90fcb9a35913cc27b98e09aff`.

The historical baseline endpoint, successful baseline binary, prior receipt,
failure log, old candidate endpoint, and prior packet hashes remain unchanged.
Root must complete the baseline-specific 472b probe rebind and grant a quiet
slot before any qualification run; this record makes no acceptance claim.
