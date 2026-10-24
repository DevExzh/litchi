# XLS F3 repaired-overlay preparation

This is a compact source-only follow-up for the preserved 472b F3 build
failure. It makes no acceptance claim and records no Cargo, native probe,
calibration, timing, allocation, or ABBA execution.

- Candidate endpoint: `/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/xls-f3-candidate-472b-repair-472b`, HEAD `472b6705475b073e7fe08f6e11ee17fb1e29f838`, tree `ebc1cee05f124c0c63977aff07a23288c19e7d00`.
- Baseline endpoint: `/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/xls-f3-baseline-9b1-repair-472b`, HEAD `9b1f96d6dce1506dc9d0080f0696b42ba592c397`, tree `1d69a761f6d89107c9212b5e7f8d4917b1f4b1f0`.
- Corrected harness leaf: `bedd3f4d189b82de7d04d9afc2a333ab411fc88e`; driver-default leaf: `0ff64955ef19a0bf6319c18ecab85648ab563381`.
- Repaired sparse fixture: `test-data/poi/test-data/spreadsheet/54016.xls`, 984576 bytes, SHA-256 `2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a`; Apache POI path `test-data/spreadsheet/54016.xls`, Apache-2.0, with tracked `PROVENANCE.md` and `NOTICE`.
- Source-only checks passed: current `OleFile::open(reader, OleFileLimits)` signature; every F3 arity site has `OleFileLimits::default()`; candidate and baseline probe identities match harness ref; complete `validate_options` accepts the 3-route × 2-corpus matrix and production order `8d,ef,157,c7,a443`; fixture hashes match; active inputs contain no external checkout path.

The failed compile emitted eight total errors: seven one-argument F3
`OleFile::open` call sites (two corpus and five qualification calls) plus the
missing fixture. The uncompiled stale consumers in `xls_numeric.rs` and
`lib.rs` were intentionally preserved; they are outside this F3 compile
repair and are still represented by the frozen overlay identity.

Fresh build commands, target/TMP roots, source manifests, overlay hashes,
preserved failed log/target, and the future quiet-driver command are in:

`/home/zhuhe/code/litchi-worktrees/.coord/sheet-preparation-691941/sheet-native-readiness-472b-f3-repair.json` (SHA-256 `99a0c1f4968889e841c447caa6dfb53c76f65031175b65fd0e5dcffc37a5c7e8`).

Source-preparation manifest SHA-256:
`9968f2fcacf858cee6d69840bf15db174419743913d8d0028b0ddc7baff7d6da`.

Independent source-only review is requested from `inventory_drawingml_limits`.
The root coordinator controls the next Cargo grant; the historical corrected
packet and successful/failed build evidence remain untouched.
