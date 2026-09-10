# Neutral XLDM extraction freeze v3

This snapshot isolates the neutral XLDM extraction from the later Xldm140
identity projection. It is based on `906be17c2408e863bb26f2faf8f79361809597e6` and materialized at `/var/tmp/litchi-xldm-extraction-validation-20260910`.
Primary remains untouched and no worktree changes were staged or committed.

Included:

- new `litchi-xldm` neutral crate containing the bounded outer storage,
  metadata, native, generated, OLAP, compression, and validation modules;
- XLSX `package::xldm` public facade reexporting neutral types/modules while
  mapping the historical top-level `inspect`, `inspect_shared`, and `write`
  errors back to `litchi_xlsx::Error`;
- XLSX-only test fixture helper, preventing a reverse XLSX dependency;
- borrowed source-sharing APIs (`Storage::source_bytes` and
  `inspect_shared`);
- standalone lock updates for `tools/native-resave` and
  `tools/perf-baseline`.

Excluded:

- `identity.rs` and identity-specific OLAP validation;
- historical `docs/performance/results/**/Cargo.lock` evidence snapshots;
- primary repository edits, staging, and commits.

Focused gates in this worktree:

```
cargo test -p litchi-xldm --lib --no-fail-fast                         # 50 passed
cargo test -p litchi-xlsx --lib workbook::data_model::package::tests --no-fail-fast  # 27 passed
cargo test -p litchi-xlsx --test data_model_native --no-fail-fast       # 23 passed
cargo test -p litchi-xlsx --test xldm_native_storage --no-fail-fast     # 2 passed
cargo check --manifest-path tools/native-resave/Cargo.toml --features xlsx --locked
cargo check --manifest-path tools/perf-baseline/Cargo.toml --locked
```

Artifacts:

- `/var/tmp/litchi-xldm-extraction-v3-20260910.patch` — complete patch including new source files and tracked locks;
- `/var/tmp/litchi-xldm-extraction-v3-source-20260910.tar.gz` — deterministic archive of new files carried outside git diff;
- `/var/tmp/litchi-xldm-extraction-v3-source-sha256-20260910.txt` — SHA-256 list for those new files;
- `/var/tmp/litchi-xldm-standalone-locks-v1-20260910.patch` — standalone lock-only patch;
- `/var/tmp/litchi-xldm-standalone-locks-v1-20260910.md` — lock scope and commands;
- `/var/tmp/litchi-xldm-extraction-v3-status-20260910.txt` — exact status/name-status/stat snapshot.

Patch SHA-256: `683f64988c5e8bc272822b3f33267474272976f60a19bdff453d15b3ab36546b`
Archive SHA-256: `dd6c381f91120b82cd78bfc4922115a7bb8596450b021526651a0cc29d8cebfc`
Source-hash-list SHA-256: `8bfb544a810e2885ec00aec679fcaa1f001b76a65e3c18f72ceae3d8d0222c4d`
