# 0495 supplemental existing fuzz smoke

All four existing targets passed 1,000 executions each in the final capture:
DOCX `parse_docx`, `tail_append`, `source_backed_tail_append_stream`, and OPC
`parse_opc`. Each exited zero with no crash artifact or AddressSanitizer error.
This is a bounded smoke, not an exhaustive safety result or direct fuzz coverage
of every new managed-edit API. Focused production tests cover that API.

`final-runs/*/started.json` and `terminal.json` retain exact argv, selected
environment, binary hashes, initial seed hashes, timestamps, exit status, and
raw-output hashes. Final corpus mutations and raw logs are retained alongside.
The worker's preliminary 4,000 executions are separately preserved; the final
4,000 repeated them to add exact command and binary custody. No performance
claim follows from either run.

The binaries used Rust 1.98.1 with AddressSanitizer and sanitizer coverage via
`RUSTC_BOOTSTRAP=1`; see `plan.json` for compiler flags and tool versions.
The generic `llvm-symbolizer` command was unavailable. Build logs and separate
standalone fuzz workspace locks are retained under `logs/` and `fuzz-sources/`.

To reproduce, create a disposable checkout at the revision in `plan.json`,
apply `candidate-source.patch`, and copy `workspace-Cargo.lock` to its root.
Copy each retained `fuzz-sources/<crate>/Cargo.lock` into that crate's fuzz
workspace. Use the environment in `plan.json` with new disposable target and
TMPDIR paths, then build both fuzz manifests with:

```sh
cargo build --release --locked --offline --manifest-path crates/litchi-docx/fuzz/Cargo.toml --target x86_64-unknown-linux-gnu --bins
cargo build --release --locked --offline --manifest-path crates/litchi-opc/fuzz/Cargo.toml --target x86_64-unknown-linux-gnu --bins
```

Offline builds require the recorded dependencies already available locally.
Recreate each initial corpus using the names and hashes in its `started.json`,
taking generic seeds from `inputs/docx` and streaming seeds from
`inputs/source-backed`. Execute the retained argv with new binary, corpus, and
artifact-directory paths and its recorded runtime environment. Seeded fuzzing
can still vary with toolchain, platform, and sanitizer behavior.

`candidate-source-check.json` records all 31 candidate Rust files, checked
byte-for-byte against the primary workspace before the final runs. The base
revision and full patch preserve their reconstruction. `cleanup.json` confirms
the isolated checkout and target were removed after a process-reference audit.
The main 0495 performance seal was not changed. `inventory.json` authenticates
this supplemental bundle's files except itself.
