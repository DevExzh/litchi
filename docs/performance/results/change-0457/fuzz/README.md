# Change 0457 bounded fuzz lanes

`run-fuzz.py` stages two independent cargo-fuzz-style crates under
`/tmp/litchi-goal-0457/fuzz-zip` and `fuzz-xml`. The ZIP lane uses the current
`parse_zip.rs` target with six retained ZIP seeds plus a small native ODP. The
XML lane uses the current common `scan_xml.rs` target with deterministic good,
bad, token, depth, BOM, and namespace seeds. Each staged manifest points to
the repository crate through an absolute path, so the temporary crates do not
depend on a copied checkout.

Every lane runs three serialized stages through the immutable parent
`../check.py`: lock generation, ASAN release build, and smoke execution. The
stage tags are unique (`fuzz-0457-{zip,xml}-{lock,build,smoke}`); existing
tags, staging directories, or evidence directories cause an immediate refusal.
A failed stage is not retried. Cargo uses the shared
`target/fuzz-asan` directory and the existing Cargo cache; the runner never
deletes either one.

The pinned ASAN flags match change 0456. Smoke runs use 1,000 iterations,
seed 457, a ten-second per-input timeout, a 1 MiB ZIP maximum, and a 64 KiB
XML maximum. Before lock generation, the runner archives the exact source,
adapted manifests, seed corpus, manifests, and SHA256 inventory. It archives
the generated lockfile before each build. After both smoke stages it copies
the binaries and parent receipts/logs and records hashes for them and every
staged corpus file.

The source and seed snapshots are sealed by `SHA256SUMS`. The runner also
compares the snapshots with the current Rust targets and manifests before it
creates `/tmp`; source drift aborts the run instead of silently testing an
older target.

Run only after the 0457 source epoch is frozen:

```sh
python3 -B docs/performance/results/change-0457/fuzz/run-fuzz.py
```

This directory currently contains preparation tooling and deterministic input
evidence. The command has not been run here.
