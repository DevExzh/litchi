# 0362ul69c — Formula writer depth and append-transaction return

Date: 2026-10-25. Parent: 0362ul69b. Status: source repair complete; the
coordinator still owns Cargo, rustdoc and performance validation.

This return closes the remaining writer review points:

- The exact depth walk now has single- and double-rule regression coverage. A
  ruled wrapper contributes the same two encoded levels in both implementations
  (`write_ruled` enters the template and its slot LINE, and `depth_ruled`
  admits those same levels); the enclosing root LINE makes the single-rule
  case actual depth three and a double-rule case actual depth four. Review
  confirmed that this was already the runtime model, so this return does not
  claim a ruled-depth correction. The walk uses the same `admit_depth` helper
  as `NodeWriter`.
- `write_nodes_into` checkpoints the caller's original length and truncates
  any records emitted before a later encoding error. Unsupported characters and
  NUL-containing font names are also rejected during preflight, before output
  begins, while the checkpoint protects future low-level refusal paths.
  Existing caller prefixes are covered by focused failure tests.
- Character-width and font-name validation are folded into the existing
  preflight string/work scan, so there is no separate full AST pass for those
  leaves and no change to successful MTEF bytes.

No public rollback or deprecated alias was added. No MTEF semantic value,
wire spelling or selected limit profile changed. Source-only checks were
`rustfmt --edition 2024`, `git diff --check`, and repository size/path guards;
no Cargo/native/rustdoc command, binary, large evidence artifact or upstream
checkout path is part of this return.

## Coordinator validation and hydration plan

This packet remains source-only until the coordinator grants a quiet Cargo
slot. The eventual validation must hydrate the exact reviewed root lock, not
resolve dependencies again:

| lock | bytes | SHA-256 | use |
| --- | ---: | --- | --- |
| root `Cargo.lock` | 139,606 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` | workspace Formula and consumer checks |
| `tools/native-resave/Cargo.lock` | 56,975 | `51c25d3d6f36e20b50f0b19e39b1aa9c7b49777154122ecc95d3acd05c041378` | original native-resave tool lock |
| `tools/perf-baseline/Cargo.lock` | 57,227 | `eac86284f67136aa856a8dcb6f383178b89ca92b26434a3de1d7408d68f2690c` | original perf-baseline tool lock |
| `tools/perf-execution/Cargo.lock` | 26,543 | `2e657990eedc21de6c4ce31ac48be2955f1b75f05d1697c0054c3de5ca914340` | original perf-execution tool lock |

The canonical root lock is retained outside the checkout at
`(deleted-coord-artifact)/4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1.Cargo.lock`.
Hydration copies that file into the disposable validation source and verifies
the size and digest before any build. The tracked tool lock files above are
also copied or used in place, with their recorded digests checked. No
`cargo generate-lockfile` is permitted for this packet, and no lock is
silently replaced by a newly resolved one.

Once the slot is available, run the complete Formula suite with the root lock,
then the direct feature consumers (`litchi-doc --features formula`,
`litchi-ops --features formula`, and the `litchi` facade with
`--features formula`; the all-feature consumer pass covers the formula-enabled
PPT path),
including all targets and doctests where the gate selects them. Run the same
matrix under strict Clippy (`--all-targets --all-features -- -D warnings`) and
the repository API/glossary gate. The original native-resave, perf-baseline,
and perf-execution locks remain available for the coordinator's tool checks;
those checks are consumer/build validation, not a substitute for the Formula
suite.

The required Formula boundary tests are the raw input/depth/work/item/string
and output-limit edges, the split OMML Text/CDATA cumulative-string case, the
MTEF mapped and embellished-string cases, exact single/double ruled depth,
and the prefixed-buffer rollback cases for unsupported characters and NUL
font names. After correctness and Clippy pass, the coordinator may schedule
the existing ABBA lane. Until those receipts exist, this record makes no
build, consumer, Clippy, fidelity, or performance claim.

The coordinator's command matrix is:

```text
cargo test --offline --locked -p litchi-formula --all-features --all-targets --no-fail-fast
cargo test --offline --locked -p litchi-doc --features formula --all-targets --no-fail-fast
cargo test --offline --locked -p litchi-ops --features formula --all-targets --no-fail-fast
cargo test --offline --locked -p litchi --features formula --all-targets --no-fail-fast
cargo clippy --offline --locked -p litchi-formula -p litchi-doc -p litchi-ops -p litchi --all-targets --all-features -- -D warnings
cargo check --offline --locked --manifest-path tools/native-resave/Cargo.toml --all-targets --all-features
cargo check --offline --locked --manifest-path tools/perf-baseline/Cargo.toml --all-targets --all-features
cargo check --offline --locked --manifest-path tools/perf-execution/Cargo.toml --all-targets --all-features
```

These are planned commands only; none was run while the coordinator's four
heavy-build slots were occupied.
