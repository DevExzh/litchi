# 0362dvfa — DrawingML diagram raw review return

- Status: historical return for the `e980` repair; its focused source checks
  passed, but their root-lock provenance was not captured and they are not
  final acceptance evidence. The superseding `9786`/52ed packet is recorded in
  `0362dvfb`. ABBA measurement, native builds and the final coordinator gate
  remain intentionally unclaimed.
- This is a docs-only receipt. It makes no source change after
  `e98076db975242b74e0653f06d6094a1e63bb6cb`.
- Base: `2f7db7ee7fa01893e9099217fc4243687d196f63`.
- Source tree at validation: commit
  `e98076db975242b74e0653f06d6094a1e63bb6cb`, tree
  `c4b83818c24063c562403e3e461eb525b8e04001`.
- `tools/perf-baseline/unified-smartart/Cargo.lock` SHA-256:
  `7e80d6b7256ae348253671373934172c9139cece5b9c7ff14a20d20367e938d9`.

## Superseded qualification status

The source checked here was `e98076db975242b74e0653f06d6094a1e63bb6cb` and
the historical focused run reported 245 `litchi-drawingml` tests, consumer
library checks and a standalone harness check. The commands used an external
target, but no root-lock pre/post receipt or persisted command logs were made;
the lock was therefore unknown/old (`6d7895e6056c931fc19780ff4d0a031230e2ac79b9b30a9060bc8b6fad6747da`
is only the later preserved
diagnostic lock hash, not evidence that this historical run used it). These
results remain diagnostic context for the review and must not be treated as
52ed acceptance. The final source correction at
`9786dcf45056009768f66579eb1e90faac8145c3` passed the 52ed-locked focused
packet, with logs and exact lock provenance, in
`0362dvfb-drawingml-diagram-raw-review-return.md`.

## Review returns and dispositions

The first `025ae` return was **NOT READY** for five issues. The `e980` source
repair dispositions are:

1. `edit::read_drawing` now maintains a global shape/node count across
   siblings, rather than charging only the open-element stack. Empty shapes
   are included, with a sibling-budget regression.
2. Drawing parsing tracks the root, stack and depth to reject EOF with an
   unclosed root or nested element. Nested empty `dsp:drawing` roots are also
   rejected.
3. `Parsed::read` now requires one closed data-model root and rejects duplicate
   roots, outside-root text/markup and trailing non-whitespace data.
4. `authoring::Spec::from_parts` and `recognize` validate direct inputs against
   the selected `Limits` and regenerate through bounded `data_xml` and
   `layout_xml`; they no longer call unbounded raw generators.
5. The public PPTX `shape::diagram::graphic_frame` generator takes the
   canonical `&diagram::Limits`; all live consumers and tests were migrated.

The second `8f49re` return was **NOT READY** for six remaining issues. The
same source repair closes them as follows:

1. `read_drawing::Event::Empty` rejects a nested empty `dsp:drawing` root.
2. `DiagramDefinition` tracks root-seen/root-closed state and rejects nested
   or duplicate roots, unterminated input and non-whitespace trailing markup.
3. `SmartArt::parse_data_xml` requires the DGM data-model root and namespace,
   rejects wrong roots, plaintext, empty roots and trailing/unterminated input,
   and validates the complete point graph.
4. Generic SmartArt data and cached-drawing generators now traverse every
   nested node, emit the corresponding presentation points and parent
   connections, and cache every descendant shape. A nested round-trip test
   proves that descendants and references are retained rather than flattened
   into dangling links or omitted drawings.
5. Resource-scaled `Spec::parse_outline` and `Spec::outline` now require
   `&Limits`, preflight input/output bounds and reserve only checked capacity.
6. Behavior regressions cover the global shape budget, nested empty roots,
   definition/data-model grammar, nested fidelity and custom `graphic_frame`.
   The canonical API list includes the authoring and graphic-frame surfaces;
   pending ABBA/native validation is stated truthfully below.

## Focused validation packet

All commands below used the repository memory-capped cargo shim, one build job,
offline dependencies, and the external target
`/tmp/uo2-drawingml-diagram-raw-target.f3Kak0`. The temporary directory was
`/tmp/uo2-drawingml-diagram-raw-tmp.5GD8hv`.

| Command | Result |
| --- | --- |
| `PATH=(deleted-coord-artifact):$PATH CARGO_TARGET_DIR=/tmp/uo2-drawingml-diagram-raw-target.f3Kak0 TMPDIR=/tmp/uo2-drawingml-diagram-raw-tmp.5GD8hv CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo test -p litchi-drawingml --all-features --offline --locked --lib --tests --no-fail-fast --jobs 1` | Passed: 245 `litchi-drawingml` unit tests and all integration suites. |
| `PATH=(deleted-coord-artifact):$PATH CARGO_TARGET_DIR=/tmp/uo2-drawingml-diagram-raw-target.f3Kak0 TMPDIR=/tmp/uo2-drawingml-diagram-raw-tmp.5GD8hv CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo check -p litchi-docx -p litchi-pptx -p litchi-ops --all-features --offline --locked --lib --jobs 1` | Passed. |
| `PATH=(deleted-coord-artifact):$PATH CARGO_TARGET_DIR=/tmp/uo2-drawingml-diagram-raw-target.f3Kak0 TMPDIR=/tmp/uo2-drawingml-diagram-raw-tmp.5GD8hv CARGO_BUILD_JOBS=1 CARGO_INCREMENTAL=0 CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0 RUST_TEST_THREADS=1 cargo check --manifest-path tools/perf-baseline/unified-smartart/Cargo.toml --all-features --offline --locked --jobs 1` | Passed after the tracked harness lockfile was refreshed for the current workspace dependency set. |
| `PATH=(deleted-coord-artifact):$PATH CARGO_TARGET_DIR=/tmp/uo2-drawingml-diagram-raw-target.f3Kak0 TMPDIR=/tmp/uo2-drawingml-diagram-raw-tmp.5GD8hv CARGO_BUILD_JOBS=1 cargo fmt --all -- --check` | Passed. |
| `git diff --check` | Passed before the source commit. |

The external target measured `1,199,565,942` bytes and the temporary directory
measured `0` bytes at the final source check, both below the 20 GB lane cap.
The focused commands wrote output to interactive tool streams rather than
files; therefore there are no persisted per-command log files and no log
SHA-256 values to report. This is deliberately recorded as unavailable rather
than inventing log hashes. The source, tree and lock hashes above are the
reproducible identifiers for the checked packet.

## Failed or non-evidence attempts

- An early pre-`e980` compile of the repair had ordinary source errors (a
  `usize`/`Option` mismatch, `Namespace::as_ref` misuse, an unreachable match
  arm and an unused assignment). Those were fixed before the `e980` source
  commit and that failed compile was not used as qualification evidence.
- The first standalone-harness check was attempted with `--locked` before its
  tracked lockfile contained the current workspace package entries. Cargo
  correctly refused to update the stale lockfile; the lockfile was then
  refreshed in the source commit and the exact locked command above passed.
- A temporary multi-filter Cargo invocation returned Cargo usage rather than
  running tests; it contributes no test result.
- Before the coordinator prohibited a repository-wide capture, an accidental
  `python3 tools/check_api_glossary.py` was started in this leaf and then
  terminated. It spawned the observed Cargo/rustdoc process `3109410`, which
  was documenting `litchi-ooxml-common` because that broad glossary traversal
  had reached that workspace crate. It used the default workspace target and
  parallelism, not the focused external target above, and produced no accepted
  result or evidence. No further Cargo command was started for this receipt.

At the final process check, no owned DrawingML test, rustdoc or glossary
process remained. A separate chart-lane Cargo test was still running in
`/home/zhuhe/code/litchi-worktrees/uo2-drawingml-chart-raw` with its own
`(deleted-coord-artifact)`; it was unrelated and was left untouched.

No fixtures, binaries, large evidence artifacts or `3rdparty/` references are
introduced by this return. The receipt does not claim ABBA performance or
native/full-workspace qualification; those remain coordinator work.
