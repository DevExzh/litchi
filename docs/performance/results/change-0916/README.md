# 0916 evidence and reproduction

The [change record](../../0916-docx-section-commit-target-reuse.md) defines the
claim and its limits. `native-protocol.json` was frozen before measurement;
`manifest.json` binds its bytes, the runner, the unchanged 0914 validator, both
executables, and the source/build receipts. Never overwrite that packet when
regenerating an experiment.

## Replay retained evidence

```sh
python3 -B docs/performance/results/change-0916/native-compare.py replay
python3 -B docs/performance/results/change-0916/profile-replay.py
```

Native replay uses the committed compressed reports and receipts. It checks
their hashes, the frozen schedule and timestamp ordering, every 0914 report
gate, GNU time metrics, pooled statistics, and the paired bootstrap. It does
not require the external native reports or executable. Profile replay also
checks the external diagnostic captures and retained baseline executable.

`quality-receipt.json` indexes command, timestamp, exit and log evidence for
DOCX all-targets tests, doctests, formatting, Clippy, warning-denied rustdoc,
and the boundary checker. `capture-barrier.json` records coordinated process
idleness before timing. Each native child records host load before and after;
none of these records establishes whole-machine quiescence.

## Source and build

The retained baseline archive is
`/home/zhuhe/litchi-performance-0914/source-HEAD-code.tar` (171,120,640 bytes,
SHA-256 `7b77f1875adcdc3c822351b11dcf621f4ebff501fe501cab80fd4fb3f9a59571`).
Its source revision is `a4ca85e28cdea151176534b7f8dc912253a1ca1e`.
Regenerate it using the 0914 capture helper's `prepare` action, with
`LITCHI_PERF_0914_ROOT` set to a new external directory and
`LITCHI_PERF_0914_REVISION` set to that revision; see the
[0914 record](../../0914-docx-section-layout-baseline.md).

Extract the archive into a new external directory, enter its `litchi/` root,
and apply this packet's `candidate.patch` with `patch -p1`. The patch changes
only `src/section/layout.rs`, `src/source_backed.rs`, and
`tests/source_backed_sections.rs` under `crates/litchi-docx`. `source-audit.json`
compares all 9,357 archive files and their modes. The source receipt records
the patch hash and the before/after hashes of all changed files.

Build from that root with an external target directory:

```sh
taskset -c 14,15 env CARGO_TARGET_DIR=/new/external/target CARGO_TERM_COLOR=never \
  cargo build --locked --offline -j2 --release \
  --manifest-path tools/perf-baseline/Cargo.toml --bin litchi-perf-baseline
```

The recorded candidate executable is 69,973,984 bytes, SHA-256
`5857a0e06296634328e3d9f29fafbd3a60baf8faeb7bf148f99d8cd444a7297c`.
It remains outside Git at
`/home/zhuhe/litchi-performance-0916/candidate/target/release/litchi-perf-baseline`.
`receipts/build-receipt.json` binds the actual command, toolchain, effective
environment, timestamps, output, and executable hash. Rebuild paths and
toolchains can affect executable bytes; compare the recorded inputs as well.

Quality checks additionally need the retained full fixture archive (1,426
files), root workspace lock, and two SVG runtime fixtures listed with sizes,
hashes, and regeneration information in `quality-receipt.json`. The standalone
harness uses its own unchanged lock. For the original quality layout, extract
the fixture archive at the candidate root (it contains `fixtures/test-data/`),
restore the links in `source-tree-links.txt`, and copy the root lock and SVG
fixtures to their original relative locations. `quality-layout.json` records
the additional tracked checker, policy, ADR and performance documents and two
fixture staging copies. It pins the exact Git revision and hashes needed to
reproduce those quality-only extras; they are not release source changes.

The full prebuild file manifest is retained externally at
`/home/zhuhe/litchi-performance-0916/candidate/receipts/source-tree-files.sha256`:
1,311,966 bytes, SHA-256
`1265151d60ccdf64a4aa0dd6adc4126eaf41ca894c97214c5895fa4ab715e8be`.
`receipts/source-receipt.json` contains its exact regeneration command. The
manifest was reproduced unchanged after the build and quality checks. It is
not committed because it exceeds the evidence size limit.

## Fresh timing

In a separate Git checkout, copy only `native-compare.py` and
`native-protocol.json` into a new `docs/performance/results/change-0916-recapture/`
directory; leave its manifest, reports and receipts absent. Keep the existing
0914 validator at its relative location. Set `LITCHI_PERF_0916_ROOT` to a new
external directory for every action. Run `native-compare.py --help` for the
`freeze`, `run`, `reduce`, and `replay` actions. Freeze with both
`--source-receipt` and `--build-receipt`; the runner refuses to replace a
manifest or captured child. A rebuilt baseline with a different hash requires
a new premeasurement protocol, leaving this retained protocol unchanged.
Native controls use
CPU 12, eight alternating pairs, 30 warmups and 200 samples per child. Finish
all coordinated build, test and profiler work before the first child.

The large Callgrind files and their reproduction commands are indexed by
`profile-reduced.json`. They remain external; guest instructions are not native
cycles or an estimate of native hashing cost.
