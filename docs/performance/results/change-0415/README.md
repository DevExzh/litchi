# Change 0415 evidence

This is a ZIP64 capability and resource observation bundle, not a speedup
claim. Control is `916f60c38`; production candidate is `caf0d9394`.

## Retained observations

- `protocol.json`, `environment.json` and `identities.json` bind scope,
  toolchain, source files and executables. `capture.json` records exact guard
  argv, clean worktree revisions, executable hashes and process order.
- `guards/` retains 300 samples and 30 warmups per row in A1/B1/B2/A2 order.
  `followup/` retains the triggered 1,000-sample, 100-warmup repetition using
  the same binaries. Every process has a `/usr/bin/time -v` report.
- `observations/` separates file-output writing, Rust full readback and Python
  full readback for borrowed/owned writers at 64 MiB, 256 MiB and 4 GiB + 1.
  These use zero input. Files are caller-selected scratch on `/tmp` tmpfs;
  output storage and page-cache memory are outside process RSS.
- `large_output.rs` is a separate physical-output check using a repeated
  pseudo-random block. It makes the compressed Deflate payload itself cross
  the ZIP32 size boundary. Its process RSS includes writing and subsequent
  bounded Rust readback. Its timings are not ABBA guard observations.
  The retained `large-output-probe.zst` binds the exact auxiliary executable to
  its recorded hash. Writer timing starts after file creation and input-block
  generation; process RSS includes those setup steps.
- `interop.py` independently creates a four-member generic OPC package using
  Python `zipfile`/zlib. `large.bin` is exactly 4 GiB. All four members are
  drained with a 64 KiB buffer, retaining count, CRC and SHA-256. Its global
  tail is ordinary ZIP32; its large member uses ZIP64 size metadata. The Rust
  readback records this distinction explicitly.
- `profile/` retains the candidate's owned zero-stream diagnostic: four
  4-GiB-plus-one operations into an observed `io::sink`, frame-pointer sampling
  at 199 Hz on CPU 2, raw perf data, symbolized stacks and a flame graph. These
  profiled timings are excluded from the unprofiled results.
- `checks/` contains build, test, gate and fuzz logs plus exact lockfiles.
  Preliminary failed checks are retained alongside their passing replacements.

The ordinary probes retain a fixed 64 KiB input block. Borrowed transport
copies in its existing 16 KiB window; consuming entries receive 64 KiB chunks.
Sink byte/write-count/CRC observations are inside the guard timer. Input block
generation and output-oracle comparisons are outside it. The repeated random
block is generated from the xorshift seed in `probe.rs`; it is longer than the
Deflate history window. No cold, native-device, remote or scaling claim follows.

## Replay retained results

From this directory, `sha256sum --check SHA256SUMS` verifies every inventoried
artifact. From the repository root:

```sh
python3 docs/performance/results/change-0415/summarize.py \
  --root docs/performance/results/change-0415 --output /tmp/0415-guards.json
cmp /tmp/0415-guards.json docs/performance/results/change-0415/guard-summary.json
python3 docs/performance/results/change-0415/summarize.py \
  --root docs/performance/results/change-0415/followup \
  --samples 1000 --warmups 100 --output /tmp/0415-followup.json
cmp /tmp/0415-followup.json docs/performance/results/change-0415/followup/guard-summary.json
python3 docs/performance/results/change-0415/verify.py
```

The guard verifier checks every raw sample vector, exact medians and
nearest-rank p95/p99, output byte/write/CRC equality and mandatory process RSS.
It distinguishes candidate regressions from within-revision drift and retains
every positive delta above 5%. `verify.py` additionally checks capture order,
binary/revision bindings, cross-implementation readback and capability results.

## Fresh capture

Use clean detached worktrees at the control and candidate revisions. Create a
standalone Cargo package outside both trees with an empty `[workspace]`,
edition 2021, and these dependencies:

```toml
[dependencies]
soapberry-zip = { path = "/absolute/selected-tree/crates/soapberry-zip" }
crc32fast = "1"
serde_json = "1"
```

Point a `[[bin]]` named `zip64-stream-probe` at the absolute retained `probe.rs`
path, and a second named `zip64-large-output-probe` at `large_output.rs`.
Copy `checks/probe-Cargo.lock` into that package as `Cargo.lock`. Build each
selected tree in sequence with:

```sh
env CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=1 \
  RUSTFLAGS='-C force-frame-pointers=yes -C force-unwind-tables=yes' \
  cargo +1.98.1 build --release --locked --manifest-path /absolute/probe/Cargo.toml
```

Copy executables to distinct paths before changing the dependency path and
building the other revision. Adapt `capture.py`'s `--control`, `--candidate`,
`--control-tree`, `--candidate-tree` and `--output` arguments to these paths.
Use the defaults for the original protocol and `--samples 1000 --warmups 100`
for the follow-up. `observe.py` takes `--probe`, `--scratch-prefix` and `--output`.
The physical-output probe takes an output path followed by `4294967297`.

Generate the independent producer with `python3 interop.py generate PATH`.
`probe verify PATH` drains every member through the indexed Rust reader;
`python3 interop.py verify PATH` independently verifies every member. Retained
time reports and capture manifests contain exact original argv. Serialize all
measurements against builds, tests and analysis workloads.

The profile was recorded with `perf record -F 199 --call-graph fp`, followed by
`taskset -c 2 CANDIDATE guard owned zeros 4294967297 3 1`. Postprocess with
`perf report --stdio --no-children --no-inline` and `flamegraph --perfdata DATA
--no-inline`. Use `DEBUGINFOD_URLS=''`. Symbolized `stacks.folded`, perf script
and the SVG remain readable after task executables have been removed.

## Correctness gates

Use Rust/Cargo 1.98.1 and the retained workspace lockfile in a disposable
worktree. The integrated all-feature/all-target ZIP, OPC and ODF common run
passes 1,222 tests; 43 doctests pass and two are ignored. Both extra tests that
are ignored in the normal target run were executed successfully:

```sh
cargo +1.98.1 test --release --locked -p soapberry-zip \
  --test streaming_zip64_deflate -- --ignored --nocapture
env LITCHI_0415_PYTHON_ZIP=PATH cargo +1.98.1 test --locked -p litchi-opc \
  --test external_zip64_source -- --ignored --nocapture
```

Warning-denied rustdoc and scoped Clippy pass. The unexempted Clippy command
still encounters pre-existing lints; the passing command uses only the five
previously documented command-level allowances: `chunks_exact_to_as_chunks`,
`err_expect`, `bool_assert_comparison`, `large_enum_variant`, and
`redundant_pattern_matching`. New findings were fixed. Workspace formatting
still reports only unchanged `litchi-docx/tests/glossary_authoring.rs`; all
changed Rust files pass formatting. Crate boundaries, CRUD coverage, the eight
strict registered claims and report classification pass.

The existing ZIP and OPC fuzz targets each run 1,000 iterations from three
deterministic seeds, using seed 415 and AddressSanitizer plus LLVM sanitizer
coverage. The Cargo fuzz driver is unavailable, so the existing binaries are
built directly with `RUSTC_BOOTSTRAP=1` on the recorded 1.98.1 toolchain. Exact
flags, target, seed hashes and lockfiles are retained in `fuzz.json` and `checks`.
This is a bounded fuzz smoke run, not exhaustive malformed-input coverage.
